#!/usr/bin/env python3
"""
Enrichit agences_immobilieres.xlsx avec Firecrawl (/v2/map puis /v2/scrape).

- Plusieurs cles API en rotation (round-robin). Une cle a court de credits (402) ou
  invalide (401) est retiree ; une cle limitee (429) est mise en pause puis reprise.
- Aucune limite sur le nombre d'agences, d'URLs mappees (limit=100000) ni de pages.
- Reprise automatique : chaque agence terminee est ecrite dans state/results.jsonl.
- Resultat : un seul xlsx, une ligne par agence, identifiant = nom de l'agence.

Codes de sortie : 0 = tout est traite | 10 = budget temps atteint, il reste du travail
                  3 = toutes les cles sont mortes | 4 = aucun progres | 2 = config invalide
"""
from __future__ import annotations

import argparse
import collections
import itertools
import json
import os
import queue
import re
import signal
import sys
import threading
import time
import unicodedata
from datetime import datetime, timezone
from urllib.parse import urldefrag, urlparse

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Font, PatternFill

# --------------------------------------------------------------------------- config


def env_int(name: str, default: int) -> int:
    try:
        return int(os.environ.get(name, "") or default)
    except ValueError:
        return default


def env_bool(name: str, default: bool = False) -> bool:
    v = (os.environ.get(name) or "").strip().lower()
    if not v:
        return default
    return v in ("1", "true", "yes", "oui", "on")


BASE_URL = os.environ.get("FIRECRAWL_BASE_URL", "https://api.firecrawl.dev/v2").rstrip("/")
INPUT_XLSX = os.environ.get("INPUT_XLSX", "agences_immobilieres.xlsx")
OUTPUT_XLSX = os.environ.get("OUTPUT_XLSX", "output/agences_firecrawl.xlsx")
STATE_DIR = os.environ.get("STATE_DIR", "state")
TIME_BUDGET_MIN = env_int("TIME_BUDGET_MIN", 330)  # 0 = illimite
PER_KEY_CONCURRENCY = env_int("PER_KEY_CONCURRENCY", 2)  # = navigateurs simultanes du plan (Free 2, Hobby 5, Standard 25)
RATE_PER_MIN = env_int("RATE_PER_MIN", 10)  # requetes/min par cle ET par endpoint (Free 10, Hobby 100, Standard 500) ; 0 = off
MAX_PAGES_PER_SITE = env_int("MAX_PAGES_PER_SITE", 0)  # 0 = aucune limite
SCRAPE_ALL = env_bool("SCRAPE_ALL", False)  # True = scrape toutes les URLs mappees
MAX_ATTEMPTS = env_int("MAX_ATTEMPTS", 2)  # tentatives par site en cas d'erreur transitoire
MD_CHARS = min(env_int("MD_CHARS", 5000), 32000)  # markdown garde par page dans le xlsx
SCRAPE_TIMEOUT_MS = env_int("SCRAPE_TIMEOUT_MS", 45000)
PROXY = os.environ.get("FIRECRAWL_PROXY", "basic")  # basic evite les pages a 5 credits
CELL_MAX = 32000

PAGE_TIERS = [
    # (priorite, regex sur le chemin de l'URL)
    (1, r"(?:^|[/_-])(equipe|equipes|team|collaborateurs?|conseillers?|negociateurs?|notre-equipe|l-equipe|direction|gerant|dirigeants?|fondateurs?|nos-agents)(?:$|[/_.-])"),
    (2, r"(?:^|[/_-])(contact|contacts|contactez-nous|nous-contacter|nous-joindre|agence|agences|a-propos|apropos|qui-sommes-nous|presentation|nous-connaitre|notre-agence|about|l-agence|nos-agences)(?:$|[/_.-])"),
    (3, r"(?:^|[/_-])(mentions|mentions-legales|legal|legales|cgu|cgv|honoraires|rgpd|confidentialite|politique-de-confidentialite|informations-legales)(?:$|[/_.-])"),
]
PAGE_TIERS = [(p, re.compile(r, re.I)) for p, r in PAGE_TIERS]
PAGE_EXCLUDE = re.compile(
    r"(?:^|[/_-])(annonce|annonces|bien|biens|vente|ventes|achat|acheter|louer|location|locations|programme|programmes|"
    r"blog|actualite|actualites|article|articles|news|offre|offres|detail|fiche|ref)(?:$|[/_.-])",
    re.I,
)
ASSET_EXT = re.compile(
    r"\.(?:jpe?g|png|gif|webp|svg|ico|css|js|json|xml|zip|mp4|mp3|woff2?|ttf|docx?|xlsx?|pptx?|pdf)(?:$|\?)", re.I
)

# --------------------------------------------------------------------------- utils


def now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def log(*a) -> None:
    print(datetime.now().strftime("%H:%M:%S"), *a, flush=True)


def clean(v) -> str:
    return "" if v is None else str(v).strip()


def norm_name(s) -> str:
    return " ".join(clean(s).split()).casefold()


def fold(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn").lower()


def xl(v):
    """Valeur sure pour une cellule Excel."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return v
    s = ILLEGAL_CHARACTERS_RE.sub("", str(v))
    return s[:CELL_MAX]


def norm_site(url: str) -> str:
    url = clean(url).split()[0] if clean(url) else ""
    if not url:
        return ""
    if not re.match(r"^https?://", url, re.I):
        url = "https://" + url
    return url


def site_key(url: str) -> str:
    p = urlparse(url)
    host = (p.hostname or "").lower()
    host = host[4:] if host.startswith("www.") else host
    return host + p.path.rstrip("/")


def norm_link(u: str) -> str:
    u, _ = urldefrag(u.strip())
    p = urlparse(u)
    host = (p.hostname or "").lower()
    path = p.path.rstrip("/") or ""
    return f"{p.scheme}://{host}{path}" + (f"?{p.query}" if p.query else "")


# --------------------------------------------------------------------------- rotation de cles


class AllKeysDead(Exception):
    pass


class Stopped(Exception):
    pass


class ApiError(Exception):
    def __init__(self, status, msg, retriable=False):
        super().__init__(f"HTTP {status}: {msg}")
        self.status = status
        self.retriable = retriable


class Key:
    def __init__(self, value: str):
        self.value = value
        self.label = "fc-…" + value[-4:]
        self.dead = ""  # raison si morte
        self.cool_until = 0.0
        self.strikes = 0
        self.inflight = 0
        self.last_used = 0
        self.stamps = {"map": collections.deque(), "scrape": collections.deque()}
        self.calls_ok = 0


class KeyPool:
    def __init__(self, keys: list[str], per_key: int, stop: threading.Event):
        self.keys = [Key(k) for k in keys]
        self.per_key = max(1, per_key)
        self.stop = stop
        self.cv = threading.Condition()
        self.tick = itertools.count(1)

    def acquire(self, kind: str = "scrape") -> Key:
        with self.cv:
            while True:
                if self.stop.is_set():
                    raise Stopped()
                alive = [k for k in self.keys if not k.dead]
                if not alive:
                    raise AllKeysDead()
                now = time.time()
                limit = max(1, int(RATE_PER_MIN * 0.9)) if RATE_PER_MIN > 0 else 0  # marge de 10 %
                for k in alive:
                    dq = k.stamps[kind]
                    while dq and dq[0] <= now - 60:
                        dq.popleft()
                ready = [
                    k for k in alive
                    if k.cool_until <= now and k.inflight < self.per_key and (not limit or len(k.stamps[kind]) < limit)
                ]
                if ready:
                    k = min(ready, key=lambda x: (x.inflight, x.last_used))  # rotation equitable
                    k.inflight += 1
                    k.last_used = next(self.tick)
                    k.stamps[kind].append(now)
                    return k
                waits = [k.cool_until - now for k in alive if k.cool_until > now]
                if limit:
                    waits += [k.stamps[kind][0] + 60 - now for k in alive if len(k.stamps[kind]) >= limit]
                self.cv.wait(timeout=max(0.2, min(min(waits) if waits else 1.0, 5.0)))

    def release(self, k: Key) -> None:
        with self.cv:
            k.inflight -= 1
            self.cv.notify_all()

    def ok(self, k: Key) -> None:
        with self.cv:
            k.calls_ok += 1
            k.strikes = 0

    def kill(self, k: Key, reason: str) -> None:
        with self.cv:
            if not k.dead:
                k.dead = reason
                log(f"[cles] {k.label} retiree : {reason} ({sum(1 for x in self.keys if not x.dead)} restantes)")
            self.cv.notify_all()

    def cool(self, k: Key, seconds: float) -> None:
        with self.cv:
            k.strikes += 1
            k.cool_until = max(k.cool_until, time.time() + seconds)
            self.cv.notify_all()

    def summary(self) -> str:
        return " ".join(f"{k.label}:{'X' if k.dead else k.calls_ok}" for k in self.keys)


class Firecrawl:
    def __init__(self, pool: KeyPool, stop: threading.Event):
        self.pool = pool
        self.stop = stop
        self.local = threading.local()

    def session(self) -> requests.Session:
        if not hasattr(self.local, "s"):
            self.local.s = requests.Session()
        return self.local.s

    def call(self, path: str, payload: dict, read_timeout: float) -> dict:
        kind = "map" if path == "/map" else "scrape"
        errors = 0
        throttled = 0
        while True:
            k = self.pool.acquire(kind)  # peut lever AllKeysDead / Stopped
            try:
                r = self.session().post(
                    BASE_URL + path,
                    json=payload,
                    headers={"Authorization": f"Bearer {k.value}", "Content-Type": "application/json"},
                    timeout=(15, read_timeout),
                )
            except requests.RequestException as e:
                self.pool.release(k)
                errors += 1
                if errors > 4:
                    raise ApiError("net", str(e)[:200], retriable=True)
                self._sleep(min(30, 2**errors))
                continue
            self.pool.release(k)
            sc = r.status_code
            if sc == 200:
                try:
                    data = r.json()
                except ValueError:
                    raise ApiError(sc, "reponse non JSON", retriable=True)
                self.pool.ok(k)
                return data
            if sc == 401:
                self.pool.kill(k, "401 cle invalide")
                continue
            if sc == 402:
                self.pool.kill(k, "402 credits epuises")
                continue
            if sc == 429:
                throttled += 1
                try:
                    wait = float(r.headers.get("Retry-After", ""))
                except ValueError:
                    wait = min(60.0, 5.0 * 2 ** min(k.strikes, 4))
                self.pool.cool(k, wait)
                if throttled > 60:
                    raise ApiError(sc, "rate limit persistant", retriable=True)
                continue
            if sc >= 500 or sc == 408:
                errors += 1
                if errors > 3:
                    raise ApiError(sc, r.text[:200], retriable=True)
                self._sleep(min(30, 2**errors))
                continue
            raise ApiError(sc, r.text[:300], retriable=False)  # 400/403/404... propre au site

    def _sleep(self, s: float) -> None:
        self.stop.wait(s)

    def map(self, url: str) -> list[str]:
        d = self.call(
            "/map",
            {"url": url, "limit": 100000, "sitemap": "include", "location": {"country": "FR", "languages": ["fr-FR"]}},
            read_timeout=180,
        )
        out = []
        for l in d.get("links") or []:
            u = l.get("url") if isinstance(l, dict) else l
            if u:
                out.append(u)
        return out

    def scrape(self, url: str) -> dict:
        d = self.call(
            "/scrape",
            {
                "url": url,
                "formats": ["markdown", "links"],
                "onlyMainContent": False,  # on veut le footer / mentions legales
                "timeout": SCRAPE_TIMEOUT_MS,
                "proxy": PROXY,
                "location": {"country": "FR", "languages": ["fr-FR"]},
            },
            read_timeout=SCRAPE_TIMEOUT_MS / 1000 + 45,
        )
        return d.get("data") or {}


# --------------------------------------------------------------------------- extraction

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)*\.[A-Za-z]{2,}")
EMAIL_JUNK = re.compile(
    r"(\.(png|jpe?g|gif|webp|svg|css|js)$|sentry|wixpress|example\.|domain\.com|@2x|your-?email|votre-?email|"
    r"@email\.|noreply|no-reply|donotreply|u003e)",
    re.I,
)
FREE_DOMAINS = {
    "gmail.com", "orange.fr", "wanadoo.fr", "hotmail.com", "hotmail.fr", "outlook.com", "outlook.fr", "yahoo.fr",
    "yahoo.com", "free.fr", "sfr.fr", "neuf.fr", "laposte.net", "icloud.com", "live.fr", "live.com", "me.com",
    "bbox.fr", "numericable.fr", "aol.com", "gmx.fr", "protonmail.com", "proton.me",
}
GENERIC_LOCAL = re.compile(
    r"^(contact|contacts|info|infos|information|accueil|agence|bonjour|hello|secretariat|administration|admin|"
    r"location|locations|gestion|vente|ventes|transaction|transactions|syndic|compta|comptabilite|direction|"
    r"commercial|service|reception|rh|recrutement|communication|mail|webmaster|immo|immobilier|estimation|"
    r"copropriete|rgpd|dpo|support|reservation|rendezvous|rdv|devis|conciergerie|location-gestion)\b",
    re.I,
)
PHONE_RE = re.compile(r"(?<!\d)(?:\+33[\s.]?|0033[\s.]?|0)[1-9](?:[\s.\-]?\d{2}){4}(?!\d)")
SIRET_RE = re.compile(r"(?i)siret\D{0,20}((?:\d[\s.]?){14})")
SIREN_RE = re.compile(r"(?i)siren\D{0,20}((?:\d[\s.]?){9})(?!\d)")
CARTE_RE = re.compile(r"(?i)\b(CPI\s?\d{4}\s?\d{4}\s?\d{3}\s?\d{3})")
URL_RE = re.compile(r"https?://[^\s)\]>\"'<]+")

ROLE = (
    r"(?i:co-?g[ée]rant(?:e)?|g[ée]rant(?:e)?|directeur(?:\s+g[ée]n[ée]ral|\s+de\s+la\s+publication)?|directrice(?:\s+g[ée]n[ée]rale|\s+de\s+la\s+publication)?|"
    r"dirigeant(?:e)?|pr[ée]sident(?:e)?|fondateur|fondatrice|responsable\s+de\s+la\s+publication|"
    r"repr[ée]sent[ée]e?\s+l[ée]galement\s+par|PDG|titulaire\s+de\s+la\s+carte)"
)
TITLE = r"(?:M\.|Mme|Mlle|Madame|Monsieur|Mr\.?)"
WORD = r"[A-ZÉÈÀÂÊÎÔÛÇ][A-Za-zÀ-ÿ'’\-]+"
P_ROLE_NAME = re.compile(
    rf"({ROLE})\s*(?:de\s+(?:l['’]agence|la\s+soci[ée]t[ée])\s*)?(?:est\s+|[:,\-–—]\s*)?(?:{TITLE}\s+)?({WORD}(?:\s+{WORD}){{1,2}})"
)
P_NAME_ROLE = re.compile(rf"({WORD}(?:\s+{WORD}){{1,2}})\s*[,\-–—(:|]\s*({ROLE})")
NAME_STOP = {
    "mentions", "legales", "légales", "societe", "société", "agence", "immobilier", "immobiliere", "immobilière",
    "contact", "notre", "nos", "le", "la", "les", "du", "de", "des", "directeur", "directrice", "gerant", "gérant",
    "gerante", "gérante", "president", "président", "presidente", "présidente", "siret", "siren", "rcs", "tel",
    "tél", "email", "mail", "adresse", "capital", "hebergeur", "hébergeur", "site", "web", "equipe", "équipe",
    "publication", "cabinet", "groupe", "reseau", "réseau", "france", "conditions", "generales", "générales",
    "politique", "confidentialite", "confidentialité", "cookies", "sas", "sarl", "sasu", "eurl", "sci", "et",
    "qui", "est", "responsable", "directeur", "general", "général", "generale", "générale",
}
SOCIAL_RULES = [
    ("facebook", re.compile(r"https?://(?:[\w-]+\.)?facebook\.com/(?!sharer|share|plugins|tr\b|dialog|login)[^\s)\]>\x22'<]+", re.I)),
    ("instagram", re.compile(r"https?://(?:www\.)?instagram\.com/(?!p/|explore|accounts)[^\s)\]>\x22'<]+", re.I)),
    ("linkedin", re.compile(r"https?://(?:[\w-]+\.)?linkedin\.com/(?:company|school)/[^\s)\]>\x22'<]+", re.I)),
    ("linkedin_personnes", re.compile(r"https?://(?:[\w-]+\.)?linkedin\.com/in/[^\s)\]>\x22'<]+", re.I)),
    ("youtube", re.compile(r"https?://(?:www\.)?youtube\.com/(?:channel|c|user|@)[^\s)\]>\x22'<]*", re.I)),
    ("twitter_x", re.compile(r"https?://(?:www\.)?(?:twitter|x)\.com/(?!intent|share|home)[^\s)\]>\x22'<]+", re.I)),
    ("tiktok", re.compile(r"https?://(?:www\.)?tiktok\.com/@[^\s)\]>\x22'<]+", re.I)),
]


def host_base(url: str) -> str:
    h = (urlparse(url).hostname or "").lower()
    return h[4:] if h.startswith("www.") else h


def extract_emails(text: str, site_host: str):
    text = re.sub(r"\s*[\[(]\s*(?:at|arobase)\s*[\])]\s*", "@", text, flags=re.I)
    seen, nomin, gener, autres = set(), [], [], []
    for m in EMAIL_RE.findall(text):
        e = m.strip(".-_").lower()
        if e in seen or EMAIL_JUNK.search(e):
            continue
        seen.add(e)
        local, _, dom = e.partition("@")
        same = bool(site_host) and (dom == site_host or dom.endswith("." + site_host) or site_host.endswith("." + dom))
        if not (same or dom in FREE_DOMAINS):
            autres.append(e)
        elif GENERIC_LOCAL.match(local):
            gener.append(e)
        else:
            nomin.append(e)
    return nomin, gener, autres


def extract_phones(text: str) -> list[str]:
    out = []
    for m in PHONE_RE.findall(text):
        d = re.sub(r"\D", "", m)
        if d.startswith("0033"):
            d = "0" + d[4:]
        elif d.startswith("33") and len(d) == 11:
            d = "0" + d[2:]
        if len(d) == 10:
            f = " ".join(d[i : i + 2] for i in range(0, 10, 2))
            if f not in out:
                out.append(f)
    return out


def trim_name(name: str) -> str:
    toks = name.split()
    while toks and fold(toks[0].strip("-'’")) in {fold(x) for x in NAME_STOP}:
        toks.pop(0)
    while toks and fold(toks[-1].strip("-'’")) in {fold(x) for x in NAME_STOP}:
        toks.pop()
    return " ".join(toks) if len(toks) >= 2 else ""


def extract_managers(text: str) -> list[tuple[str, str]]:
    found, seen = [], set()
    for m in P_ROLE_NAME.finditer(text):
        role, name = m.group(1), trim_name(m.group(2))
        if name and name.lower() not in seen:
            seen.add(name.lower())
            found.append((name, role.strip().lower()))
    for m in P_NAME_ROLE.finditer(text):
        name, role = trim_name(m.group(1)), m.group(2)
        if name and name.lower() not in seen:
            seen.add(name.lower())
            found.append((name, role.strip().lower()))
    return found[:6]


def email_for_name(name: str, emails: list[str]) -> str:
    toks = [fold(t) for t in re.split(r"[\s\-']+", name) if len(t) > 1]
    if len(toks) < 2:
        return ""
    first, last = toks[0], toks[-1]
    for e in emails:
        loc = fold(e.split("@")[0])
        if (len(last) >= 3 and last in loc) or loc in (first, first + last, first[0] + last):
            return e
    return ""


def extract_socials(text: str, links: list[str]) -> dict:
    blob = "\n".join(links) + "\n" + text
    out = {}
    for key, rx in SOCIAL_RULES:
        vals = []
        for m in rx.finditer(blob):
            u = m.group(0).rstrip(".,;:'\"")
            if u not in vals:
                vals.append(u)
        if vals:
            out[key] = " | ".join(vals[:10 if key == "linkedin_personnes" else 3])
    return out


def rank_pages(links: list[str], home: str) -> list[str]:
    base = host_base(home)
    cand = {}
    for u in links:
        p = urlparse(u)
        h = host_base(u)
        if not h or not (h == base or h.endswith("." + base)):
            continue
        if ASSET_EXT.search(u):
            continue
        path = (p.path or "/").lower()
        if SCRAPE_ALL:
            tier = 9
        else:
            if PAGE_EXCLUDE.search(path):
                continue
            tier = next((t for t, rx in PAGE_TIERS if rx.search(path)), None)
            if tier is None:
                continue
        key = norm_link(u)
        depth = path.strip("/").count("/")
        if key not in cand or (tier, depth) < cand[key][0]:
            cand[key] = ((tier, depth, len(path)), u)
    home_key = norm_link(home)
    ordered = [u for _, u in sorted(cand.values(), key=lambda x: x[0])]
    ordered = [u for u in ordered if norm_link(u) != home_key]
    if MAX_PAGES_PER_SITE > 0:
        ordered = ordered[: max(0, MAX_PAGES_PER_SITE - 1)]
    return [home] + ordered


# --------------------------------------------------------------------------- traitement d'un site


def process_site(fc: Firecrawl, site: str) -> tuple[dict, list[dict]]:
    """Retourne (champs agregés, pages). Peut lever Stopped / AllKeysDead."""
    rec = {"site_url": site, "scraped_at": now_iso()}
    host = host_base(site)
    map_links, map_err = [], ""
    try:
        map_links = fc.map(site)
    except ApiError as e:
        map_err = str(e)
    rec["n_urls_mappees"] = len(map_links)
    rec["urls_mappees"] = "\n".join(map_links)[:CELL_MAX]
    rec["erreur_map"] = map_err

    targets = rank_pages(map_links, site)
    pages, texts, all_links, home_fail = [], [], [], None
    for i, u in enumerate(targets):
        try:
            d = fc.scrape(u)
            md = d.get("markdown") or ""
            meta = d.get("metadata") or {}
            links = [x for x in (d.get("links") or []) if isinstance(x, str)]
            pages.append({"url": u, "titre": clean(meta.get("title")), "http": meta.get("statusCode"), "erreur": clean(meta.get("error")), "markdown": md[:MD_CHARS]})
            texts.append(md)
            all_links += links
            if i == 0:
                rec["site_titre"] = clean(meta.get("title"))
                rec["site_description"] = clean(meta.get("description"))
                rec["site_url_finale"] = clean(meta.get("url"))
        except ApiError as e:
            pages.append({"url": u, "titre": "", "http": e.status if isinstance(e.status, int) else None, "erreur": str(e)[:300], "markdown": ""})
            if i == 0:
                home_fail = e
    rec["n_pages_scrapees"] = sum(1 for p in pages if p["markdown"])
    rec["pages_scrapees"] = "\n".join(p["url"] for p in pages if p["markdown"])[:CELL_MAX]

    text = "\n\n".join(texts)
    nomin, gener, autres = extract_emails(text + "\n" + "\n".join(all_links), host)
    rec["emails_nominatifs"] = "; ".join(nomin)
    rec["emails_generiques"] = "; ".join(gener)
    rec["emails_autres_domaines"] = "; ".join(autres[:15])
    rec["telephones"] = "; ".join(extract_phones(text)[:15])
    mgrs = extract_managers(text)
    rec["gerants_detectes"] = "; ".join(f"{n} ({r})" for n, r in mgrs)
    best = mgrs[0][0] if mgrs else ""
    rec["gerant_principal_detecte"] = best
    rec["email_gerant_probable"] = next((email_for_name(n, nomin) for n, _ in mgrs if email_for_name(n, nomin)), "")
    m = SIRET_RE.search(text)
    rec["siret"] = re.sub(r"\D", "", m.group(1)) if m else ""
    m = SIREN_RE.search(text)
    rec["siren"] = re.sub(r"\D", "", m.group(1)) if m else (rec["siret"][:9] if rec["siret"] else "")
    m = CARTE_RE.search(text)
    rec["carte_pro"] = re.sub(r"\s+", " ", m.group(1)).upper() if m else ""
    rec.update({f"social_{k}": v for k, v in extract_socials(text, all_links).items()})

    if rec["n_pages_scrapees"] > 0:
        rec["statut"] = "ok"
    elif home_fail is not None and not home_fail.retriable:
        rec["statut"] = "site_inaccessible"  # definitif (400/403/404...)
        rec["erreur"] = str(home_fail)
    else:
        rec["statut"] = "error"  # transitoire : sera retente
        rec["erreur"] = str(home_fail or map_err)
    return rec, pages


# --------------------------------------------------------------------------- donnees d'entree


def read_sheet(ws):
    it = ws.iter_rows(values_only=True)
    header = [clean(h) for h in next(it)]
    rows = [list(r) + [None] * (len(header) - len(r)) for r in it if any(clean(c) for c in r)]
    return header, [r[: len(header)] for r in rows]


def load_agencies(path: str):
    wb = load_workbook(path, read_only=True, data_only=True)
    sheets = wb.worksheets
    h1, r1 = read_sheet(sheets[0])
    h2, r2 = read_sheet(sheets[1]) if len(sheets) > 1 else ([], [])
    h3, r3 = read_sheet(sheets[2]) if len(sheets) > 2 else ([], [])
    wb.close()

    # --- identifiants : nom de l'agence ; homonymes departages par code postal + ville
    names = [clean(r[0]) for r in r1]
    counts: dict[str, int] = {}
    for n in names:
        counts[norm_name(n)] = counts.get(norm_name(n), 0) + 1
    used: set[str] = set()
    ids = []
    for n, r in zip(names, r1):
        i = " ".join(n.split())
        if counts[norm_name(n)] > 1:
            m = re.search(r"(\d{5})\s+([^,]+)$", clean(r[1]) if len(r) > 1 else "")
            if m:
                i = f"{i} ({m.group(1)} {m.group(2).strip()})"
        base, k = i, 2
        while i.casefold() in used:
            i = f"{base} #{k}"
            k += 1
        used.add(i.casefold())
        ids.append(i)

    # --- feuille 2 (gerants) : memes lignes que la feuille 1 -> par position, sinon par nom
    def by_name(rows):
        d: dict[str, list] = {}
        for r in rows:
            d.setdefault(norm_name(r[0]), []).append(r)
        return d

    g_pos = len(r2) == len(r1)
    g_name = by_name(r2)
    a_name = by_name(r3)
    seen: dict[str, int] = {}
    agencies = []
    for idx, (i, r) in enumerate(zip(ids, r1)):
        nn = norm_name(r[0])
        k = seen.get(nn, 0)
        seen[nn] = k + 1
        g = r2[idx] if g_pos else (g_name.get(nn, [])[k] if k < len(g_name.get(nn, [])) else None)
        a = a_name.get(nn, [])[k] if k < len(a_name.get(nn, [])) else None
        agencies.append({"id": i, "row": r, "g": g, "a": a, "site": norm_site(r[3]) if len(r) > 3 else ""})
    return (h1, h2, h3), agencies


# --------------------------------------------------------------------------- etat (reprise)


class State:
    def __init__(self, directory: str):
        self.dir = directory
        os.makedirs(directory, exist_ok=True)
        self.rpath = os.path.join(directory, "results.jsonl")
        self.ppath = os.path.join(directory, "pages.jsonl")
        self.results: dict[str, dict] = {}
        self.pages: dict[tuple, dict] = {}
        self.lock = threading.Lock()
        self._load()
        self.rf = open(self.rpath, "a", encoding="utf-8")
        self.pf = open(self.ppath, "a", encoding="utf-8")

    def _load(self):
        for path, store, keyf in (
            (self.rpath, self.results, lambda o: o["id"]),
            (self.ppath, self.pages, lambda o: (o["id"], o["url"])),
        ):
            if not os.path.exists(path):
                continue
            with open(path, encoding="utf-8") as f:
                for line in f:
                    try:
                        o = json.loads(line)
                        store[keyf(o)] = o
                    except (ValueError, KeyError):
                        pass  # ligne tronquee par un arret brutal

    def is_done(self, agency_id: str) -> bool:
        r = self.results.get(agency_id)
        return bool(r) and (r.get("statut") != "error" or r.get("attempts", 1) >= MAX_ATTEMPTS)

    def write(self, records: list[dict], pages: list[dict]):
        with self.lock:
            for p in pages:
                self.pf.write(json.dumps(p, ensure_ascii=False) + "\n")
                self.pages[(p["id"], p["url"])] = p
            self.pf.flush()
            for r in records:
                self.rf.write(json.dumps(r, ensure_ascii=False) + "\n")
                self.results[r["id"]] = r
            self.rf.flush()

    def close(self):
        self.rf.close()
        self.pf.close()


# --------------------------------------------------------------------------- export xlsx

SCRAPE_COLS = [
    ("statut", "Statut scraping"), ("erreur", "Erreur"), ("site_url", "Site (URL utilisée)"),
    ("site_url_finale", "Site (URL finale)"), ("site_titre", "Titre du site"), ("site_description", "Description du site"),
    ("emails_nominatifs", "E-mails nominatifs (trouvés)"), ("emails_generiques", "E-mails génériques (trouvés)"),
    ("emails_autres_domaines", "E-mails autres domaines"), ("telephones", "Téléphones"),
    ("gerant_principal_detecte", "Gérant / dirigeant détecté (à vérifier)"), ("gerants_detectes", "Dirigeants détectés (nom + rôle)"),
    ("email_gerant_probable", "E-mail du gérant (probable)"), ("siret", "SIRET"), ("siren", "SIREN"), ("carte_pro", "Carte pro"),
    ("social_facebook", "Facebook"), ("social_instagram", "Instagram"), ("social_linkedin", "LinkedIn (page)"),
    ("social_linkedin_personnes", "LinkedIn (personnes)"), ("social_youtube", "YouTube"), ("social_twitter_x", "X / Twitter"),
    ("social_tiktok", "TikTok"), ("n_urls_mappees", "Nb URLs mappées"), ("n_pages_scrapees", "Nb pages scrapées"),
    ("pages_scrapees", "Pages scrapées"), ("urls_mappees", "Toutes les URLs mappées"), ("erreur_map", "Erreur map"),
    ("note", "Note"), ("scraped_at", "Date scraping (UTC)"),
]


def head_cell(ws, text):
    c = WriteOnlyCell(ws, value=text)
    c.font = Font(bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor="1F4E78")
    return c


def export_xlsx(heads, agencies, state: State, out_path: str):
    h1, h2, h3 = heads
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    tmp = out_path + ".tmp"
    wb = Workbook(write_only=True)

    ws = wb.create_sheet("Agences")
    ws.freeze_panes = "B2"
    ws.column_dimensions["A"].width = 45
    cols = ["ID (nom de l'agence)"] + h1 + h2[1:] + h3[1:] + [t for _, t in SCRAPE_COLS]
    ws.append([head_cell(ws, c) for c in cols])
    for a in agencies:
        rec = state.results.get(a["id"], {})
        g = (a["g"] or [None] * len(h2))[1:]
        av = (a["a"] or [None] * len(h3))[1:]
        row = [a["id"]] + a["row"] + g + av
        row += [rec.get(k, "" if rec else None) for k, _ in SCRAPE_COLS]
        if not a["site"] and not rec:
            row[-len(SCRAPE_COLS)] = "no_site"
        ws.append([xl(v) for v in row])

    wp = wb.create_sheet("Pages")
    wp.freeze_panes = "A2"
    wp.append([head_cell(wp, c) for c in ["ID (nom de l'agence)", "URL", "Titre", "HTTP", "Erreur", f"Markdown (max {MD_CHARS} car.)"]])
    for (aid, _), p in sorted(state.pages.items(), key=lambda kv: (kv[0][0], kv[1].get("ord", 0))):
        wp.append([xl(aid), xl(p["url"]), xl(p.get("titre")), p.get("http"), xl(p.get("erreur")), xl(p.get("markdown"))])
    wb.save(tmp)
    os.replace(tmp, out_path)
    log(f"[export] {out_path} : {len(agencies)} agences, {len(state.pages)} pages")


# --------------------------------------------------------------------------- main


def load_keys() -> list[str]:
    raw = [os.environ.get("FIRECRAWL_API_KEYS", "")]
    raw += [v for k, v in sorted(os.environ.items()) if re.fullmatch(r"FIRECRAWL_API_KEY(?:_\d+)?", k)]
    keys = []
    for blob in raw:
        for k in re.split(r"[\s,;]+", blob or ""):
            if k.strip() and k.strip() not in keys:
                keys.append(k.strip())
    return keys


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--export-only", action="store_true", help="regenere le xlsx depuis l'etat sans appeler Firecrawl")
    ap.add_argument("--max-sites", type=int, default=0, help="debug : ne traiter que N sites")
    args = ap.parse_args()

    heads, agencies = load_agencies(INPUT_XLSX)
    state = State(STATE_DIR)
    log(f"{len(agencies)} agences chargees, {len(state.results)} deja en etat")

    if args.export_only:
        export_xlsx(heads, agencies, state, OUTPUT_XLSX)
        return 0

    keys = load_keys()
    if not keys:
        log("ERREUR : aucune cle. Definir FIRECRAWL_API_KEYS (separees par virgule ou retour a la ligne).")
        return 2
    log(f"{len(keys)} cle(s) Firecrawl en rotation")

    # --- sans site : resultat immediat ; avec site : regroupe par site (franchises = 1 seul scraping)
    groups: dict[str, list[dict]] = {}
    for a in agencies:
        if state.is_done(a["id"]):
            continue
        if not a["site"]:
            state.write([{"id": a["id"], "statut": "no_site", "attempts": 1, "scraped_at": now_iso()}], [])
            continue
        groups.setdefault(site_key(a["site"]), []).append(a)
    todo = list(groups.values())
    if args.max_sites:
        todo = todo[: args.max_sites]
    total = len(todo)
    log(f"{total} site(s) a traiter")

    stop = threading.Event()
    reason = {"v": ""}
    signal.signal(signal.SIGTERM, lambda *_: (reason.update(v="signal"), stop.set()))
    pool = KeyPool(keys, PER_KEY_CONCURRENCY, stop)
    fc = Firecrawl(pool, stop)
    q: "queue.Queue[list[dict]]" = queue.Queue()
    for g in todo:
        q.put(g)
    deadline = time.time() + TIME_BUDGET_MIN * 60 if TIME_BUDGET_MIN > 0 else None
    done = itertools.count(1)
    progress = {"n": 0}
    t0 = time.time()

    def worker():
        while not stop.is_set():
            if deadline and time.time() > deadline:
                reason["v"] = reason["v"] or "deadline"
                stop.set()
                return
            try:
                grp = q.get_nowait()
            except queue.Empty:
                return
            first = grp[0]
            try:
                rec, pages = process_site(fc, first["site"])
            except Stopped:
                return
            except AllKeysDead:
                reason["v"] = "keys_dead"
                stop.set()
                return
            except Exception as e:  # bug imprevu sur un site : on le note et on continue
                rec, pages = {"site_url": first["site"], "statut": "error", "erreur": f"exception: {e!r}"[:300], "scraped_at": now_iso()}, []
            records = []
            for n, a in enumerate(grp):
                r = dict(rec)
                r["id"] = a["id"]
                r["attempts"] = state.results.get(a["id"], {}).get("attempts", 0) + 1
                if n > 0:
                    r["note"] = f"site partagé avec « {first['id']} » (pages listées sur cette ligne)"
                records.append(r)
            for i, p in enumerate(pages):
                p["id"], p["ord"] = first["id"], i
            state.write(records, pages)
            progress["n"] += 1
            c = next(done)
            if c % 25 == 0:
                el = time.time() - t0
                log(f"[{c}/{total}] {c / el * 60:.1f} sites/min | cles {pool.summary()}")

    n_threads = max(1, min(PER_KEY_CONCURRENCY * len(keys), 150))
    threads = [threading.Thread(target=worker, daemon=True) for _ in range(n_threads)]
    for t in threads:
        t.start()
    try:
        while any(t.is_alive() for t in threads):
            time.sleep(1)
    except KeyboardInterrupt:
        reason["v"] = "signal"
        stop.set()
    for t in threads:
        t.join(timeout=300)

    state.close()
    state = State(STATE_DIR)  # relecture propre
    export_xlsx(heads, agencies, state, OUTPUT_XLSX)

    remaining = sum(1 for a in agencies if not state.is_done(a["id"]))
    ok = sum(1 for r in state.results.values() if r.get("statut") == "ok")
    log(f"Fin : {progress['n']} sites traites ce run | {ok} agences ok | {remaining} restantes | cles {pool.summary()} | arret={reason['v'] or 'fini'}")
    if remaining == 0:
        return 0
    if reason["v"] == "keys_dead":
        log("Toutes les cles sont epuisees/invalides : ajoutez des cles puis relancez (la reprise est automatique).")
        return 3
    if progress["n"] == 0:
        return 4
    return 10


if __name__ == "__main__":
    sys.exit(main())
