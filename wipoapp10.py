#!/usr/bin/env python3
"""WIPO PCT weekly publication search.

Tkinter desktop app. Downloads bibliographic rows from the documented
PATENTSCOPE REST service when you paste your own access token, or loads a
local CSV, XML, or XLSX file. No sample rows are included.
#Get a patentscope web token
Optional Excel support: pip install openpyxl
"""

import csv
import datetime
import os
import re
import socket
import threading
import time
import tkinter as tk
import webbrowser
import xml.etree.ElementTree as ET
from tkinter import filedialog, messagebox, ttk
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    openpyxl = None

COUNTRY_MAP = {
    "JP": {"code": "JP", "name": "Japan"},
    "CN": {"code": "CN", "name": "China"},
    "AU": {"code": "AU", "name": "Australia"},
    "AT": {"code": "AT", "name": "Austria"},
    "DE": {"code": "DE", "name": "Germany"},
    "FR": {"code": "FR", "name": "France"},
    "KR": {"code": "KR", "name": "Republic of Korea"},
    "CA": {"code": "CA", "name": "Canada"},
    "UK": {"code": "GB", "name": "United Kingdom (GB)"},
    "IT": {"code": "IT", "name": "Italy"},
}

SERVICE_URL = "https://patentscope.wipo.int/search/rest/patents"
TOKEN_HELP_URL = "https://patentscope.wipo.int/search/en/help/webservice.jsf"
SEARCH_PAGE = "https://patentscope.wipo.int/search/en/result.jsf?query="
USER_AGENT = "WipoPctDesktopSearch/1.0 (personal bibliographic lookup)"
PAGE_SIZE = 50
MAX_DOWNLOAD = 500
MAX_BYTES = 20 * 1024 * 1024

PUB_KEYS = {
    "publicationnumber",
    "internationalpublicationnumber",
    "wonumber",
    "pn",
}
RECORD_TAGS = {
    "patent",
    "document",
    "record",
    "docpage",
    "item",
    "entry",
    "wopatent",
    "patentdocument",
    "queryresult",
}
NAME_KEYS = {"name", "orgname", "applicantname", "agentname", "organizationname"}


class PatentscopeError(Exception):
    def __init__(self, message, status=None):
        super().__init__(message)
        self.status = status


def default_pub_date(today=None):
    today = today or datetime.date.today()
    delta = (today.weekday() - 3) % 7
    if delta == 0:
        delta = 7
    return today - datetime.timedelta(days=delta)


def tag_key(tag):
    if not isinstance(tag, str):
        return ""
    return re.sub(r"[^a-z0-9]", "", tag.rsplit("}", 1)[-1].lower())


def own_text(node):
    parts = []
    if node.text and node.text.strip():
        parts.append(node.text.strip())
    for child in list(node):
        if child.tail and child.tail.strip():
            parts.append(child.tail.strip())
    return " ".join(parts)


def element_text(node):
    return " ".join("".join(node.itertext()).split())


def norm_date(value):
    if isinstance(value, datetime.datetime):
        return value.date().isoformat()
    if isinstance(value, datetime.date):
        return value.isoformat()
    digits = re.sub(r"\D", "", str(value or ""))
    if len(digits) >= 8:
        year, month, day = int(digits[:4]), int(digits[4:6]), int(digits[6:8])
        if 1900 <= year <= 2100 and 1 <= month <= 12 and 1 <= int(day) <= 31:
            return "%04d-%02d-%02d" % (year, month, day)
    return ""


def parse_user_date(text):
    found = norm_date(text)
    if not found:
        raise ValueError("Publication date must be YYYY-MM-DD.")
    return datetime.datetime.strptime(found, "%Y-%m-%d").date()


def clean_class(value):
    value = " ".join((value or "").upper().split())
    if not value:
        return ""
    if not re.fullmatch(r"[A-Z0-9][A-Z0-9 /_,.\-]{0,40}", value):
        raise ValueError(
            "Classification may contain only letters, digits, spaces, and / - _ , ."
        )
    return value


def looks_like_pub(value):
    compact = re.sub(r"[^A-Za-z0-9]", "", str(value or "")).upper()
    return re.match(r"WO\d{4}\d{5,7}", compact) is not None


def normalize_pub_num(value):
    raw = " ".join(str(value or "").split())
    compact = re.sub(r"[^A-Za-z0-9]", "", raw).upper()
    match = re.match(r"WO(\d{4})(\d{5,7})", compact)
    if match:
        return "WO/%s/%s" % (match.group(1), match.group(2))
    return raw or "N/A"


def get_patentscope_url(pub_num):
    compact = re.sub(r"[^A-Za-z0-9]", "", str(pub_num or "")).upper()
    match = re.match(r"(WO\d{4}\d{5,7})", compact)
    if match:
        return "https://patentscope.wipo.int/search/en/detail.jsf?docId=" + match.group(1)
    return SEARCH_PAGE + quote(str(pub_num or "WO"))


def clean_text(value, limit=500):
    text = " ".join(str(value or "").split())
    return text[:limit] if text else "N/A"


def make_record(pub, country, applicant, law_firm, ipc, cpc, date, title):
    code = (country or "").strip().upper()
    if not re.fullmatch(r"[A-Z]{2}", code) or code == "WO":
        code = "N/A"
    return {
        "pub_num": normalize_pub_num(pub) if pub else "N/A",
        "country": code,
        "applicant": clean_text(applicant),
        "law_firm": clean_text(law_firm),
        "ipc_code": clean_text(ipc, 240),
        "cpc_code": clean_text(cpc, 240),
        "date": norm_date(date) or "N/A",
        "title": clean_text(title, 800),
    }


def dedupe(records):
    kept, seen = [], set()
    for rec in records:
        pub = rec.get("pub_num") or ""
        if not pub or pub == "N/A" or pub in seen:
            continue
        seen.add(pub)
        kept.append(rec)
    return kept


def redact(text, token):
    text = text or ""
    if token:
        text = text.replace(token, "***")
    return text


def parent_map(root):
    mapping = {}
    for parent in root.iter():
        for child in list(parent):
            mapping[child] = parent
    return mapping


def is_descendant(node, ancestor, parents):
    current = parents.get(node)
    while current is not None:
        if current is ancestor:
            return True
        current = parents.get(current)
    return False


def record_elements(root):
    found = [elem for elem in root.iter() if tag_key(elem.tag) in RECORD_TAGS]
    if not found:
        found = []
        for elem in root.iter():
            for child in list(elem):
                key = tag_key(child.tag)
                text = own_text(child) or element_text(child)
                if key in PUB_KEYS and looks_like_pub(text):
                    found.append(elem)
                    break
    if not found:
        return []
    parents = parent_map(root)
    outer = []
    for elem in found:
        if any(other is not elem and is_descendant(elem, other, parents) for other in found):
            continue
        if elem is not root:
            outer.append(elem)
    return outer


def names_under(elem, role_tags):
    found = []
    for node in elem.iter():
        if tag_key(node.tag) not in role_tags:
            continue
        picked = ""
        for sub in node.iter():
            if sub is not node and tag_key(sub.tag) in NAME_KEYS:
                picked = own_text(sub) or element_text(sub)
                if picked:
                    break
        if not picked:
            picked = own_text(node)
        if picked and picked not in found and len(picked) < 300:
            found.append(picked)
    return found


def applicant_country(elem):
    for node in elem.iter():
        if tag_key(node.tag) != "applicant":
            continue
        for sub in node.iter():
            if tag_key(sub.tag) not in {"country", "countrycode"}:
                continue
            code = (own_text(sub) or element_text(sub)).upper()
            if re.fullmatch(r"[A-Z]{2}", code) and code != "WO":
                return code
    return ""


def publication_number(elem):
    for node in elem.iter():
        if tag_key(node.tag) in PUB_KEYS:
            text = own_text(node) or element_text(node)
            if looks_like_pub(text):
                return text
    parents = parent_map(elem)
    for node in elem.iter():
        if tag_key(node.tag) != "documentid":
            continue
        keys = []
        current = node
        while current in parents:
            current = parents[current]
            keys.append(tag_key(current.tag))
            if current is elem:
                break
        if "applicationreference" in keys:
            continue
        country = number = ""
        for sub in node.iter():
            key = tag_key(sub.tag)
            if key == "country":
                country = (own_text(sub) or "").upper()
            elif key in {"docnumber", "number"}:
                number = own_text(sub)
        digits = re.sub(r"\D", "", number or "")
        if country == "WO" and len(digits) >= 10:
            return "WO" + digits
    return ""


def publication_date(elem):
    parents = parent_map(elem)
    for node in elem.iter():
        key = tag_key(node.tag)
        if key in {"publicationdate", "pubdate"}:
            found = norm_date(own_text(node) or element_text(node))
            if found:
                return found
        if key != "date":
            continue
        keys = []
        current = node
        while current in parents:
            current = parents[current]
            keys.append(tag_key(current.tag))
            if current is elem:
                break
        if "publicationreference" in keys:
            found = norm_date(own_text(node))
            if found:
                return found
    return ""


def title_of(elem):
    fallback = ""
    for node in elem.iter():
        if tag_key(node.tag) not in {"title", "inventiontitle"}:
            continue
        text = own_text(node) or element_text(node)
        if not text or len(text) > 800:
            continue
        lang = ""
        for attr, val in node.attrib.items():
            if str(attr).lower().endswith("lang"):
                lang = str(val).lower()
        if lang.startswith("en"):
            return text
        fallback