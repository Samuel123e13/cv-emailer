"""Find published professional emails for a company domain via the Hunter.io API.

This uses Hunter's official, public API (https://hunter.io/api-documentation) —
a sanctioned, Terms-of-Service-compliant way to look up business email
addresses that companies have published. No scraping is involved.

Only the standard library is used (urllib), so there are no extra dependencies.
"""

from __future__ import annotations

import json
import urllib.parse
import urllib.request
import urllib.error
from dataclasses import dataclass

API_BASE = "https://api.hunter.io/v2"

# Position keywords that suggest a recruiting / hiring contact. Used only to
# *rank* results so the most relevant people show up first — nothing is hidden.
RECRUITER_HINTS = (
    "recruit", "talent", "hiring", "people", "human resources", "hr",
    "staffing", "acquisition", "university", "campus", "intern", "early career",
)


class HunterError(Exception):
    pass


@dataclass
class Contact:
    name: str
    email: str
    company: str
    position: str
    department: str
    confidence: int
    type: str  # "personal" | "generic"

    def is_recruiterish(self) -> bool:
        text = f"{self.position} {self.department}".lower()
        return any(h in text for h in RECRUITER_HINTS)

    def as_recipient(self) -> dict:
        """Shape expected by the rest of the app (store/sender)."""
        return {"name": self.name, "email": self.email, "company": self.company}


def _request(path: str, params: dict) -> dict:
    url = f"{API_BASE}/{path}?" + urllib.parse.urlencode(params)
    try:
        with urllib.request.urlopen(url, timeout=30) as resp:
            return json.load(resp)
    except urllib.error.HTTPError as exc:
        try:
            payload = json.load(exc)
            errs = payload.get("errors") or [{}]
            detail = errs[0].get("details") or errs[0].get("id") or exc.reason
        except Exception:
            detail = exc.reason
        if exc.code == 401:
            raise HunterError("Hunter rejected the API key (401). Check the key in Settings.")
        if exc.code == 429:
            raise HunterError("Hunter rate/usage limit reached (429). You may be out of free searches.")
        raise HunterError(f"Hunter API error {exc.code}: {detail}")
    except urllib.error.URLError as exc:
        raise HunterError(f"Could not reach Hunter.io: {exc.reason}")


def account_info(api_key: str) -> dict:
    """Return {'plan': str, 'used': int, 'available': int}."""
    data = _request("account", {"api_key": api_key}).get("data", {})
    searches = (data.get("requests", {}) or {}).get("searches", {}) or {}
    return {
        "plan": data.get("plan_name", "?"),
        "used": searches.get("used"),
        "available": searches.get("available"),
    }


def normalise_domain(text: str) -> str:
    """Accept 'Acme', 'acme.com', 'https://www.acme.com/careers' -> 'acme.com'."""
    text = text.strip()
    if not text:
        return ""
    if "://" in text:
        text = urllib.parse.urlparse(text).netloc
    text = text.split("/")[0]
    if text.lower().startswith("www."):
        text = text[4:]
    return text.lower()


def domain_search(api_key: str, domain: str, limit: int = 10,
                  department: str | None = None,
                  email_type: str | None = None) -> list[Contact]:
    """Look up emails for one company domain. Costs one Hunter search.

    ``department`` (e.g. "hr") and ``email_type`` ("personal"/"generic") are
    optional server-side filters. Results are sorted so recruiter-like
    positions and higher confidence come first.
    """
    domain = normalise_domain(domain)
    if not domain:
        return []
    params = {"domain": domain, "limit": limit, "api_key": api_key}
    if department:
        params["department"] = department
    if email_type:
        params["type"] = email_type

    data = _request("domain-search", params).get("data", {})
    org = data.get("organization") or domain
    contacts: list[Contact] = []
    for e in data.get("emails", []) or []:
        first = (e.get("first_name") or "").strip()
        last = (e.get("last_name") or "").strip()
        name = " ".join(p for p in (first, last) if p)
        contacts.append(Contact(
            name=name,
            email=(e.get("value") or "").strip(),
            company=org,
            position=(e.get("position") or "").strip(),
            department=(e.get("department") or "").strip(),
            confidence=int(e.get("confidence") or 0),
            type=(e.get("type") or "").strip(),
        ))

    contacts.sort(key=lambda c: (not c.is_recruiterish(), -c.confidence))
    return contacts
