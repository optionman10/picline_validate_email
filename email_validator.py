"""
email_validator.py
------------------
Core email validation logic: syntax, naming conventions, domain checks
(disposable blocklist, typo suggestions, DNS MX lookup with caching).

No GUI or Excel dependencies here so this module is unit-testable.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field, asdict
from enum import Enum

try:
    import dns.resolver
    import dns.exception
    HAS_DNS = True
except ImportError:  # pragma: no cover
    HAS_DNS = False


class Severity(str, Enum):
    ERROR = "ERROR"
    WARNING = "WARNING"


class Status(str, Enum):
    VALID = "VALID"
    INVALID = "INVALID"
    VALID_WITH_WARNINGS = "VALID_WITH_WARNINGS"


@dataclass
class Issue:
    code: str
    severity: Severity
    message: str


@dataclass
class ValidationResult:
    email: str
    status: Status
    issues: list[Issue] = field(default_factory=list)
    suggestion: str | None = None

    @property
    def error_description(self) -> str:
        """Human-readable description of all issues (errors first)."""
        ordered = sorted(self.issues, key=lambda i: i.severity != Severity.ERROR)
        return "; ".join(f"[{i.severity.value}] {i.message}" for i in ordered) or "OK"

    def to_dict(self) -> dict:
        d = asdict(self)
        d["status"] = self.status.value
        return d


# --------------------------------------------------------------------------
# Static data
# --------------------------------------------------------------------------

DISPOSABLE_DOMAINS: set[str] = {
    "mailinator.com", "tempmail.com", "temp-mail.org", "10minutemail.com",
    "guerrillamail.com", "guerrillamail.net", "yopmail.com", "trashmail.com",
    "fakeinbox.com", "sharklasers.com", "getnada.com", "dispostable.com",
    "maildrop.cc", "mintemail.com", "throwawaymail.com", "mailnesia.com",
    "tempinbox.com", "moakt.com", "emailondeck.com", "spamgourmet.com",
}

# Popular domains used for typo suggestions (difflib close-match).
POPULAR_DOMAINS: list[str] = [
    "gmail.com", "yahoo.com", "hotmail.com", "outlook.com", "aol.com",
    "icloud.com", "live.com", "msn.com", "mail.com", "protonmail.com",
    "proton.me", "ymail.com", "googlemail.com", "comcast.net", "me.com",
    "yahoo.co.uk", "yahoo.co.in", "hotmail.co.uk", "rediffmail.com",
    "zoho.com", "gmx.com", "gmx.net", "fastmail.com", "yandex.com",
]

GENERIC_LOCAL_PARTS: set[str] = {
    "info", "admin", "administrator", "noreply", "no-reply", "support",
    "sales", "contact", "help", "helpdesk", "office", "mail", "postmaster",
    "webmaster", "abuse", "marketing", "hr", "enquiries", "enquiry",
    "service", "services", "team", "hello", "root", "test", "demo",
}

# RFC 5322-ish practical pattern: dot-atom local part, valid domain labels.
_EMAIL_RE = re.compile(
    r"^(?=.{1,254}$)"                                  # total length
    r"(?P<local>[A-Za-z0-9!#$%&'*+/=?^_`{|}~-]+"
    r"(?:\.[A-Za-z0-9!#$%&'*+/=?^_`{|}~-]+)*)"         # dot-atom local
    r"@"
    r"(?P<domain>(?:[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}"
    r"[A-Za-z0-9])?\.)+[A-Za-z]{2,63})$"               # labels + TLD
)

_DOMAIN_ONLY_RE = re.compile(
    r"^(?:[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?\.)+[A-Za-z]{2,63}$"
)

_DNS_TIMEOUT = 4.0


# --------------------------------------------------------------------------
# DNS MX cache
# --------------------------------------------------------------------------

class DomainChecker:
    """MX-record lookup with per-domain caching. Degrades gracefully offline."""

    def __init__(self, timeout: float = _DNS_TIMEOUT, enabled: bool = True):
        self.timeout = timeout
        self.enabled = enabled and HAS_DNS
        self._cache: dict[str, bool | None] = {}
        self._resolver = None
        if self.enabled:
            self._resolver = dns.resolver.Resolver()
            self._resolver.lifetime = timeout
            self._resolver.timeout = timeout

    def has_mx(self, domain: str) -> bool | None:
        """True=mail-capable, False=definitely not, None=couldn't determine."""
        key = domain.lower()
        if key in self._cache:
            return self._cache[key]
        result: bool | None = None
        if self.enabled and self._resolver is not None:
            try:
                answers = self._resolver.resolve(key, "MX")
                result = len(answers) > 0
                if not result:
                    # Some domains have A record but no MX (implicit MX rule)
                    a = self._resolver.resolve(key, "A")
                    result = len(a) > 0
            except dns.resolver.NXDOMAIN:
                result = False
            except dns.resolver.NoAnswer:
                result = False
            except (dns.resolver.NoNameservers, dns.exception.Timeout,
                    dns.resolver.YXDOMAIN, Exception):
                result = None  # network down / blocked — don't penalize email
        self._cache[key] = result
        return result


# --------------------------------------------------------------------------
# Validation
# --------------------------------------------------------------------------

def _suggest_domain(domain: str) -> str | None:
    import difflib
    matches = difflib.get_close_matches(domain.lower(), POPULAR_DOMAINS, n=1, cutoff=0.82)
    return matches[0] if matches else None


def _looks_random(local: str) -> bool:
    """Heuristic: long, vowel-poor alphanumeric mix like 'xk7qz9w2b'."""
    cleaned = re.sub(r"[^a-z0-9]", "", local.lower())
    if len(cleaned) < 10:
        return False
    vowels = sum(c in "aeiou" for c in cleaned)
    digits = sum(c.isdigit() for c in cleaned)
    return vowels / len(cleaned) < 0.20 and digits > 0


def validate_email(
    email: object,
    domain_checker: DomainChecker | None = None,
    check_dns: bool = True,
) -> ValidationResult:
    """Validate one email address. Returns a ValidationResult (never raises)."""
    issues: list[Issue] = []
    raw = email

    def err(code: str, msg: str):
        issues.append(Issue(code, Severity.ERROR, msg))

    def warn(code: str, msg: str):
        issues.append(Issue(code, Severity.WARNING, msg))

    # -- basic sanity -------------------------------------------------------
    if raw is None or (isinstance(raw, float) and raw != raw):  # NaN
        return ValidationResult("", Status.INVALID,
                                [Issue("EMPTY", Severity.ERROR, "Empty cell")])
    text = str(raw)
    if text != text.strip():
        warn("WHITESPACE", "Leading/trailing whitespace around the email")
    text = text.strip()

    if not text:
        return ValidationResult(text, Status.INVALID,
                                [Issue("EMPTY", Severity.ERROR, "Empty value")])

    if text.count("@") != 1:
        err("FORMAT", f"Must contain exactly one '@' (found {text.count('@')})")
        return _finish(text, issues)

    m = _EMAIL_RE.match(text)
    local, _, domain = text.partition("@")

    # -- syntax -------------------------------------------------------------
    if not m:
        if len(text) > 254:
            err("LENGTH", "Total length exceeds 254 characters")
        if not local:
            err("LOCAL_EMPTY", "Local part (before '@') is empty")
        elif len(local) > 64:
            err("LOCAL_LENGTH", "Local part exceeds 64 characters")
        elif local.startswith(".") or local.endswith("."):
            err("LOCAL_DOTS", "Local part starts or ends with a dot")
        elif ".." in local:
            err("LOCAL_DOTS", "Local part contains consecutive dots")
        elif re.search(r"[^A-Za-z0-9!#$%&'*+/=?^_`{|}~.-]", local):
            bad = sorted(set(re.findall(r"[^A-Za-z0-9!#$%&'*+/=?^_`{|}~.-]", local)))
            err("LOCAL_CHARS", f"Illegal character(s) in local part: {' '.join(bad)}")
        if not domain:
            err("DOMAIN_EMPTY", "Domain part (after '@') is empty")
        elif not _DOMAIN_ONLY_RE.match(domain):
            err("DOMAIN_FORMAT", "Domain has invalid format (labels/TLD)")
    else:
        # Regex passed — still apply naming-convention *warnings*.
        if len(local) > 64:
            err("LOCAL_LENGTH", "Local part exceeds 64 characters")

    # -- naming conventions (warnings) --------------------------------------
    if local:
        if ".." in local and not any(i.code == "LOCAL_DOTS" for i in issues):
            err("LOCAL_DOTS", "Local part contains consecutive dots")
        if local.lower() in GENERIC_LOCAL_PARTS:
            warn("GENERIC", f"'{local}' is a generic/role mailbox, not a person")
        if local.isdigit():
            warn("ALL_DIGITS", "Local part is all digits — likely low quality")
        if _looks_random(local):
            warn("RANDOM", "Local part looks randomly generated")

    # -- domain checks -------------------------------------------------------
    suggestion: str | None = None
    if domain and _DOMAIN_ONLY_RE.match(domain or ""):
        d = domain.lower()
        if d in DISPOSABLE_DOMAINS:
            err("DISPOSABLE", f"'{d}' is a known disposable/temporary email domain")
        else:
            sug = _suggest_domain(d)
            if sug and sug != d:
                suggestion = f"{local}@{sug}"
                warn("TYPO", f"Possible typo — did you mean '{sug}'?")
            if check_dns and domain_checker is not None:
                mx = domain_checker.has_mx(d)
                if mx is False:
                    err("NO_MX", f"Domain '{d}' does not exist or accepts no mail (no MX record)")
                # mx is None → network issue; leave unreported

    return _finish(text, issues, suggestion)


def _finish(text: str, issues: list[Issue], suggestion: str | None = None) -> ValidationResult:
    has_error = any(i.severity == Severity.ERROR for i in issues)
    has_warn = any(i.severity == Severity.WARNING for i in issues)
    if has_error:
        status = Status.INVALID
    elif has_warn:
        status = Status.VALID_WITH_WARNINGS
    else:
        status = Status.VALID
    return ValidationResult(text, status, issues, suggestion)


__all__ = [
    "Severity", "Status", "Issue", "ValidationResult",
    "DomainChecker", "validate_email",
    "DISPOSABLE_DOMAINS", "POPULAR_DOMAINS", "GENERIC_LOCAL_PARTS",
]
