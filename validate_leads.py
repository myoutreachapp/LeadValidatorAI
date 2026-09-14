#!/usr/bin/env python3
"""
validate_leads.py — MyOutreach lead validation pipeline

Usage:
    python validate_leads.py --leads leads.xlsx --config campaign_config.yaml \
        --titles job_titles_reference.yaml --output validated_output.xlsx \
        [--tal-dir /path/to/tal/files] [--reference leadscale_export.csv]

One script, swap the --config file per campaign. See campaign_config_*.yaml
for the format.

WHAT THIS DOES NOT DO YET (stubbed — needs your API keys):
  - ZeroBounce email validation  -> see validate_email_zerobounce()
  - LinkedIn profile append       -> see append_linkedin_profile()
Both are stubbed with clear TODOs. Everything else runs for real.
"""

import argparse
import os
import re
import sys
from pathlib import Path

import pandas as pd
import phonenumbers
import requests
from phonenumbers import geocoder
import yaml
from rapidfuzz import fuzz

ZEROBOUNCE_API_KEY = os.environ.get("ZEROBOUNCE_API_KEY")

try:
    import pycountry
except ImportError:
    pycountry = None


# ---------------------------------------------------------------------------
# Loaders
# ---------------------------------------------------------------------------

def load_config(path):
    with open(path) as f:
        return yaml.safe_load(f)


def load_job_titles(path):
    with open(path) as f:
        titles = yaml.safe_load(f)
    return expand_titles_with_acronyms(titles)


_ACRONYM_SKIP_WORDS = {"of", "and", "for", "the", "&", "-", "/"}


def derive_acronym(title):
    """'Chief Information Officer' -> 'CIO'. Returns None for titles that don't
    look like a multi-word official title (need at least 2 significant words)."""
    words = re.split(r"[\s/,-]+", str(title).strip())
    letters = [w[0] for w in words if w and w.lower() not in _ACRONYM_SKIP_WORDS and w[0].isalpha()]
    if len(letters) >= 2:
        return "".join(letters).upper()
    return None


def expand_titles_with_acronyms(job_titles_ref):
    """A reference list containing 'Chief Information Officer' but not 'CIO' will
    badly fail to match a real-world title like 'CIO at J&T Bank' — acronyms and
    their spelled-out form share almost no characters, so fuzzy matching alone
    can't bridge that gap no matter the threshold. Auto-deriving the acronym from
    every multi-word reference title fixes this everywhere, not just for titles
    someone remembered to list explicitly."""
    expanded = {}
    for function, titles in (job_titles_ref or {}).items():
        all_titles = list(titles)
        for t in titles:
            acronym = derive_acronym(t)
            if acronym and acronym not in all_titles:
                all_titles.append(acronym)
        expanded[function] = all_titles
    return expanded


def load_leads(path):
    path = Path(path)
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path)
    return pd.read_excel(path)


DOMAIN_COLUMN_ALIASES = ["domain", "website", "url", "web address", "web", "site", "company website", "company url"]


def normalize_domain(value):
    """Strips protocol, www, path, and trailing slashes so 'https://www.acme.com/en'
    and 'acme.com' both normalize to 'acme.com' for matching."""
    if value is None:
        return ""
    v = str(value).strip().lower()
    v = re.sub(r"^https?://", "", v)
    v = re.sub(r"^www\.", "", v)
    v = v.split("/")[0]  # drop any path
    return v.strip()


ALT_DOMAIN_COLUMN_ALIASES = [
    "alternate domains", "alternate domain", "also known as", "aka domains", "aka domain",
    "other domains", "backup domains", "alt domains", "additional domains", "secondary domains",
]


COMPANY_NAME_COLUMN_ALIASES = [
    "company name", "company", "account name", "account", "company/account name", "org name", "organization",
]


def normalize_company_name(value):
    """Light normalization only (case, whitespace, trailing punctuation) — this is
    for an EXACT match on name, same strictness level as the domain match, not a
    fuzzy comparison."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    v = str(value).strip().lower()
    v = re.sub(r"\s+", " ", v)
    v = v.rstrip(".,")
    return v


def split_domain_list(value):
    """Splits a cell that may contain several domains separated by ; , / or newlines."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return []
    parts = re.split(r"[;,/\n]+", str(value))
    return [normalize_domain(p) for p in parts if p.strip()]


def load_tal(tal_file, sheet):
    tal_file = Path(tal_file)
    if tal_file.suffix.lower() == ".csv":
        df = pd.read_csv(tal_file)
    else:
        df = pd.read_excel(tal_file, sheet_name=sheet)
    # normalize column names we rely on regardless of exact casing in the sheet
    cols = {c.lower().strip(): c for c in df.columns}
    domain_col = next((cols[alias] for alias in DOMAIN_COLUMN_ALIASES if alias in cols), None)
    if not domain_col:
        raise ValueError(
            f"TAL file has no domain/website column — found: {list(df.columns)}. "
            f"Expected one of these column names: {DOMAIN_COLUMN_ALIASES}"
        )
    alt_domain_col = next((cols[alias] for alias in ALT_DOMAIN_COLUMN_ALIASES if alias in cols), None)
    name_col = next((cols[alias] for alias in COMPANY_NAME_COLUMN_ALIASES if alias in cols), None)

    df["_domain_norm"] = df[domain_col].apply(normalize_domain)
    if alt_domain_col:
        df["_alt_domains"] = df[alt_domain_col].apply(split_domain_list)
    else:
        df["_alt_domains"] = [[] for _ in range(len(df))]
    # every domain (primary + aliases) a row will match on
    df["_all_domains"] = df.apply(lambda r: set([r["_domain_norm"]] + r["_alt_domains"]), axis=1)
    df["_name_norm"] = df[name_col].apply(normalize_company_name) if name_col else ""
    return df


def load_reference_file(path):
    """Integrate / Leadscale export used to catch company-name typos before hard-failing."""
    if path is None:
        return None
    path = Path(path)
    df = pd.read_csv(path) if path.suffix.lower() == ".csv" else pd.read_excel(path)
    cols = {c.lower(): c for c in df.columns}
    company_col = cols.get("company_name") or cols.get("company")
    if not company_col:
        return None
    return set(df[company_col].astype(str).str.strip().str.lower())


# ---------------------------------------------------------------------------
# Field access (uses column_mapping so the same script works across sources)
# ---------------------------------------------------------------------------

def get_field(row, mapping, field):
    col = mapping.get(field)
    if col is None or col not in row.index:
        return None
    val = row[col]
    if pd.isna(val):
        return None
    return val


def domain_from_email(email):
    if not email or "@" not in str(email):
        return None
    return str(email).split("@")[-1].strip().lower()


# ---------------------------------------------------------------------------
# Standardization
# ---------------------------------------------------------------------------

def standardize_name(name):
    if not name:
        return name
    return str(name).strip().title()


def standardize_row_formatting(row, mapping):
    """Returns a dict of cleaned display fields — does not touch pass/fail logic."""
    out = {}
    for field in ("first_name", "last_name", "company"):
        val = get_field(row, mapping, field)
        out[field] = standardize_name(val) if val else val
    email = get_field(row, mapping, "email")
    out["email"] = str(email).strip().lower() if email else None
    return out


# ---------------------------------------------------------------------------
# Individual checks — each returns (passed: bool, reason: str or None)
# ---------------------------------------------------------------------------

_zerobounce_cache = {}

# Statuses ZeroBounce returns. "valid" and "catch-all" are treated as passes —
# "valid" passes clean. "catch-all" also passes (not a bounce risk like
# "invalid"/"spamtrap"), but ZeroBounce genuinely can't confirm the specific
# mailbox exists — so it's flagged for a manual burner-email test rather than
# silently trusted. Adjust if your delivery team wants a stricter policy.
_ZB_PASS_STATUSES = {"valid", "catch-all"}
_ZB_MANUAL_CHECK_STATUSES = {"catch-all"}


def check_email_valid_zerobounce(email):
    """Returns (ok, block_reason, manual_check_note). manual_check_note is
    informational only — it never blocks the lead, it just tells the team to
    run the manual burner-email test before treating this address as fully live."""
    if not email:
        return False, "no email", None

    if email in _zerobounce_cache:
        status = _zerobounce_cache[email]
    else:
        if not ZEROBOUNCE_API_KEY:
            raise RuntimeError(
                "ZEROBOUNCE_API_KEY environment variable not set. "
                "Run: export ZEROBOUNCE_API_KEY=your_key_here  (before running this script)"
            )
        try:
            r = requests.get(
                "https://api.zerobounce.net/v2/validate",
                params={"api_key": ZEROBOUNCE_API_KEY, "email": email},
                timeout=15,
            )
            r.raise_for_status()
            data = r.json()
            status = data.get("status", "unknown")
        except requests.RequestException as e:
            return False, f"zerobounce: request failed ({e})", None
        _zerobounce_cache[email] = status

    if status in _ZB_PASS_STATUSES:
        note = None
        if status in _ZB_MANUAL_CHECK_STATUSES:
            note = ("zerobounce: catch-all — send manual test email from a burner account; "
                    "a reply like 'user left' or 'mailbox not found' means this address is invalid")
        return True, None, note
    return False, f"zerobounce:{status}", None



US_STATE_ABBREVIATIONS = {
    "alabama": "AL", "alaska": "AK", "arizona": "AZ", "arkansas": "AR", "california": "CA",
    "colorado": "CO", "connecticut": "CT", "delaware": "DE", "florida": "FL", "georgia": "GA",
    "hawaii": "HI", "idaho": "ID", "illinois": "IL", "indiana": "IN", "iowa": "IA",
    "kansas": "KS", "kentucky": "KY", "louisiana": "LA", "maine": "ME", "maryland": "MD",
    "massachusetts": "MA", "michigan": "MI", "minnesota": "MN", "mississippi": "MS", "missouri": "MO",
    "montana": "MT", "nebraska": "NE", "nevada": "NV", "new hampshire": "NH", "new jersey": "NJ",
    "new mexico": "NM", "new york": "NY", "north carolina": "NC", "north dakota": "ND", "ohio": "OH",
    "oklahoma": "OK", "oregon": "OR", "pennsylvania": "PA", "rhode island": "RI",
    "south carolina": "SC", "south dakota": "SD", "tennessee": "TN", "texas": "TX", "utah": "UT",
    "vermont": "VT", "virginia": "VA", "washington": "WA", "west virginia": "WV", "wisconsin": "WI",
    "wyoming": "WY", "district of columbia": "DC",
}


def normalize_state(state, country=None):
    """Returns a 2-letter US state code, 'CANADA', or the cleaned-up input if
    neither — good enough to key into a region map."""
    if not state:
        if country and str(country).strip().lower() == "canada":
            return "CANADA"
        return None
    s = str(state).strip()
    if len(s) == 2:
        return s.upper()
    return US_STATE_ABBREVIATIONS.get(s.lower(), s)


def build_region_lookup(region_definitions):
    """region_definitions: {'Southwest': ['NV','UT',...], ...} -> {'NV': 'Southwest', ...}"""
    lookup = {}
    for region, codes in (region_definitions or {}).items():
        for code in codes:
            lookup[code.strip().upper()] = region  # last region wins if a code is listed twice
    return lookup


def check_region_quota(state, country, region_lookup, region_counts, target_per_region):
    """Only meaningful once a lead has otherwise passed everything else — caller should
    only call this for provisionally-passing leads, same pattern as cap_per_company."""
    if not region_lookup or not target_per_region:
        return True, None, None  # region distribution not configured for this campaign
    key = normalize_state(state, country)
    region = region_lookup.get(key) if key else None
    if region is None:
        return True, None, None  # lead's location isn't part of any defined region — no quota applies
    region_counts[region] = region_counts.get(region, 0) + 1
    if region_counts[region] > target_per_region:
        return False, f"region quota exceeded ({region}: {target_per_region} max)", region
    return True, None, region


SENIORITY_RANK_NAMES = {"Individual Contributor": 0, "Manager": 1, "Director": 2, "VP": 3, "SVP": 4, "EVP": 5, "Chief": 6}

# Checked highest-rank first so a title matching multiple levels (e.g. "VP, Sr. Manager")
# gets credited for the highest one it actually contains.
_SENIORITY_KEYWORDS = [
    (6, ["chief ", " ceo", " cto", " cio", " ciso", " coo", " cfo",
         "chief information officer", "chief technology officer", "chief information security officer"]),
    (5, ["evp", "executive vice president"]),
    (4, ["svp", "senior vice president"]),
    (3, ["vp", "vice president"]),
    (2, ["director", "head of"]),
    (1, ["senior manager", "manager"]),
]


def detect_seniority_rank(job_title):
    """Returns an integer rank (0-6) based on seniority keywords in the title.
    0 means no manager/director/VP/etc. keyword was found — individual contributor
    level (Engineer, Architect, Lead, Analyst, Specialist, etc.)."""
    t = " " + str(job_title).lower() + " "

    # "Associate/Assistant Director" reads as a plain "Director" keyword match,
    # but is a genuinely lower level in standard corporate title conventions —
    # check this BEFORE the general ladder so it doesn't get full Director credit.
    if re.search(r"\b(associate|assistant)\s+director\b", t):
        return 1

    # A consulting/advisory title that happens to mention a C-level acronym
    # (e.g. "Consultant - CIO Advisory", "CTO Advisory Services Consultant")
    # describes their PRACTICE AREA, not their own role — don't let the
    # acronym read as if they actually hold that executive title. Genuine
    # seniority words (Director, VP, Partner, etc.) still count normally;
    # only the chief/C-suite tier is skipped for these titles.
    if re.search(r"\bconsultant\b|\badvisory\b|\badvisor\b", t):
        for rank, keywords in _SENIORITY_KEYWORDS:
            if rank >= 6:
                continue
            if any(kw in t for kw in keywords):
                return rank
        return 0

    for rank, keywords in _SENIORITY_KEYWORDS:
        if any(kw in t for kw in keywords):
            return rank
    return 0


def check_seniority(job_title, minimum_seniority):
    if not minimum_seniority:
        return True, None  # no seniority floor configured for this campaign
    if not job_title:
        return False, "no job title to assess seniority"
    required_rank = SENIORITY_RANK_NAMES.get(minimum_seniority)
    if required_rank is None:
        return True, None  # unrecognized config value — don't silently fail every lead
    actual_rank = detect_seniority_rank(job_title)
    if actual_rank < required_rank:
        return False, f"seniority below minimum ({minimum_seniority}) — title: '{job_title}'"
    return True, None


def check_phone(phone, expected_country):
    """Phone VALIDITY (parseable, real number) is a hard check. Region-vs-country
    match is informational only — people keep old numbers after moving states,
    provinces, or countries, so a mismatch there isn't a reliable fail signal."""
    if not phone:
        return False, "no phone", None
    phone_str = str(phone).strip()
    # try parsing with expected country as a hint for national-format numbers
    region_hint = None
    if pycountry and expected_country:
        try:
            region_hint = pycountry.countries.lookup(expected_country).alpha_2
        except LookupError:
            region_hint = None
    try:
        parsed = phonenumbers.parse(phone_str, region_hint)
    except phonenumbers.NumberParseException:
        # retry assuming it's missing a leading '+' but has a full country code
        try:
            parsed = phonenumbers.parse("+" + re.sub(r"\D", "", phone_str))
        except phonenumbers.NumberParseException:
            return False, "phone: unparseable", None

    if not phonenumbers.is_valid_number(parsed):
        return False, "phone: invalid syntax", None

    # Toll-free numbers (800/888/877/866/855/844/833) aren't tied to any real
    # region — they're shared across the whole NANP — so there's nothing
    # meaningful to note here.
    if phonenumbers.number_type(parsed) == phonenumbers.PhoneNumberType.TOLL_FREE:
        return True, None, None

    phone_region = geocoder.region_code_for_number(parsed)  # e.g. 'US'
    expected_region = region_hint
    if expected_region and phone_region and phone_region != expected_region:
        note = f"phone region ({phone_region}) differs from stated country ({expected_region}) — common if the person moved or kept an old number, not necessarily wrong"
        return True, None, note
    return True, None, None


def check_country_allowed(country, allowed_countries):
    """Empty/unset allowed_countries means no restriction — for global-targeting
    campaigns, so you don't have to hand-list all ~200 countries."""
    if not allowed_countries:
        return True, None
    if not country:
        return False, "no country"
    if str(country).strip() not in allowed_countries:
        return False, f"country '{country}' not in permitted geo list"
    return True, None


def check_industry_allowed(industry, permitted_industries):
    """Empty/unset permitted_industries means no restriction — backward-compatible
    with campaigns that don't define this (e.g. industry implied entirely by TAL)."""
    if not permitted_industries:
        return True, None
    if not industry:
        return False, "no industry to check against permitted list"
    industry_norm = str(industry).strip().lower()
    if any(term.strip().lower() == industry_norm for term in permitted_industries):
        return True, None
    return False, f"industry '{industry}' not in permitted list {permitted_industries}"


def get_geo_bucket(country, geo_function_rules):
    """A bucket with an empty 'countries' list matches ANY country — same
    global-campaign convention as check_country_allowed."""
    for bucket, cfg in geo_function_rules.items():
        if not cfg["countries"] or country in cfg["countries"]:
            return bucket
    return None


DOWNGRADE_QUALIFIERS = {"associate", "assistant", "interim", "acting"}


def normalize_title_for_matching(title):
    """Strips punctuation (commas, pipes, parens, ampersands, etc.) to spaces
    before fuzzy matching. rapidfuzz's token_set_ratio tokenizes on whitespace
    only — it treats 'workplace,' and 'workplace' as two entirely different
    tokens, which silently wrecks the score for any real-world title with
    commas or other punctuation in it (which is most of them)."""
    return re.sub(r"[^\w\s]", " ", str(title)).strip()


def check_job_function(job_title, country, geo_function_rules, job_titles_ref, fuzzy_threshold=80):
    """Never hard-fails on fuzzy match quality — job title is used to find the best
    matching function and record a confidence note, but a mediocre or even poor fuzzy
    score does not reject the lead on its own. Only a genuinely missing title, or a
    country with no function-rule bucket configured, blocks the lead here."""
    if not job_title:
        return False, "no job title", None

    bucket = get_geo_bucket(country, geo_function_rules)
    if bucket is None:
        return False, "country not mapped to a function-rule bucket", None

    allowed_functions = geo_function_rules[bucket]["allowed_functions"]
    title_lower = normalize_title_for_matching(job_title).lower()

    # First pass: literal containment. If a reference title (e.g. "CIO") appears
    # as a clear whole word/phrase anywhere in the lead's title, that's a confident
    # match regardless of how much other descriptive text surrounds it — a title
    # like "CIO at J&T Bank" or "CIO, Member of the Executive Board" has the real
    # signal right there; the rest is just company/context, not noise that should
    # dilute the match. Fuzzy scoring alone can under-count this when a short,
    # strong signal is surrounded by a lot of unrelated extra words.
    #
    # Exception: a downgrade qualifier immediately before the matched phrase
    # (e.g. "Associate Director", "Assistant Director") means the person doesn't
    # actually hold that title — it just contains it as a substring. Skip the
    # automatic containment win here and let it fall through to fuzzy scoring
    # instead, rather than blindly accepting it.
    for function in allowed_functions:
        for ref_title in job_titles_ref.get(function, []):
            if len(ref_title) < 2:
                continue
            ref_norm = normalize_title_for_matching(ref_title).lower()
            m = re.search(r"\b" + re.escape(ref_norm) + r"\b", title_lower)
            if m:
                preceding_words = title_lower[:m.start()].split()
                preceding_word = preceding_words[-1] if preceding_words else ""
                if preceding_word in DOWNGRADE_QUALIFIERS:
                    continue
                return True, None, function

    # Fall back to fuzzy scoring for titles that don't literally contain a
    # reference phrase, but are still plausibly related through different wording.
    #
    # Same downgrade-qualifier concern applies here, and matters even more:
    # token_set_ratio scores by token PRESENCE, not word order — so "Associate
    # Director | IT Product Management" scores a perfect 100 against reference
    # title "IT Director", because both "it" and "director" appear somewhere in
    # the lead's title, even though not adjacent and not the person's real title.
    # Skip any reference title containing "director" from fuzzy scoring when the
    # lead's own title contains a downgraded "Associate/Assistant Director" —
    # this only blocks credit from Director-branded reference titles; Architect/
    # Lead-style reference titles are untouched and still score normally.
    has_downgraded_director = bool(re.search(
        r"\b(" + "|".join(DOWNGRADE_QUALIFIERS) + r")\s+director\b", title_lower))

    best_function, best_score = None, 0
    for function in allowed_functions:
        for ref_title in job_titles_ref.get(function, []):
            ref_norm = normalize_title_for_matching(ref_title).lower()
            if has_downgraded_director and "director" in ref_norm:
                continue
            score = fuzz.token_set_ratio(title_lower, ref_norm)
            if score > best_score:
                best_score, best_function = score, function

    if best_score >= 90:
        note = None  # confident match, nothing to flag
        return True, note, best_function
    elif best_score >= fuzzy_threshold:
        note = f"fuzzy match (score {best_score:.0f}, function: {best_function}) — confirm manually"
        return True, note, best_function
    elif best_score >= 50:
        note = f"weak fuzzy match (score {best_score:.0f}, closest function: {best_function}) — please review"
        return True, note, best_function
    else:
        # genuinely no plausible relation to any allowed function — this is the
        # floor that was accidentally missing: everything above still passes with
        # a note (that's the intentional leniency fix), but a title with nothing
        # in common with any target function has to actually fail here, or
        # nothing is left gating job function/seniority at all for campaigns
        # that use an explicit title list instead of the seniority dropdown.
        return False, f"job title '{job_title}' doesn't plausibly match any allowed function (best score {best_score:.0f}, closest: {best_function})", None


def check_tal_match(domain, company, tal_df):
    """Matches on EITHER an exact domain match OR an exact company-name match —
    both are equally strict (no fuzzy logic), so either one is sufficient."""
    domain_norm = domain.lower() if domain else None
    name_norm = normalize_company_name(company) if company else None
    if not domain_norm and not name_norm:
        return False, "no domain or company name to match against TAL", None

    domain_hit = tal_df["_all_domains"].apply(lambda domains: bool(domain_norm) and domain_norm in domains)
    name_hit = tal_df["_name_norm"].apply(lambda n: bool(name_norm) and n == name_norm) if "_name_norm" in tal_df else False
    match = tal_df[domain_hit | name_hit]
    if match.empty:
        return False, "company/domain not found on TAL (exact match required, on either domain or company name)", None
    return True, None, match.iloc[0]


def check_reference_match(company_name, reference_set):
    """Cross-check against Integrate/Leadscale export — catches typos that would
    otherwise silently fail the lead on the platform side."""
    if reference_set is None:
        return True, None  # no reference file provided, skip
    if not company_name:
        return False, "no company name to check against reference export"
    name_norm = str(company_name).strip().lower()
    if name_norm in reference_set:
        return True, None
    # fuzzy-suggest the likely intended match so a human can fix it fast
    best_match, best_score = None, 0
    for ref_name in reference_set:
        score = fuzz.ratio(name_norm, ref_name)
        if score > best_score:
            best_score, best_match = score, ref_name
    if best_score >= 85:
        return False, f"company name typo suspected — closest reference match: '{best_match}' ({best_score}%)"
    return False, "company name not found in Integrate/Leadscale export"


def check_exclusions(company, job_title, industry, exclusions, skip_industry_check=False):
    """skip_industry_check=True when the lead is already confirmed on the TAL —
    a pre-vetted account list overrides the generic industry exclusion, same as
    it overrides the permitted-industries inclusion list."""
    for term in exclusions.get("company_name_contains", []):
        if company and term.lower() in str(company).lower():
            return False, f"excluded: company name contains '{term}'"
    for term in exclusions.get("job_title_contains", []):
        if job_title and term.lower() in str(job_title).lower():
            return False, f"excluded: job title contains '{term}'"
    if not skip_industry_check:
        for term in exclusions.get("industry_contains", []):
            if industry and term.lower() in str(industry).lower():
                return False, f"excluded: industry contains '{term}'"
    return True, None


def check_watch_time(minutes, minimum):
    if minutes is None:
        return True, None  # no watch-time data in this file, skip
    try:
        return (float(minutes) >= minimum, None if float(minutes) >= minimum
                else f"watch time {minutes} < {minimum} min minimum")
    except (ValueError, TypeError):
        return False, "watch time value unparseable"


def check_qualifying_questions(row, mapping, questions_cfg):
    for question_col, acceptable_answers in questions_cfg.items():
        if question_col not in row.index:
            continue
        answer = row[question_col]
        if pd.isna(answer) or str(answer).strip() not in acceptable_answers:
            return False, f"qualifying question '{question_col}' answer not acceptable"
    return True, None



# Maps internal (verbose, debug-style) reason strings to a clean category a
# supplier/delivery team can action for a chargeback dispute. Add patterns as
# you discover new reason shapes — anything unmatched falls into "Other / QA Review".
REJECTION_CATEGORY_RULES = [
    ("duplicate email", "Duplicate"),
    ("no email", "Invalid Contact Info"),
    ("zerobounce", "Invalid Contact Info"),
    ("phone: unparseable", "Invalid Contact Info"),
    ("phone: invalid syntax", "Invalid Contact Info"),
    ("no phone", "Invalid Contact Info"),
    ("country code", "Invalid Contact Info"),
    ("not in permitted geo list", "Wrong Geo"),
    ("not in permitted list", "Not ICP Match — Industry"),
    ("not mapped to a function-rule bucket", "Wrong Geo"),
    ("does not match allowed functions", "Not ICP Match — Job Title/Function"),
    ("seniority below minimum", "Below Seniority Requirement"),
    ("not found on TAL", "Not on Target Account List"),
    ("company name not found", "Company Name Mismatch / Typo"),
    ("company name typo suspected", "Company Name Mismatch / Typo"),
    ("cap per company exceeded", "Cap Per Company Exceeded"),
    ("region quota exceeded", "Region Quota Exceeded"),
    ("excluded:", "Excluded — Competitor/Restricted"),
    ("watch time", "Below Engagement Threshold"),
    ("qualifying question", "Qualifying Question Failed"),
    ("REVIEW:", "Needs Manual QA Review"),
]


def categorize_rejection(reasons_str):
    """Returns (category, is_manual_qa_flag) for the first matching rule."""
    if not reasons_str:
        return None, False
    for pattern, category in REJECTION_CATEGORY_RULES:
        if pattern.lower() in reasons_str.lower():
            return category, category == "Needs Manual QA Review"
    return "Other / QA Review", True


def append_linkedin_profile(first_name, last_name, company):
    """
    STUB — wire up Clay (or Apollo/RocketReach) here.
    Return a LinkedIn profile URL string, or None if no match found.
    """
    return None  # placeholder


# ---------------------------------------------------------------------------
# Main pipeline
# ---------------------------------------------------------------------------

def validate(leads_path, config_path, titles_path, tal_dir, reference_path, output_path):
    config = load_config(config_path)
    job_titles_ref = load_job_titles(titles_path)
    df = load_leads(leads_path)
    mapping = config["column_mapping"]

    tal_df = None
    if config.get("tal", {}).get("file"):
        tal_file = Path(tal_dir or ".") / config["tal"]["file"]
        tal_df = load_tal(tal_file, config["tal"]["sheet"])

    reference_set = load_reference_file(reference_path)

    results = []
    seen_emails = set()
    company_counts = {}
    region_counts = {}
    region_lookup = build_region_lookup(config.get("region_distribution", {}).get("regions"))
    target_per_region = config.get("region_distribution", {}).get("target_per_region")

    for idx, row in df.iterrows():
        reasons = []
        clean = standardize_row_formatting(row, mapping)

        email = clean["email"]
        first_name = clean["first_name"]
        last_name = clean["last_name"]
        company = clean["company"]
        job_title = get_field(row, mapping, "job_title")
        country = get_field(row, mapping, "country")
        phone = get_field(row, mapping, "phone")
        watch_time = get_field(row, mapping, "watch_time_minutes")
        industry = get_field(row, mapping, "industry")
        state = get_field(row, mapping, "state")

        domain = normalize_domain(get_field(row, mapping, "domain")) or domain_from_email(email)

        # --- duplicate check ---
        if email and email in seen_emails:
            reasons.append("duplicate email in this batch")
        else:
            if email:
                seen_emails.add(email)

        # --- email validity (catch-all passes but gets flagged for manual test) ---
        ok, reason, email_manual_check_note = check_email_valid_zerobounce(email)
        if not ok:
            reasons.append(reason)

        # --- phone (validity blocks; region-vs-country mismatch is a note only) ---
        ok, reason, phone_region_note = check_phone(phone, country)
        if not ok:
            reasons.append(reason)

        # --- country / geo ---
        ok, reason = check_country_allowed(country, config["allowed_countries"])
        if not ok:
            reasons.append(reason)

        # --- TAL match (checked BEFORE industry — a pre-vetted TAL account
        #     overrides the generic permitted-industries/exclusion rules below,
        #     since the account was already specifically approved).
        #     tal_required=False means the TAL is a PRIORITY list, not a gate —
        #     e.g. "prioritize TAL before targeting demographics" — so a lead not
        #     on the TAL isn't blocked here, it just falls through to the normal
        #     industry/exclusion checks below like any other lead. Defaults to
        #     True (hard gate) to match every campaign already built before this. ---
        on_tal = False
        tal_required = config.get("tal_required", True)
        if tal_df is not None:
            ok, reason, _tal_row = check_tal_match(domain, company, tal_df)
            on_tal = ok
            if not ok and tal_required:
                reasons.append(reason)

        # --- industry (skipped entirely if the account is already on the TAL) ---
        if not on_tal:
            ok, reason = check_industry_allowed(industry, config.get("permitted_industries"))
            if not ok:
                reasons.append(reason)

        # --- job function (fuzzy title-list match — informational note only, never blocks) ---
        matched_function = None
        function_match_note = None
        ok, function_match_note, matched_function = check_job_function(
            job_title, country, config["geo_function_rules"], job_titles_ref
        )
        if not ok:  # missing title, unmapped country, or genuinely no plausible function match (score<50)
            reasons.append(function_match_note)
            function_match_note = None

        # --- seniority floor (separate from function — e.g. "Manager or above") ---
        ok, reason = check_seniority(job_title, config.get("minimum_seniority"))
        if not ok:
            reasons.append(reason)

        # --- reference export match (Integrate/Leadscale typo check) ---
        ok, reason = check_reference_match(company, reference_set)
        if not ok:
            reasons.append(reason)

        # --- exclusions (industry exclusion skipped if already on the TAL) ---
        ok, reason = check_exclusions(company, job_title, industry, config["exclusions"], skip_industry_check=on_tal)
        if not ok:
            reasons.append(reason)

        # --- watch time ---
        ok, reason = check_watch_time(watch_time, config["watch_time_minutes_min"])
        if not ok:
            reasons.append(reason)

        # --- qualifying questions ---
        ok, reason = check_qualifying_questions(row, mapping, config.get("qualifying_questions", {}))
        if not ok:
            reasons.append(reason)

        # --- cap per company (checked last, after everything else has determined
        #     whether this lead is otherwise a pass) ---
        provisional_pass = len(reasons) == 0
        company_key = (domain or company or "").lower()
        if provisional_pass and company_key:
            company_counts[company_key] = company_counts.get(company_key, 0) + 1
            if company_counts[company_key] > config["cap_per_company"]:
                reasons.append(f"cap per company exceeded ({config['cap_per_company']} max)")

        # --- region quota (e.g. "43 leads max per region, so every regional
        #     marketer gets some") — also only counted against otherwise-passing leads ---
        matched_region = None
        provisional_pass = len(reasons) == 0  # re-check: cap-per-company may have just added a reason
        if provisional_pass:
            ok, reason, matched_region = check_region_quota(state, country, region_lookup, region_counts, target_per_region)
            if not ok:
                reasons.append(reason)

        li_url = append_linkedin_profile(first_name, last_name, company) if not reasons else None

        results.append({
            "first_name": first_name,
            "last_name": last_name,
            "email": email,
            "email_manual_check_note": email_manual_check_note,
            "phone": phone,
            "phone_region_note": phone_region_note,
            "company": company,
            "domain": domain,
            "industry": industry,
            "job_title": job_title,
            "matched_function": matched_function,
            "function_match_note": function_match_note,
            "country": country,
            "matched_region": matched_region,
            "on_tal": on_tal if tal_df is not None else None,
            "linkedin_url": li_url,
            "status": "PASS" if not reasons else "FAIL",
            "reasons": "; ".join(reasons) if reasons else "",
        })

    results_df = pd.DataFrame(results)
    passed_df = results_df[results_df["status"] == "PASS"].drop(columns=["status", "reasons"])
    failed_df = results_df[results_df["status"] == "FAIL"].copy()

    # Passed leads flagged catch-all — need the manual burner-email test before
    # being treated as fully confirmed, even though they weren't hard-rejected.
    manual_email_check_df = passed_df[passed_df["email_manual_check_note"].notna()]

    # Supplier-facing rejection report — clean category + detail, not internal debug strings.
    cat_and_flag = failed_df["reasons"].apply(categorize_rejection)
    failed_df["rejection_category"] = cat_and_flag.apply(lambda t: t[0])
    failed_df["needs_manual_qa"] = cat_and_flag.apply(lambda t: t[1])
    supplier_report_df = failed_df[[
        "first_name", "last_name", "email", "company", "domain",
        "rejection_category", "reasons",
    ]].rename(columns={"reasons": "rejection_detail"})

    summary_rows = [
        ("Lead target", config["campaign"].get("lead_target")),
        ("Leads in this file", len(results_df)),
        ("Passed this run", len(passed_df)),
        ("...of which need manual email check (catch-all)", len(manual_email_check_df)),
        ("Failed this run", len(failed_df)),
        ("Deadline", config["campaign"].get("term_end")),
        ("Note", "Cumulative delivered/remaining tracked manually by the team"),
    ]
    summary_df = pd.DataFrame(summary_rows, columns=["Metric", "Value"])

    fail_reason_counts = (
        failed_df["reasons"].str.split("; ").explode().value_counts()
        if not failed_df.empty else pd.Series(dtype=int)
    )
    reason_breakdown_df = fail_reason_counts.rename_axis("Reason").reset_index(name="Count")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        reason_breakdown_df.to_excel(writer, sheet_name="Summary", index=False, startrow=len(summary_df) + 2)
        passed_df.to_excel(writer, sheet_name="Passed - Client Ready", index=False)
        failed_df.to_excel(writer, sheet_name="Failed - Internal Detail", index=False)
        supplier_report_df.to_excel(writer, sheet_name="Supplier Rejection Report", index=False)
        manual_email_check_df.to_excel(writer, sheet_name="Needs Manual Email Check", index=False)

    print(f"Done. {len(passed_df)} passed, {len(failed_df)} failed. Output: {output_path}")
    return output_path


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--leads", required=True)
    ap.add_argument("--config", required=True)
    ap.add_argument("--titles", required=True)
    ap.add_argument("--tal-dir", default=".")
    ap.add_argument("--reference", default=None)
    ap.add_argument("--output", default="validated_output.xlsx")
    args = ap.parse_args()

    validate(args.leads, args.config, args.titles, args.tal_dir, args.reference, args.output)
