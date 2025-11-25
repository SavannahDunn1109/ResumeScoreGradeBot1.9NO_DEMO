import io
import os
import re
from datetime import date
import pandas as pd
import streamlit as st
from docx import Document
from PyPDF2 import PdfReader

# MSAL must be installed via requirements.txt in the app root.
try:
    import msal
except ModuleNotFoundError:
    st.error(
        "MSAL is not installed. Make sure `msal` is listed in requirements.txt at the app root, "
        "then Restart the app environment."
    )
    st.stop()

# Office365/SharePoint
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

# ======================== CONFIG ========================
SITE_URL = "https://eleven090.sharepoint.com/sites/Recruiting"
LIBRARY = "Shared Documents"
FOLDER = "Active Resumes"

st.set_page_config(page_title="Resume Scorer from SharePoint", layout="wide")
st.title("📄 Resume Scorer from SharePoint")

# ======================== AUTH HELPERS ========================
@st.cache_resource(show_spinner=False)
def connect_with_azure_app(site_url: str):
    """
    Force GUID-tenant authority via MSAL, then inject the token into SharePoint client.
    Requires: SharePoint → Application → Sites.Selected (Application) + site-level grant.
    """
    try:
        s = st.secrets["sharepoint_azure"]
        tenant_id = s["tenant_id"]     # MUST be the GUID
        client_id = s["client_id"]
        client_secret = s["client_secret"]
        site_url = s.get("site_url", site_url)

        # 🔎 Debug (safe): confirm we're not using eleven-09.com anywhere
        st.write({
            "tenant_id": tenant_id,
            "client_id": client_id[:8] + "...",
            "site_url": site_url
        })

        authority = f"https://login.microsoftonline.com/{tenant_id}"
        scopes = ["https://eleven090.sharepoint.com/.default"]

        app = msal.ConfidentialClientApplication(
            client_id=client_id,
            client_credential=client_secret,
            authority=authority,
        )
        token = app.acquire_token_for_client(scopes=scopes)
        assert "access_token" in token, f"MSAL error: {token}"

        ctx = ClientContext(site_url).with_access_token(token["access_token"])
        ctx.web.get().execute_query()  # sanity ping
        return ctx

    except KeyError:
        msg = """Missing secrets. Add to .streamlit/secrets.toml:
@st.cache_resource(show_spinner=False)
def connect_with_azure_app(site_url: str):
    """
    MSAL app-only auth using GUID tenant authority.
    Permissions: SharePoint -> Application -> Sites.Selected (with site-level grant).
    """
    try:
      
def _missing_sharepoint_secrets_error() -> None:
    msg = """
    import importlib

def _browser_cookie_available() -> bool:
    return importlib.util.find_spec("browser_cookie3") is not None

def _get_fedauth_rtfa():
    """Read FedAuth/rtFa from Chrome/Edge if browser_cookie3 is present."""
    
    @st.cache_resource(show_spinner=False)
def connect_with_azure_app(site_url: str):
    """
    MSAL app-only auth using GUID tenant authority.
    Permissions: SharePoint -> Application -> Sites.Selected (with site-level grant).
    """
    
        # Derive host from site_url to build correct resource scope
        parsed = urlparse(site_url)
        host = parsed.netloc

        authority = f"https://login.microsoftonline.com/{tenant_id}"
        scopes = [f"https://{host}/.default"]

        app = msal.ConfidentialClientApplication(
            client_id=client_id,
            client_credential=client_secret,
            authority=authority,
        )
        token = app.acquire_token_for_client(scopes=scopes)

        if "access_token" not in token:
            raise RuntimeError(f"MSAL error: {token}")

        ctx = ClientContext(site_url).with_access_token(token["access_token"])
        ctx.web.get().execute_query()  # sanity ping
        return ctx

    except KeyError:
        # secrets.toml is missing [sharepoint_azure]
        msg = (
            "Missing secrets for [sharepoint_azure] in .streamlit/secrets.toml.\n\n"
            "Add something like:\n"
            "[sharepoint_azure]\n"
            'tenant_id    = "YOUR_TENANT_GUID"\n'
            'client_id    = "YOUR_APP_CLIENT_ID"\n'
            'client_secret= "YOUR_CLIENT_SECRET_VALUE"\n'
            'site_url     = "https://eleven090.sharepoint.com/sites/Recruiting"\n'
        )
        raise RuntimeError(msg)

    except Exception as e:
        raise RuntimeError(f"Azure App auth failed: {e}")
...

[sharepoint_azure]
tenant_id    = "b7c46a1e-ef8c-4ba8-aeaf-0a29d31fb1be"   # GUID
client_id    = "090e3e87-bef3-45b7-b27c-57f5cee20845"   # your app id
client_secret= "YOUR_SECRET_VALUE"                      # VALUE column
site_url     = "https://eleven090.sharepoint.com/sites/Recruiting"
"""
    raise RuntimeError(msg)

import importlib

def _browser_cookie_available() -> bool:
    return importlib.util.find_spec("browser_cookie3") is not None

def _get_fedauth_rtfa():
    """Read FedAuth/rtFa from Chrome/Edge if browser_cookie3 is present."""
    if not _browser_cookie_available():
        raise RuntimeError(
            "Local (browser cookies) mode requires the 'browser-cookie3' package, "
            "which isn't available here. Use 'Azure App (client secret)' instead, "
            "or install it locally with: pip install browser-cookie3"
        )

    import browser_cookie3  # lazy import

    def pick(cj):
        fedauth = rtfa = None
        for c in cj:
            if c.domain.endswith("sharepoint.com"):
                n = c.name.lower()
                if n == "fedauth":
                    fedauth = c.value
                elif n == "rtfa":
                    rtfa = c.value
        return fedauth, rtfa

    # Try Chrome then Edge
    try:
        f, r = pick(browser_cookie3.chrome(domain_name=".sharepoint.com"))
        if f and r:
            return f, r
    except Exception:
        pass
    try:
        f, r = pick(browser_cookie3.edge(domain_name=".sharepoint.com"))
        if f and r:
            return f, r
    except Exception:
        pass
    return None, None


def connect_with_browser_cookies(site_url: str):
    # Use existing browser session (MFA already done). Local dev only.
    fedauth, rtfa = _get_fedauth_rtfa()
    if not (fedauth and rtfa):
        raise RuntimeError(
            "No SharePoint cookies found. Open the site in Chrome/Edge (non-incognito), "
            "sign in and complete MFA, then try again."
        )

    ctx = ClientContext(site_url)

    def _auth(req):
        req.set_header("Cookie", f"FedAuth={fedauth}; rtFa={rtfa}")

    # Monkey-patch request auth and sanity-check
    ctx.authentication_context._authenticate = _auth
    ctx.web.get().execute_query()
    return ctx

# ======================== FILE HELPERS ========================
def download_file(ctx, file_url):
    response = File.open_binary(ctx, file_url)
    return io.BytesIO(response.content)

def extract_text_from_pdf(file_bytes):
    text = ""
    reader = PdfReader(file_bytes)
    for page in reader.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + "\n"
    return text

def extract_text_from_docx(file_bytes):
    doc = Document(file_bytes)
    return "\n".join([p.text for p in doc.paragraphs])

# ======================== EXPERIENCE HELPERS ========================
MONTHS = {
    "jan": 1, "january": 1, "feb": 2, "february": 2, "mar": 3, "march": 3,
    "apr": 4, "april": 4, "may": 5, "jun": 6, "june": 6, "jul": 7, "july": 7,
    "aug": 8, "august": 8, "sep": 9, "sept": 9, "september": 9, "oct": 10, "october": 10,
    "nov": 11, "november": 11, "dec": 12, "december": 12,
}

def _mk_date(y: int, m: int) -> date:
    m = min(max(1, m), 12)
    return date(int(y), int(m), 15)

def _parse_month(token: str):
    if not token:
        return None
    return MONTHS.get(token.strip().lower())

def _parse_year(token: str):
    if not token:
        return None
    m = re.match(r"(19|20)\d{2}$", token.strip())
    return int(m.group(0)) if m else None

def _present_to_date() -> date:
    today = date.today()
    return date(today.year, today.month, 15)

def _extract_date_ranges(text: str):
    t = text.replace("\u2013", "-").replace("\u2014", "-")
    ranges = []

    pat_year_year = re.compile(
        r"\b(?P<y1>(?:19|20)\d{2})\s*[-to]+\s*(?P<y2>Present|Current|(?:19|20)\d{2})\b",
        flags=re.I
    )
    for m in pat_year_year.finditer(t):
        y1 = _parse_year(m.group("y1")); y2tok = m.group("y2")
        if not y1:
            continue
        start = _mk_date(y1, 6)
        if y2tok.lower() in ("present", "current"):
            end = _present_to_date()
        else:
            y2 = _parse_year(y2tok)
            if not y2:
                continue
            end = _mk_date(y2, 6)
        if end > start:
            ranges.append((start, end))

    if not ranges:
        return []
    ranges.sort(key=lambda r: r[0])
    merged = [ranges[0]]
    for s, e in ranges[1:]:
        last_s, last_e = merged[-1]
        if s <= last_e:
            merged[-1] = (last_s, max(last_e, e))
        else:
            merged.append((s, e))
    return merged

def _years_from_ranges(text: str) -> float:
    merged = _extract_date_ranges(text)
    total_months = 0
    for s, e in merged:
        diff = (e.year - s.year) * 12 + (e.month - s.month)
        total_months += max(0, diff)
    return round(total_months / 12.0, 1)

def _years_from_phrases(text: str) -> int:
    best = 0
    for m in re.finditer(r"\b([1-4]?\d)\s*\+?\s*[- ]?\s*(?:years?|yrs?)\b", text, flags=re.I):
        best = max(best, int(m.group(1)))
    return best

def estimate_years_experience(text: str):
    yrs_ranges = _years_from_ranges(text)
    yrs_phrases = _years_from_phrases(text)
    if yrs_ranges >= 0.5:
        return yrs_ranges, "ranges"
    return float(yrs_phrases), "phrases"

def classify_level(years: float, jr_max: int, mid_max: int) -> str:
    if years <= jr_max:
        return "Junior"
    elif years <= mid_max:
        return "Mid"
    else:
        return "Senior"

# ======================== REQUIREMENTS & SCORING ========================
uploaded_req_file = st.file_uploader("📄 Upload Requirements (.txt)", type=["txt"])
KEYWORDS = []
if uploaded_req_file:
    req_lines = uploaded_req_file.read().decode("utf-8").splitlines()
    for line in req_lines:
        line = line.strip()
        if line and not line.endswith(":"):
            KEYWORDS.append(line)
    st.success(f"✅ Loaded {len(KEYWORDS)} keywords.")
else:
    st.warning("⚠️ Please upload a requirements .txt file to begin scoring.")
    st.stop()

st.subheader("⚙️ Scoring & Filters")
exp_points_per_year = st.number_input("Points per year of experience", 0, 50, 5, 1)
jr_max = st.number_input("Max years for JUNIOR", 0, 10, 2, 1)
mid_max = st.number_input("Max years for MID", jr_max, 25, 6, 1)
enforce_min = st.checkbox("Enforce minimum years?", value=False)
min_years_required = st.number_input("Minimum years", 0, 30, 3, 1)

def score_resume(text: str):
    kw_score = 0
    found = []
    lower_text = text.lower()
    for kw in KEYWORDS:
        if kw.lower() in lower_text:
            kw_score += 10
            found.append(kw)
    years, src = estimate_years_experience(text)
    exp_score = years * exp_points_per_year
    total = kw_score + exp_score
    return {
        "years": years, "years_source": src,
        "level": classify_level(years, jr_max, mid_max),
        "kw_score": kw_score, "exp_score": exp_score,
        "total": total, "keywords_found": ", ".join(found),
    }

# ======================== MODE & CONNECTION ========================
st.sidebar.markdown("### Run mode")
mode = st.sidebar.radio(
    "Choose how to connect",
    (["Azure App (client secret)", "Demo (no SharePoint)"] +
     (["Local (browser cookies)"] if _browser_cookie_available() else [])),
    index=0,
)

ctx = None
if mode == "Azure App (client secret)":
    try:
        with st.spinner("Connecting to SharePoint with Azure App…"):
            ctx = connect_with_azure_app(SITE_URL)
        st.success("✅ Connected via Azure App")
    except Exception as e:
        st.error(str(e))
        st.stop()

elif mode == "Local (browser cookies)":
    if st.button("🔐 Connect using my browser session"):
        try:
            with st.spinner("Connecting via your browser session…"):
                ctx = connect_with_browser_cookies(SITE_URL)
            st.session_state.ctx = ctx
            st.success("✅ Connected with browser cookies")
        except Exception as e:
            st.error(f"❌ Connect failed: {e}")
    ctx = st.session_state.get("ctx")

else:
    st.info("🎬 Demo mode: Not connecting to SharePoint. Upload or test locally below.")

# ======================== MAIN: LIST, SCORE, EXPORT ========================
data = []

if ctx:
    try:
        folder_url = f"{LIBRARY}/{FOLDER}"
        folder = ctx.web.get_folder_by_server_relative_url(folder_url)
        files = folder.files
        ctx.load(files)
        ctx.execute_query()

        for file in files:
            filename = file.properties["Name"]
            if not filename.lower().endswith((".pdf", ".docx")):
                continue
            file_url = file.properties["ServerRelativeUrl"]
            file_bytes = download_file(ctx, file_url)
            text = extract_text_from_pdf(file_bytes) if filename.lower().endswith(".pdf") else extract_text_from_docx(file_bytes)
            result = score_resume(text)
            if enforce_min and result["years"] < float(min_years_required):
                continue
            data.append({
                "File Name": filename,
                "Est. Years": result["years"],
                "Level": result["level"],
                "Experience Source": result["years_source"],
                "Keyword Score": result["kw_score"],
                "Experience Score": result["exp_score"],
                "Total Score": result["total"],
                "Keywords Found": result["keywords_found"],
            })
    except Exception as e:
        st.error(f"Error reading SharePoint folder: {e}")

df = pd.DataFrame(data)
if not df.empty:
    df = df.sort_values(
        ["Level", "Est. Years", "Total Score"],
        ascending=[True, False, False]
    ).reset_index(drop=True)

st.dataframe(df, use_container_width=True)

if not df.empty:
    output = io.BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    st.download_button("📥 Download Excel Report", output, file_name="resume_scores.xlsx")

    if ctx and st.button("📤 Upload Excel to SharePoint"):
        try:
            output.seek(0)
            target_folder = ctx.web.get_folder_by_server_relative_url(LIBRARY)
            target_folder.upload_file("resume_scores.xlsx", output.getvalue())
            ctx.execute_query()
            st.success("Excel uploaded to SharePoint!")
        except Exception as e:
            st.error(f"Upload failed: {e}")
