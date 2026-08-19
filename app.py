"""
Prospect Drafter — web app.

Four steps: upload the list, pick which email in the sequence, read the drafts,
put them in Outlook. Nothing is ever sent.

Two workflows, chosen at the top of the page and configured under "workflows" in
config.json. Both read the same WLCC master report, on different sheets, and both work
the same way: which email is due is read off that sheet's Status column, and the Status
column is what gets written back.

    In-bound Leads  people who came to WLCC. Sheet "Inbound Leads", Status in column C.
    Cold Call       cold outreach. Sheet "Cold Database in Work", Status in column A,
                    and rows are picked by WLCC Fit Score first.

Nothing here is workflow-specific in code: the sheet, the columns, the templates and
whether there is a score to filter on all come from config.

Sequence stages, in both workflows:
    1  first email
    2  follow-up, ~7 days later, only if they didn't reply
    3  final follow-up, ~7 days after that

Deployed on Streamlit Community Cloud as a private app.
"""

from __future__ import annotations

import base64
import io
import json
import os
import re
import time
import urllib.parse
import zipfile
from collections import Counter
from datetime import date, datetime
from pathlib import Path

import streamlit as st

import prospect_drafter as pd_lib
import xlsx_patch

APP_DIR = Path(__file__).parent
CONFIG_PATH = APP_DIR / "config.json"

# Kept separate from pd_lib.GRAPH_SCOPES (mail-only) on purpose: OneDrive access is an
# independent, optional sign-in. If the Entra app registration doesn't have Files.ReadWrite
# added yet, only the OneDrive connect button fails — the working Outlook connection is
# never put at risk by bundling an unapproved scope into it.
ONEDRIVE_SCOPES = "offline_access Files.ReadWrite User.Read"

st.set_page_config(page_title="Prospect Drafter", page_icon="✉️", layout="centered")


# ---------------------------------------------------------------------------
# Config and secrets
# ---------------------------------------------------------------------------

@st.cache_data
def load_config() -> dict:
    return pd_lib.deep_merge(pd_lib.DEFAULT_CONFIG, json.loads(CONFIG_PATH.read_text("utf-8")))


def secret(name: str, default: str = "") -> str:
    try:
        if name in st.secrets:
            return str(st.secrets[name])
    except Exception:
        pass
    return os.environ.get(name, default)


def get_config(workflow: str | None = None) -> dict:
    cfg = json.loads(json.dumps(load_config()))
    if workflow:
        cfg = pd_lib.apply_workflow(cfg, workflow)
    if secret("GRAPH_CLIENT_ID"):
        cfg["outlook"]["graph_client_id"] = secret("GRAPH_CLIENT_ID")
    if secret("GRAPH_TENANT_ID"):
        cfg["outlook"]["graph_tenant_id"] = secret("GRAPH_TENANT_ID")
    for key in ("your_name", "your_company", "signature_html"):
        val = st.session_state.get(f"sender_{key}")
        if val:
            cfg["sender"][key] = val
    return cfg


@st.cache_data(ttl=300, show_spinner=False)
def cached_ai_key_check(key: str, model: str, base_url: str | None) -> dict:
    """Cached so the live check runs once every 5 minutes, not on every widget click.

    The key value is part of the cache's own lookup key, so fixing a broken key in Secrets
    and reloading the app is checked fresh immediately rather than reusing a stale failure.
    """
    return pd_lib.check_ai_key(key, model, base_url)


# ---------------------------------------------------------------------------
# HTML <-> plain text, so nobody has to edit HTML tags
# ---------------------------------------------------------------------------

def html_to_plain(html: str) -> str:
    text = re.sub(r"<br\s*/?>", "\n", html or "", flags=re.I)
    text = re.sub(r"</p\s*>", "\n\n", text, flags=re.I)
    text = re.sub(r"<[^>]+>", "", text)
    for a, b in (("&nbsp;", " "), ("&amp;", "&"), ("&lt;", "<"), ("&gt;", ">"),
                 ("&#39;", "'"), ("&quot;", '"')):
        text = text.replace(a, b)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def plain_to_html(text: str) -> str:
    blocks = [b.strip() for b in (text or "").split("\n\n") if b.strip()]
    return "\n".join("<p>" + pd_lib.escape_html(b) + "</p>" for b in blocks)


# ---------------------------------------------------------------------------
# Spreadsheet
# ---------------------------------------------------------------------------

def save_upload(uploaded) -> Path:
    tmp = APP_DIR / "_uploads"
    tmp.mkdir(exist_ok=True)
    path = tmp / uploaded.name
    path.write_bytes(uploaded.getbuffer())
    return path


def _open_sheet(source: Path, cfg: dict):
    from openpyxl import load_workbook, Workbook

    if source.suffix.lower() == ".csv":
        import csv
        wb = Workbook()
        ws = wb.active
        for r in csv.reader(source.open(encoding="utf-8-sig")):
            ws.append(r)
        return wb, ws

    wb = load_workbook(source)
    want = cfg["spreadsheet"].get("sheet_name")
    ws = wb[want] if want in wb.sheetnames else wb[wb.sheetnames[0]]
    return wb, ws


def spreadsheet_with_progress(source: Path, updates: dict[str, dict], cfg: dict) -> bytes:
    """Copy of the sheet with Touches and the contact dates brought up to date.

    `updates` maps lowercase email -> {"touches": int, "date": date}.
    Columns that don't exist yet are appended to the header row.
    """
    wb, ws = _open_sheet(source, cfg)
    colmap = cfg["spreadsheet"]["columns"]
    hdr_row = int(cfg["spreadsheet"].get("header_row", 1))

    headers: dict[str, int] = {}
    for cell in ws[hdr_row]:
        if cell.value is not None and str(cell.value).strip():
            headers[str(cell.value).strip().lower()] = cell.column

    def column_for(field: str, fallback_title: str) -> int | None:
        title = str(colmap.get(field) or fallback_title).strip()
        if not title:
            return None
        idx = headers.get(title.lower())
        if idx is None:
            idx = ws.max_column + 1
            ws.cell(row=hdr_row, column=idx).value = title
            headers[title.lower()] = idx
        return idx

    email_col = headers.get(str(colmap.get("email") or "Email").strip().lower())
    if not email_col:
        raise RuntimeError("Could not find the Email column to update.")

    touch_col = column_for("touches", "Touches")
    first_col = column_for("first_contact_date", "First Contact Date")
    last_col = column_for("last_contact_date", "Last Contact Date")

    for row in range(hdr_row + 1, ws.max_row + 1):
        v = ws.cell(row=row, column=email_col).value
        if not v:
            continue
        u = updates.get(str(v).strip().lower())
        if not u:
            continue
        when = u["date"].strftime("%Y-%m-%d")
        if touch_col:
            ws.cell(row=row, column=touch_col).value = u["touches"]
        if first_col and not ws.cell(row=row, column=first_col).value:
            ws.cell(row=row, column=first_col).value = when
        if last_col:
            ws.cell(row=row, column=last_col).value = when

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def workbook_sheets(source: Path) -> tuple[list[str], str | None]:
    """Visible sheet names, plus whichever sheet Excel had active when it was saved.

    Hidden sheets are left out: this workbook keeps its dropdown lists on a hidden
    "Sheet Info" tab, and offering that as somewhere to read prospects from is noise.
    """
    from openpyxl import load_workbook

    wb = load_workbook(source, read_only=True)
    try:
        visible = [ws.title for ws in wb.worksheets if ws.sheet_state == "visible"]
        try:
            active = wb.active.title if wb.active is not None else None
        except Exception:  # noqa: BLE001 - a stale active index shouldn't stop the upload
            active = None
        return (visible or list(wb.sheetnames)), active
    finally:
        wb.close()


def sheet_selector(source: Path, cfg: dict) -> str | None:
    """Show which sheet the app is about to read, and let it be changed.

    Each workflow names the sheet it expects, so the normal case is that this is already
    correct and the user just sees it confirmed. Returns None for CSV, which has no sheets.
    """
    if source.suffix.lower() == ".csv":
        st.caption("A CSV has just the one sheet.")
        return None

    try:
        names, excel_active = workbook_sheets(source)
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not open that workbook: {exc}")
        st.stop()

    want = str(cfg["spreadsheet"].get("sheet_name") or "").strip()
    label = cfg.get("workflow", {}).get("label", "this workflow")
    configured = pd_lib.find_sheet(names, want)
    default = pd_lib.resolve_sheet_name(names, want, excel_active)

    st.markdown("**Active sheet**")
    chosen = st.selectbox(
        "Active sheet",
        options=names,
        index=names.index(default),
        label_visibility="collapsed",
        format_func=lambda n: f"{n}  ·  active in Excel" if n == excel_active else n,
        help=(
            f"The {label} workflow reads “{want}”. Change this only if you keep that data "
            "on a differently named sheet."
            if want else "Which sheet holds the prospects."
        ),
    )

    if want and not configured:
        st.warning(
            f"This workbook has no **{want}** sheet, which is the one the **{label}** "
            f"workflow expects. Reading **{chosen}** instead — check that's right, or "
            "switch workflow."
        )
    elif configured and chosen != configured:
        st.warning(
            f"Reading **{chosen}**, not the **{configured}** sheet that **{label}** "
            "normally uses. The column names have to match, or nothing will be found."
        )
    else:
        st.caption(f"Reading **{chosen}** — the sheet the {label} workflow expects.")

    return chosen


# ---------------------------------------------------------------------------
# Cold Call: fit score filter and Status write-back
# ---------------------------------------------------------------------------

def fit_score_filter(prospects: list, cfg: dict) -> list:
    """Let the user pick which WLCC Fit Scores to work, and keep only those rows.

    This is step one of the Cold Call workflow: the score in column B is a human
    judgement about how well a company fits WLCC, and the whole point is to work the
    good ones first rather than the whole 800-row database.
    """
    spec = cfg["workflow"]["fit_score"]
    label = spec.get("label", "Fit score")
    present = pd_lib.fit_score_values(prospects, cfg)
    counts = Counter(p.fit_score for p in prospects)
    unscored = counts.get("", 0)

    if not present:
        st.error(
            f'No "{label}" values found in this sheet. Check the sheet and the '
            f"`fit_score` column name in the Cold Call config."
        )
        st.stop()

    st.markdown(f"**Which {label}s are you working?**")
    default = [s for s in spec.get("default", []) if s in present] or present[:1]
    chosen = st.multiselect(
        label, options=present, default=default,
        format_func=lambda s: f"{s} — {counts[s]} companies",
        label_visibility="collapsed",
        help="Pick one or more. 5 is the best fit. Only these rows go any further.",
    )
    if not chosen:
        st.warning(f"Pick at least one {label} to carry on.")
        st.stop()

    kept = [p for p in prospects if p.fit_score in set(chosen)]
    st.caption(
        f"**{len(kept)}** companies scored {', '.join(chosen)}."
        + (f" {unscored} rows have no score yet and are left out." if unscored else "")
    )
    if not kept:
        st.warning("No rows have those scores.")
        st.stop()
    return kept


def status_writeback_edits(drafts: list[dict], cfg: dict, today: date) -> dict[int, dict[str, object]]:
    """The cell edits that record this batch: Status, and the two contact dates.

    Keyed by worksheet row rather than by email address, because the cold database has
    duplicate and blank email cells and the row number is exact.
    """
    wb = cfg["workflow"]["writeback"]
    status_col = str(wb.get("status_column") or "").strip()
    first_col = str(wb.get("first_contact_column") or "").strip()
    last_col = str(wb.get("last_contact_column") or "").strip()

    edits: dict[int, dict[str, object]] = {}
    for d in drafts:
        cells: dict[str, object] = {}
        if status_col and d.get("next_status"):
            cells[status_col] = d["next_status"]
        # First Contact Date is written once and then left alone.
        if first_col and d.get("needs_first_date"):
            cells[first_col] = today
        if last_col:
            cells[last_col] = today
        if cells:
            edits[int(d["row"])] = cells
    return edits


def status_writeback(source: Path, drafts: list[dict], cfg: dict, today: date) -> bytes:
    """The workbook with this batch's Status and dates written in.

    Uses xlsx_patch rather than openpyxl on purpose. See that module's docstring: an
    openpyxl save of the WLCC master workbook drops the dropdowns and most of the
    conditional formatting on the other sheets.
    """
    if source.suffix.lower() not in (".xlsx", ".xlsm"):
        raise RuntimeError(
            "Writing the Status column back needs the Excel file itself (.xlsx), not a CSV."
        )
    edits = status_writeback_edits(drafts, cfg, today)
    if not edits:
        raise RuntimeError("Nothing to write back.")
    return xlsx_patch.patch_workbook_bytes(
        source.read_bytes(),
        cfg["spreadsheet"]["sheet_name"],
        edits,
        int(cfg["spreadsheet"].get("header_row", 1)),
    )


# ---------------------------------------------------------------------------
# Outlook
# ---------------------------------------------------------------------------

def outlook_web_link(to: str, subject: str, body_plain: str) -> str:
    q = urllib.parse.urlencode(
        {"to": to, "subject": subject, "body": body_plain}, quote_via=urllib.parse.quote
    )
    return f"https://outlook.office.com/mail/deeplink/compose?{q}"


def build_eml_zip(drafts: list[dict], cfg: dict) -> bytes:
    from email.message import EmailMessage

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for i, d in enumerate(drafts, 1):
            msg = EmailMessage()
            msg["Subject"] = d["subject"]
            msg["To"] = d["to"]
            msg["X-Unsent"] = "1"
            html = plain_to_html(d["body"]) + "\n" + cfg["sender"].get("signature_html", "")
            msg.set_content(d["body"])
            msg.add_alternative(f"<html><body>{html}</body></html>", subtype="html")
            safe = re.sub(r"[^A-Za-z0-9._-]+", "_", d["to"])[:60]
            z.writestr(f"{i:03d}_{safe}.eml", msg.as_bytes())
    return buf.getvalue()


def graph_has_reply(token: str, email: str, since: date) -> bool:
    """True if any message from this address arrived on or after `since`."""
    import requests

    address = email.replace("'", "''")          # OData escaping
    params = {
        "$filter": (
            f"from/emailAddress/address eq '{address}' "
            f"and receivedDateTime ge {since.isoformat()}T00:00:00Z"
        ),
        "$select": "id",
        "$top": "1",
    }
    r = requests.get(
        f"{pd_lib.GRAPH_BASE}/me/messages",
        headers={"Authorization": f"Bearer {token}"},
        params=params,
        timeout=30,
    )
    if r.status_code != 200:
        raise RuntimeError(f"Mailbox check failed ({r.status_code}): {r.text[:200]}")
    return bool(r.json().get("value"))


def device_code_start(cfg: dict, scope: str | None = None) -> dict:
    import requests
    r = requests.post(
        f"{pd_lib.LOGIN_BASE}/{cfg['outlook'].get('graph_tenant_id') or 'organizations'}"
        "/oauth2/v2.0/devicecode",
        data={"client_id": cfg["outlook"]["graph_client_id"], "scope": scope or pd_lib.GRAPH_SCOPES},
        timeout=30,
    )
    r.raise_for_status()
    return r.json()


def device_code_poll(cfg: dict, device_code: str, seconds: int = 150) -> dict | None:
    import requests
    tenant = cfg["outlook"].get("graph_tenant_id") or "organizations"
    deadline = time.time() + seconds
    interval = 5
    while time.time() < deadline:
        time.sleep(interval)
        r = requests.post(
            f"{pd_lib.LOGIN_BASE}/{tenant}/oauth2/v2.0/token",
            data={
                "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                "client_id": cfg["outlook"]["graph_client_id"],
                "device_code": device_code,
            },
            timeout=30,
        )
        body = r.json()
        if r.status_code == 200:
            return body
        err = body.get("error")
        if err == "authorization_pending":
            continue
        if err == "slow_down":
            interval += 5
            continue
        raise RuntimeError(body.get("error_description", err))
    return None


def encode_sharing_url(url: str) -> str:
    """Turn a normal OneDrive/SharePoint 'Copy link' URL into the token the /shares
    endpoint needs: unpadded base64url, prefixed with 'u!'."""
    b64 = base64.b64encode(url.strip().encode("utf-8")).decode("ascii")
    return "u!" + b64.rstrip("=").replace("/", "_").replace("+", "-")


def resolve_onedrive_link(token: str, url: str) -> dict:
    """Resolve a sharing link to a driveItem. Returns id, name, driveId, webUrl."""
    import requests
    encoded = encode_sharing_url(url)
    r = requests.get(
        f"{pd_lib.GRAPH_BASE}/shares/{encoded}/driveItem",
        headers={"Authorization": f"Bearer {token}", "Prefer": "redeemSharingLink"},
        params={"$select": "id,name,webUrl,size,parentReference"},
        timeout=30,
    )
    if r.status_code == 404:
        raise RuntimeError("That link doesn't seem to point to a file — check it's the right one.")
    if r.status_code == 401:
        raise RuntimeError("Sign-in has expired. Click Connect to OneDrive again.")
    if r.status_code != 200:
        raise RuntimeError(f"Could not open that link ({r.status_code}): {r.text[:200]}")
    item = r.json()
    drive_id = (item.get("parentReference") or {}).get("driveId")
    if not drive_id or not item.get("id"):
        raise RuntimeError("Microsoft didn't return enough information to identify that file.")
    return {"drive_id": drive_id, "item_id": item["id"], "name": item.get("name", "spreadsheet.xlsx")}


def onedrive_download(token: str, drive_id: str, item_id: str) -> bytes:
    import requests
    r = requests.get(
        f"{pd_lib.GRAPH_BASE}/drives/{drive_id}/items/{item_id}/content",
        headers={"Authorization": f"Bearer {token}"},
        timeout=60,
    )
    if r.status_code != 200:
        raise RuntimeError(f"Could not download the file ({r.status_code}): {r.text[:200]}")
    return r.content


def onedrive_upload(token: str, drive_id: str, item_id: str, data: bytes) -> None:
    import requests
    r = requests.put(
        f"{pd_lib.GRAPH_BASE}/drives/{drive_id}/items/{item_id}/content",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
        data=data,
        timeout=60,
    )
    if r.status_code == 423:
        raise RuntimeError(
            "The file is currently open in Excel, so OneDrive has it locked. Close it there "
            "and click this button again."
        )
    if r.status_code not in (200, 201):
        raise RuntimeError(f"Could not save back to OneDrive ({r.status_code}): {r.text[:200]}")


def sidebar_onedrive(cfg: dict) -> None:
    """Connect-to-OneDrive UI. Independent token from the Outlook one, on purpose —
    see the ONEDRIVE_SCOPES comment above."""
    if not cfg["outlook"].get("graph_client_id"):
        st.caption("Not available yet — needs the same one-off setup as Outlook.")
        return
    if st.session_state.get("onedrive_token"):
        st.success("Connected to OneDrive")
        return

    if st.button("Connect to OneDrive", use_container_width=True):
        try:
            st.session_state["onedrive_device_code"] = device_code_start(cfg, ONEDRIVE_SCOPES)
        except Exception as exc:  # noqa: BLE001
            st.error(f"Could not start sign-in: {exc}")
        st.rerun()

    dc = st.session_state.get("onedrive_device_code")
    if dc:
        st.markdown(f"**Code:** `{dc['user_code']}`")
        st.markdown(f"1. Open {dc['verification_uri']}\n2. Enter the code\n3. Sign in")
        with st.spinner("Waiting for sign-in..."):
            try:
                tok = device_code_poll(cfg, dc["device_code"])
            except Exception as exc:  # noqa: BLE001
                st.error(
                    f"Sign-in failed: {exc}\n\nIf this mentions the scope or permission not "
                    "being found, the OneDrive permission hasn't been added to the app "
                    "registration yet — ask Mo."
                )
                tok = None
        if tok:
            st.session_state["onedrive_token"] = tok["access_token"]
            st.session_state.pop("onedrive_device_code", None)
            st.rerun()
        else:
            st.warning("Timed out. Click Connect again.")


def sidebar_outlook(cfg: dict) -> None:
    """Connect-to-Outlook UI. Lives in the sidebar so both the reply check and the
    draft creation can use the same session."""
    if not cfg["outlook"].get("graph_client_id"):
        st.info("Outlook isn't connected yet — you'll use the one-at-a-time route.")
        return
    if st.session_state.get("graph_token"):
        st.success("Connected to Outlook")
        return

    if st.button("Connect to Outlook", use_container_width=True):
        try:
            st.session_state["device_code"] = device_code_start(cfg)
        except Exception as exc:  # noqa: BLE001
            st.error(f"Could not start sign-in: {exc}")
        st.rerun()

    dc = st.session_state.get("device_code")
    if dc:
        st.markdown(f"**Code:** `{dc['user_code']}`")
        st.markdown(f"1. Open {dc['verification_uri']}\n2. Enter the code\n3. Sign in")
        with st.spinner("Waiting for sign-in..."):
            try:
                tok = device_code_poll(cfg, dc["device_code"])
            except Exception as exc:  # noqa: BLE001
                st.error(f"Sign-in failed: {exc}")
                tok = None
        if tok:
            st.session_state["graph_token"] = tok["access_token"]
            st.session_state.pop("device_code", None)
            st.rerun()
        else:
            st.warning("Timed out. Click Connect again.")


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

def main() -> None:
    st.title("✉️ Prospect Drafter")
    st.caption(
        "Reads each prospect's website and writes a personalised email. **Nothing is ever "
        "sent** — you read every draft, then it goes into Outlook for you to send yourself."
    )

    # ---- AI key check, right up front --------------------------------------
    # A live check, not a guess from the key's shape: this app once flagged a real, working
    # key as broken purely because of its prefix. Only asking Google settles it, and it's
    # worth doing before anything else because a broken key means every draft below silently
    # degrades to the plain [FILL THIS IN] template - better the user hears that now than
    # discovers it three drafts in.
    base = get_config()
    ai_key = secret(base["ai"]["api_key_env"])
    key_check = cached_ai_key_check(ai_key, base["ai"]["model"], base["ai"].get("base_url"))
    if key_check["ok"] is False:
        st.error(
            f"**AI key problem** — {key_check['message']}\n\n"
            "The app still works without it: every draft will fall back to the plain "
            "template with `[FILL THIS IN]` gaps for you to complete by hand, instead of a "
            "written email."
        )
    elif key_check["ok"] is None:
        st.warning(f"Could not confirm the AI key is working — {key_check['message']}")

    # ---- Which workflow ------------------------------------------------------
    names = pd_lib.workflow_names(base)

    if len(names) > 1:
        workflow = st.radio(
            "Workflow",
            options=names,
            format_func=lambda n: pd_lib.workflow_label(base, n),
            horizontal=True,
            key="workflow",
        )
    else:
        workflow = names[0] if names else None

    cfg = get_config(workflow)
    wf = cfg.get("workflow", {})
    # Branch on what a workflow *does*, not on which one it is. Both workflows are now
    # status-driven against the same workbook, and the only differences left between them
    # live in config: which sheet, which columns, whether there is a score to filter on.
    seq_gate = str(cfg["sequence"].get("gate", "touches")).lower()
    uses_status = seq_gate == "status"
    has_fit_score = bool(wf.get("fit_score"))
    patch_writeback = str((wf.get("writeback") or {}).get("mode", "")).lower() == "patch"
    if wf.get("blurb"):
        st.caption(wf["blurb"])

    # Drafts belong to the workflow they were written for. Switching workflow clears them so
    # a batch can never be written back through the other workflow's settings. This matters
    # more now that both workflows patch the same workbook: they use different sheets and
    # different Status columns (C vs A), so a mismatched write-back would put a status into
    # the wrong column of the wrong sheet.
    if st.session_state.get("drafted_workflow") not in (None, workflow):
        for key in ("drafts", "stage", "source_path", "replied"):
            st.session_state.pop(key, None)
        st.session_state.pop("drafted_workflow", None)

    seq = cfg["sequence"]
    labels = {int(k): v for k, v in seq["labels"].items()}
    wait_days = int(seq.get("wait_days", 7))

    with st.sidebar:
        st.subheader("Your details")
        st.text_input("Your name", value=cfg["sender"]["your_name"], key="sender_your_name")
        st.text_input("Your company", value=cfg["sender"]["your_company"],
                      key="sender_your_company")
        st.text_area("Email signature", value=cfg["sender"].get("signature_html", ""),
                     key="sender_signature_html", height=100,
                     help="Added to the bottom of every draft.")
        st.divider()
        st.subheader("Outlook")
        sidebar_outlook(cfg)
        st.divider()
        st.subheader("OneDrive")
        st.caption("Optional — lets you load and save the spreadsheet without downloading it.")
        sidebar_onedrive(cfg)
        st.divider()
        if key_check["ok"] is True:
            st.caption("AI key active ✓")
        elif key_check["ok"] is False:
            st.error(key_check["message"])
        else:
            st.caption(f"AI key configured, unverified — {key_check['message']}")

    # ---- Step 1: the list ---------------------------------------------------
    st.header("1. Your prospect list")

    source_choice = st.radio(
        "Where's your spreadsheet?", ["Upload a file", "Load from OneDrive"],
        horizontal=True, label_visibility="collapsed",
    )

    path = None

    if source_choice == "Upload a file":
        st.session_state.pop("onedrive_source", None)
        uploaded = st.file_uploader(
            "Upload your spreadsheet", type=["xlsx", "xlsm", "csv"],
            help="The file you keep your prospects in. It is never modified.",
        )
        if not uploaded:
            st.info("Upload a spreadsheet to begin.")
            st.stop()
        path = save_upload(uploaded)

    else:
        if not st.session_state.get("onedrive_token"):
            st.info("Connect to OneDrive in the sidebar first.")
            st.stop()

        link = st.text_input(
            "Paste the OneDrive link to your spreadsheet",
            help="In OneDrive or Excel Online: right-click the file → Copy link.",
        )
        loaded = st.session_state.get("onedrive_source")

        if st.button("Load from OneDrive", use_container_width=True, disabled=not link.strip()):
            try:
                info = resolve_onedrive_link(st.session_state["onedrive_token"], link.strip())
                data = onedrive_download(
                    st.session_state["onedrive_token"], info["drive_id"], info["item_id"]
                )
                dest = APP_DIR / "_uploads"
                dest.mkdir(exist_ok=True)
                local_path = dest / info["name"]
                local_path.write_bytes(data)
                st.session_state["onedrive_source"] = {**info, "local_path": str(local_path)}
                st.success(f"Loaded **{info['name']}**.")
                st.rerun()
            except Exception as exc:  # noqa: BLE001
                st.error(str(exc))

        if not loaded:
            st.stop()
        st.caption(f"Using **{loaded['name']}** from OneDrive.")
        path = Path(loaded["local_path"])

    cfg["spreadsheet"]["path"] = str(path)

    chosen_sheet = sheet_selector(path, cfg)
    if chosen_sheet:
        # Everything downstream reads this, including the Cold Call write-back, so the
        # sheet that gets patched is always the sheet that was read.
        cfg["spreadsheet"]["sheet_name"] = chosen_sheet

    try:
        prospects = pd_lib.read_prospects(cfg)
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not read that file: {exc}")
        st.stop()

    st.write(f"**{len(prospects)}** rows read.")

    # ---- Pick the fit scores to work, where the sheet has them ---------------
    if has_fit_score:
        prospects = fit_score_filter(prospects, cfg)

    # The number of stages is a config choice, not a fixed 3 - Cold Call now runs a fourth,
    # final follow-up that Internal lead doesn't have. Everything below reads it from
    # `labels` rather than assuming a count, so a workflow can add or drop a stage in
    # config.json alone.
    stages = sorted(labels)

    today = date.today()
    buckets: dict[int, list] = {}
    reasons: dict[int, list[tuple]] = {}
    for stage in stages:
        ok, no = [], []
        for p in prospects:
            good, why = pd_lib.eligible_for_stage(p, stage, cfg, today)
            (ok if good else no).append(p if good else (p, why))
        buckets[stage] = ok
        reasons[stage] = no

    for col, stage in zip(st.columns(len(stages)), stages):
        col.metric(labels[stage], len(buckets[stage]))

    # ---- Step 2: which email ------------------------------------------------
    st.header("2. Which email are you sending?")

    stage = st.radio(
        "Stage",
        options=stages,
        format_func=lambda s: f"{labels[s]} — {len(buckets[s])} ready",
        horizontal=True,
        label_visibility="collapsed",
    )

    if uses_status:
        flow = (seq.get("status_flow") or {}).get(str(stage), {})
        needs = ", ".join(f"`{x}`" if x else "blank" for x in flow.get("from", [])) or "—"
        sets = pd_lib.next_status(cfg, stage)
        st.caption(
            f"Rows whose **Status** is {needs}"
            + (f", last contacted at least {wait_days} days ago" if stage > 1 else "")
            + (f". Once drafted, their Status becomes `{sets}`." if sets else ".")
        )
    elif stage == 1:
        st.caption("People who haven't been emailed yet.")
    else:
        st.caption(
            f"People who received email {stage - 1}, haven't replied, and were last "
            f"contacted at least {wait_days} days ago."
        )

    ready = buckets[stage]

    if not ready:
        st.warning(f"Nobody is ready for the {labels[stage].lower()} right now.")
        with st.expander("Why not?"):
            for p, why in reasons[stage][:40]:
                st.write(f"- **{p.company or p.email}** — {why}")
        st.stop()

    # Optional mailbox check before drafting a follow-up
    if stage > 1:
        if st.session_state.get("graph_token"):
            if st.button("Check who has replied first", use_container_width=True):
                excluded = []
                bar = st.progress(0.0, text="Checking your mailbox...")
                for i, p in enumerate(list(ready), 1):
                    since = p.latest_contact or today
                    try:
                        if graph_has_reply(st.session_state["graph_token"], p.email, since):
                            excluded.append(p.email.lower())
                    except Exception as exc:  # noqa: BLE001
                        st.warning(f"Could not check {p.email}: {exc}")
                        break
                    bar.progress(i / len(ready))
                bar.empty()
                st.session_state["replied"] = excluded
                st.success(
                    f"{len(excluded)} of these have replied — removed from the list."
                    if excluded else "Nobody on this list has replied."
                )
        else:
            st.caption(
                "Tip: connect Outlook in the sidebar and the app can check who has already "
                "replied. Until then, mark them `Replied` in the Status column yourself."
            )

    replied = set(st.session_state.get("replied") or [])
    if replied:
        ready = [p for p in ready if p.email.lower() not in replied]
        st.info(f"{len(replied)} prospect(s) excluded because they replied.")
        if not ready:
            st.stop()

    no_site = [p for p in ready if not p.website]
    if no_site:
        st.caption(
            f"{len(no_site)} of these have no website. They'll get the generic version of "
            "the template — nothing will be invented about them."
        )

    with st.expander("Check the columns were read correctly"):
        p0 = ready[0]
        st.write({
            "First name": p0.first_name, "Last name": p0.last_name, "Company": p0.company,
            "Job title": p0.job_title or "(blank)", "Email": p0.email,
            "Website": p0.website or "(blank)",
            "Emails sent so far": p0.touches,
            "Last contact": str(p0.latest_contact) if p0.latest_contact else "(none)",
            "Extra context": p0.context or "(none)",
        })

    # ---- Step 3: draft ------------------------------------------------------
    st.header("3. Write the drafts")

    template_path = pd_lib.stage_template_path(cfg, stage, APP_DIR)
    if not template_path.exists():
        st.error(f"The template for this stage is missing ({template_path.name}). Tell Mo.")
        st.stop()

    how_many = st.number_input(
        "How many prospects?", min_value=1, max_value=len(ready),
        value=min(3, len(ready)),
        help="Start small so you can check the tone before doing the whole list.",
    )

    called_first = False
    if stage == 1:
        called_first = st.checkbox(
            "I tried phoning these people first",
            value=bool(
                wf.get("called_first_default") or cfg["sender"].get("mention_prior_call")
            ),
            help=(
                "The first email can mention that you tried to call. Only tick this if you "
                "actually did — otherwise the sentence is left out."
            ),
        )
        if called_first:
            st.caption(
                "The email will say you tried to call and couldn't get through. Make sure "
                "that's true for everyone in this batch."
            )

    if st.button(f"Write the {labels[stage].lower()} drafts", type="primary",
                 use_container_width=True):
        queue = ready[: int(how_many)]
        template = template_path.read_text("utf-8")
        system_prompt = pd_lib.build_system_prompt(cfg, template, stage, called_first)
        os.environ["GEMINI_API_KEY"] = secret("GEMINI_API_KEY")

        # Some stages send fixed, approved copy rather than an AI-adapted email — see
        # is_verbatim_stage(). There's nothing for research to feed into on those stages, so
        # skip the website fetch, and the "researched" flag would otherwise wrongly mark every
        # one of these as needing a look even though nothing went wrong.
        verbatim = pd_lib.is_verbatim_stage(cfg, stage)
        fixed_subject = cfg["sequence"].get("fixed_subject")

        drafts = []
        bar = st.progress(0.0, text="Starting...")
        for i, p in enumerate(queue, 1):
            bar.progress((i - 1) / len(queue), text=f"{p.full_name or p.email} — {p.company}")
            research, warns = ("", [])
            if p.website and not verbatim:
                research, warns = pd_lib.fetch_site_text(p.website, cfg)
            try:
                out = pd_lib.call_gemini(cfg, system_prompt, pd_lib.build_user_prompt(p, research))
                ai_ok = True
            except Exception as exc:  # noqa: BLE001
                warns.append(f"AI unavailable: {exc}")
                out = pd_lib.fallback_fill(template, p, cfg)
                ai_ok = False
            drafts.append({
                "to": p.email, "name": p.full_name or p.email, "company": p.company,
                "subject": fixed_subject or out["subject"],
                "body": html_to_plain(out["body_html"]),
                "note": out["personalisation_note"],
                "researched": True if verbatim else len(research) >= 200,
                "ai_ok": ai_ok, "warnings": warns, "approved": True,
                "touches_after": p.touches + 1,
                # For the Cold Call Status write-back.
                "row": p.row,
                "next_status": pd_lib.next_status(cfg, stage),
                "needs_first_date": p.first_contact is None,
            })
            if i < len(queue):
                time.sleep(float(cfg["ai"]["delay_seconds"]))
        bar.progress(1.0, text="Done")
        st.session_state["drafts"] = drafts
        st.session_state["stage"] = stage
        st.session_state["source_path"] = str(path)
        st.session_state["drafted_workflow"] = workflow

    drafts = st.session_state.get("drafts")
    if not drafts:
        st.stop()

    drafted_stage = st.session_state.get("stage", stage)
    if drafted_stage != stage:
        st.warning(
            f"The drafts below are the **{labels[drafted_stage].lower()}**. Click the button "
            "above to write drafts for the stage you just selected."
        )

    # The sign-off follows the stage where the workflow defines one, because approved copy does
    # not sign off the same way throughout a sequence. Anything typed into the sidebar wins.
    typed = (st.session_state.get("sender_signature_html") or "").strip()
    if not typed or typed == (load_config()["sender"].get("signature_html") or "").strip():
        cfg["sender"]["signature_html"] = pd_lib.stage_signature(cfg, drafted_stage)

    # ---- Step 4: review -----------------------------------------------------
    st.header("4. Read them")
    weak = [d for d in drafts if not d["researched"] or not d["ai_ok"]]
    if weak:
        st.warning(
            f"{len(weak)} draft(s) marked **needs a look** — the website couldn't be read, so "
            "they use the plain template. Read those first."
        )
    st.caption("Untick anything you don't want. Edit the wording freely — it's just text.")

    for i, d in enumerate(drafts):
        flag = "" if (d["researched"] and d["ai_ok"]) else "  ⚠️ needs a look"
        with st.expander(f"{d['name']} — {d['company']}{flag}", expanded=(i == 0)):
            d["approved"] = st.checkbox("Include this one", value=d["approved"], key=f"ap{i}")
            st.caption(f"To: {d['to']}  ·  Angle: {d['note'] or '—'}")
            d["subject"] = st.text_input("Subject", value=d["subject"], key=f"su{i}")
            d["body"] = st.text_area("Message", value=d["body"], height=220, key=f"bo{i}")
            if d["warnings"]:
                st.caption("Notes: " + "; ".join(d["warnings"]))

    approved = [d for d in drafts if d["approved"]]
    st.info(f"**{len(approved)} of {len(drafts)}** ready to go into Outlook.")
    if not approved:
        st.stop()

    # ---- Step 5: deliver ----------------------------------------------------
    st.header("5. Put them in Outlook")
    tabs = st.tabs(["Straight into Outlook", "One at a time", "Download files"])

    with tabs[0]:
        if not cfg["outlook"].get("graph_client_id"):
            st.info(
                "Not switched on yet — it needs a one-off setup by IT. Until then use "
                "**One at a time**, which works right now."
            )
        elif not st.session_state.get("graph_token"):
            st.info("Connect to Outlook in the sidebar first.")
        elif st.button(f"Create {len(approved)} drafts in Outlook", type="primary",
                       use_container_width=True):
            done, failed = [], []
            bar = st.progress(0.0)
            for i, d in enumerate(approved, 1):
                entry = {"to": d["to"], "subject": d["subject"],
                         "body_html": plain_to_html(d["body"])}
                try:
                    pd_lib.push_graph(entry, cfg, st.session_state["graph_token"])
                    done.append(d)
                except Exception as exc:  # noqa: BLE001
                    failed.append((d, str(exc)))
                bar.progress(i / len(approved))
            if done:
                st.success(f"{len(done)} draft(s) created. Open Outlook and look in **Drafts**.")
            for d, err in failed:
                st.error(f"{d['to']}: {err}")

    with tabs[1]:
        st.caption(
            "Click a name to open a new Outlook message with everything filled in. Check it, "
            "then press Send. Works with Outlook in the browser."
        )
        for d in approved:
            body = d["body"] + "\n\n" + html_to_plain(cfg["sender"].get("signature_html", ""))
            link = outlook_web_link(d["to"], d["subject"], body)
            if len(link) > 1900:
                st.markdown(f"**{d['name']}** — too long for a one-click link, copy this:")
                st.code(f"To: {d['to']}\nSubject: {d['subject']}\n\n{body}", language=None)
            else:
                st.link_button(f"✉️ {d['name']} — {d['company']}", link, use_container_width=True)

    with tabs[2]:
        st.caption("Email files you can keep or forward. Not importable into Outlook on the web.")
        st.download_button(
            "Download the drafts as .eml files", data=build_eml_zip(approved, cfg),
            file_name=f"drafts_stage{drafted_stage}.zip", mime="application/zip",
            use_container_width=True,
        )

    # ---- Bookkeeping --------------------------------------------------------
    st.divider()
    st.subheader("Keep your spreadsheet up to date")
    st.caption(
        "This records who you just handled and when, which is what makes the follow-ups "
        "work later."
    )
    onedrive_source = st.session_state.get("onedrive_source")
    source_path = Path(st.session_state["source_path"])

    if patch_writeback:
        next_stat = pd_lib.next_status(cfg, drafted_stage)
        st.write(
            f"Sets **Status** to `{next_stat}` and updates the contact dates for the "
            f"**{len(approved)}** rows below. Everything else in the workbook is left "
            "exactly as it is."
            if next_stat else
            f"Updates the contact dates for the **{len(approved)}** rows below."
        )
        with st.expander("Which rows change"):
            for d in approved:
                bits = [f"Status → `{next_stat}`"] if next_stat else []
                if d.get("needs_first_date"):
                    bits.append("First Contact Date set")
                bits.append(f"Last Contact Date → {today:%Y-%m-%d}")
                st.write(f"- Row {d['row']} · **{d['company'] or d['to']}** — " + ", ".join(bits))
        try:
            updated_bytes = status_writeback(source_path, approved, cfg, today)
        except Exception as exc:  # noqa: BLE001
            updated_bytes = None
            st.error(f"Could not prepare the update: {exc}")
    else:
        updates = {
            d["to"].lower(): {"touches": d["touches_after"], "date": today} for d in approved
        }
        try:
            updated_bytes = spreadsheet_with_progress(source_path, updates, cfg)
        except Exception as exc:  # noqa: BLE001
            updated_bytes = None
            st.caption(f"(Could not build the updated spreadsheet: {exc})")

    if updated_bytes is not None and onedrive_source:
        if st.button("Save back to OneDrive", type="primary", use_container_width=True):
            try:
                onedrive_upload(
                    st.session_state["onedrive_token"],
                    onedrive_source["drive_id"], onedrive_source["item_id"], updated_bytes,
                )
                st.success(f"Saved back to **{onedrive_source['name']}** in OneDrive.")
            except Exception as exc:  # noqa: BLE001
                st.error(str(exc))
        st.caption("Or keep a local copy too:")

    if updated_bytes is not None:
        stem = source_path.stem if patch_writeback else "prospects"
        st.download_button(
            "Download updated spreadsheet",
            data=updated_bytes,
            file_name=f"{stem}_updated_{today:%Y-%m-%d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


if __name__ == "__main__":
    main()
