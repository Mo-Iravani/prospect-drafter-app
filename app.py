"""
Prospect Drafter — web app.

Four steps: upload the list, pick which email in the sequence, read the drafts,
put them in Outlook. Nothing is ever sent.

Sequence stages:
    1  first email
    2  follow-up, ~7 days later, only if they didn't reply
    3  final follow-up, ~7 days after that

Deployed on Streamlit Community Cloud as a private app.
"""

from __future__ import annotations

import io
import json
import os
import re
import time
import urllib.parse
import zipfile
from datetime import date, datetime
from pathlib import Path

import streamlit as st

import prospect_drafter as pd_lib

APP_DIR = Path(__file__).parent
CONFIG_PATH = APP_DIR / "config.json"

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


def get_config() -> dict:
    cfg = json.loads(json.dumps(load_config()))
    if secret("GRAPH_CLIENT_ID"):
        cfg["outlook"]["graph_client_id"] = secret("GRAPH_CLIENT_ID")
    if secret("GRAPH_TENANT_ID"):
        cfg["outlook"]["graph_tenant_id"] = secret("GRAPH_TENANT_ID")
    for key in ("your_name", "your_company", "signature_html"):
        val = st.session_state.get(f"sender_{key}")
        if val:
            cfg["sender"][key] = val
    return cfg


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


def device_code_start(cfg: dict) -> dict:
    import requests
    r = requests.post(
        f"{pd_lib.LOGIN_BASE}/{cfg['outlook'].get('graph_tenant_id') or 'organizations'}"
        "/oauth2/v2.0/devicecode",
        data={"client_id": cfg["outlook"]["graph_client_id"], "scope": pd_lib.GRAPH_SCOPES},
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
    cfg = get_config()
    seq = cfg["sequence"]
    labels = {int(k): v for k, v in seq["labels"].items()}
    wait_days = int(seq.get("wait_days", 7))

    st.title("✉️ Prospect Drafter")
    st.caption(
        "Reads each prospect's website and writes a personalised email. **Nothing is ever "
        "sent** — you read every draft, then it goes into Outlook for you to send yourself."
    )

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
        if secret("GEMINI_API_KEY"):
            st.caption("AI key configured ✓")
        else:
            st.error("No AI key. Ask Mo to add GEMINI_API_KEY in the app settings.")

    # ---- Step 1: the list ---------------------------------------------------
    st.header("1. Your prospect list")

    uploaded = st.file_uploader(
        "Upload your spreadsheet", type=["xlsx", "xlsm", "csv"],
        help="The file you keep your prospects in. It is never modified.",
    )
    if not uploaded:
        st.info("Upload a spreadsheet to begin.")
        st.stop()

    path = save_upload(uploaded)
    cfg["spreadsheet"]["path"] = str(path)

    try:
        prospects = pd_lib.read_prospects(cfg)
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not read that file: {exc}")
        st.stop()

    today = date.today()
    buckets: dict[int, list] = {}
    reasons: dict[int, list[tuple]] = {}
    for stage in (1, 2, 3):
        ok, no = [], []
        for p in prospects:
            good, why = pd_lib.eligible_for_stage(p, stage, cfg, today)
            (ok if good else no).append(p if good else (p, why))
        buckets[stage] = ok
        reasons[stage] = no

    st.write(f"**{len(prospects)}** rows read.")
    c1, c2, c3 = st.columns(3)
    for col, stage in zip((c1, c2, c3), (1, 2, 3)):
        col.metric(labels[stage], len(buckets[stage]))

    # ---- Step 2: which email ------------------------------------------------
    st.header("2. Which email are you sending?")

    stage = st.radio(
        "Stage",
        options=[1, 2, 3],
        format_func=lambda s: f"{labels[s]} — {len(buckets[s])} ready",
        horizontal=True,
        label_visibility="collapsed",
    )

    if stage == 1:
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
            value=bool(cfg["sender"].get("mention_prior_call")),
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

        drafts = []
        bar = st.progress(0.0, text="Starting...")
        for i, p in enumerate(queue, 1):
            bar.progress((i - 1) / len(queue), text=f"{p.full_name or p.email} — {p.company}")
            research, warns = ("", [])
            if p.website:
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
                "subject": out["subject"], "body": html_to_plain(out["body_html"]),
                "note": out["personalisation_note"], "researched": len(research) >= 200,
                "ai_ok": ai_ok, "warnings": warns, "approved": True,
                "touches_after": p.touches + 1,
            })
            if i < len(queue):
                time.sleep(float(cfg["ai"]["delay_seconds"]))
        bar.progress(1.0, text="Done")
        st.session_state["drafts"] = drafts
        st.session_state["stage"] = stage
        st.session_state["source_path"] = str(path)

    drafts = st.session_state.get("drafts")
    if not drafts:
        st.stop()

    drafted_stage = st.session_state.get("stage", stage)
    if drafted_stage != stage:
        st.warning(
            f"The drafts below are the **{labels[drafted_stage].lower()}**. Click the button "
            "above to write drafts for the stage you just selected."
        )

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
        "Download this and save it over your own copy. It records who you just handled and "
        "when, which is what makes the follow-ups work later."
    )
    updates = {d["to"].lower(): {"touches": d["touches_after"], "date": today} for d in approved}
    try:
        st.download_button(
            "Download updated spreadsheet",
            data=spreadsheet_with_progress(Path(st.session_state["source_path"]), updates, cfg),
            file_name=f"prospects_updated_{today:%Y-%m-%d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception as exc:  # noqa: BLE001
        st.caption(f"(Could not build the updated spreadsheet: {exc})")


if __name__ == "__main__":
    main()
