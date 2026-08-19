# Deploying the app — for Mo

One-time setup, about 40 minutes. After this she gets a URL, signs in with her work email,
and never touches Python, Colab or a terminal again.

---

## What you're building

A private web app on **Streamlit Community Cloud** (free). Private means genuinely private:
you invite viewers by email address, and everyone else gets nothing. That's the thing Netlify's
free tier couldn't do.

Two constraints worth knowing before you start:

- **You get one private app at a time** on the free tier. If you later want a second private
  app, this one has to go public or move to a paid plan.
- **Community Cloud is hosted in the US.** Prospect data passes through it. For UK GDPR purposes
  that's a transfer to consider — the data is transient (nothing is stored between sessions),
  but it's worth a note in your records.

---

## Running it on your own machine

Double-click **`run-local.cmd`** in the project folder — the one above `prospect-drafter-app`.
It opens the app at http://localhost:8501. Leave the black window open while you work; closing
it stops the app.

The first run builds a virtual environment in `prospect-drafter-app\.venv` and installs the
requirements, which takes a minute. After that, start-up is immediate: it checks the
environment by importing the packages rather than reinstalling them.

It also checks `.streamlit\secrets.toml` before starting, and warns you if the
`GEMINI_API_KEY` doesn't look like a Gemini key. Those begin with `AIza`. Without a working
key the app still reads websites and produces drafts, but they come out as the plain template
with `[FILL THIS IN]` gaps rather than written emails. Note that the app's own sidebar only
checks the key is *present*, not that it works — so trust the launcher's warning over the
"AI key configured ✓" line.

If the environment ever gets into a bad state, delete `prospect-drafter-app\.venv` and run the
launcher again. It rebuilds from scratch. One thing worth knowing: a folder in `.venv` with no
`__init__.py` imports as an empty namespace package, so a half-copied environment can pass a
naive `import streamlit` check while having nothing in it. That's why the launcher imports
`streamlit.web.cli` — the module that actually runs the app — instead.

To run it from a terminal instead:

```bash
cd prospect-drafter-app && .venv/Scripts/python.exe -m streamlit run app.py
```

Local runs read the same `config.json` and templates as the deployed app, so this is the right
place to try a template change before pushing it.

---

## Step 1 — Put the code on GitHub

1. Create a GitHub account if you don't have one: **https://github.com/signup**
2. Create a new repository — **https://github.com/new**
   - Name: `prospect-drafter`
   - **Private** (Streamlit Community Cloud can deploy from private repos)
   - Don't add a README, .gitignore or licence
3. On the empty repo page, click **uploading an existing file**
4. Drag in everything from this folder **except** `.streamlit/secrets.toml` if you ever create one:

   ```
   app.py
   prospect_drafter.py
   xlsx_patch.py
   config.json
   inbound_lead_first_contact.md
   inbound_lead_first_followup.md
   inbound_lead_second_followup.md
   coldcall_first_contact.md
   coldcall_first_followup.md
   coldcall_second_followup.md
   coldcall_third_followup.md
   requirements.txt
   sample_prospects.xlsx
   .gitignore
   .streamlit/config.toml
   ```

   `generic-templates/` is reference material only (the pre-refinement, AI-adapted versions
   of these templates) — it doesn't need to be deployed, but there's no harm including it.

5. Click **Commit changes**

> The `.gitignore` already excludes `secrets.toml` and `_uploads/`. Keys go in Streamlit's
> settings, never in the repo.

> **Watch out for GitHub's drag-and-drop uploader with dotfiles.** It can silently mishandle
> `.gitignore` and `.streamlit/config.toml` — turning `.gitignore` into a plain file literally
> named `download`, and flattening `.streamlit/config.toml` into a root-level `config.toml`.
> Both look like they uploaded fine but land in the wrong place and do nothing. If you already
> have the repo on your machine, it's more reliable to use a real git client instead — either
> the command line (`git add .gitignore .streamlit/config.toml && git commit -m "..." && git
> push`) or **GitHub Desktop** (free, https://desktop.github.com) — rather than the website's
> upload button, for these two files specifically. Afterwards, check on github.com that
> `.gitignore` and `.streamlit/config.toml` show up as their own files/folder, not as
> `download` or a root `config.toml`.

---

## Step 2 — Deploy on Streamlit

1. Go to **https://share.streamlit.io** and sign in **with GitHub**
2. Click **Create app** → **Deploy a public app from a template**? No — choose
   **Deploy from an existing repo**
3. Fill in:
   - Repository: `your-username/prospect-drafter`
   - Branch: `main`
   - Main file path: `app.py`
   - App URL: pick something like `gateway-prospect-drafter`
4. Before clicking Deploy, open **Advanced settings** and paste into **Secrets**:

   ```toml
   GEMINI_API_KEY = "your-key-from-aistudio"
   ```

5. Click **Deploy**. First build takes 2–3 minutes.

---

## Step 3 — Make it private and invite her

1. In the app's dashboard, open **Settings → Sharing**
2. Set the app to **Private**
3. Add her work email address as a viewer
4. She gets an email with a link

If her email is a Google-linked account she signs in with Google. Otherwise Streamlit emails her
a single-use link each time. Either way she never needs a GitHub account.

---

## Step 4 — Check it yourself first

Open the app, upload `sample_prospects.xlsx`, set the count to 2, and run it. You should get two
drafts using the six sample rows. Fix anything odd before you hand it over.

Then upload your real list once and confirm the column reading is right — the app shows the first
row as it understands it under "Check the columns were read correctly".

---

## Step 5 (later) — Turn on direct-to-Outlook

Until IT does the Entra app registration, she uses the **One at a time** tab, which opens a
pre-filled Outlook Web compose window per draft. That works today and needs no permissions.

Once you have the client ID from `IT-REQUEST.md`, add to the app's Secrets:

```toml
GRAPH_CLIENT_ID = "00000000-0000-0000-0000-000000000000"
GRAPH_TENANT_ID = "organizations"
```

Save, and the app reboots with the **Straight into Outlook** tab live. She clicks Connect, gets a
short code, signs in once per session, and drafts land in her own Drafts folder.

---

## Step 6 (later, optional) — Turn on "Load/save from OneDrive"

The app already has the code for this. It's off until the app registration has one more
permission than `IT-REQUEST.md` currently asks for — **`Files.ReadWrite`** (delegated). That
request hasn't been sent yet; this is here so the feature is ready the moment it has been.

What it adds once it's live: in Step 1 she can choose **Load from OneDrive**, paste a "Copy
link" from her spreadsheet in OneDrive/Excel Online instead of uploading a file, and after
drafting there's a **Save back to OneDrive** button that writes the updated `Status` /
`First Contact Date` / `Last Contact Date` cells straight back to that same file — no
download/re-upload round trip. The **Download updated spreadsheet** button stays as a
fallback either way.

It uses its own **Connect to OneDrive** button in the sidebar, separate from the Outlook one,
so if the new permission isn't approved yet only that button shows "Not available yet" — the
Outlook connection nothing else depends on it.

To switch it on when you're ready:

1. Ask IT to add `Files.ReadWrite` (delegated, work/school account) to the same Entra app
   registration used for `Mail.ReadWrite` — same consent tier, no admin approval beyond what
   `Mail.ReadWrite` already needed.
2. Nothing else changes — the same `GRAPH_CLIENT_ID` / `GRAPH_TENANT_ID` secrets cover both.
3. If she tries "Connect to OneDrive" before the permission is added, the sign-in will fail
   with a message about the scope not being found — that's expected, not a bug. Once IT adds
   it, it works on her next sign-in with no redeploy needed.

**A OneDrive-wide permission, not a one-file one.** `Files.ReadWrite` grants access to her
entire OneDrive while she's signed into the app, the same way `Mail.ReadWrite` grants access
to her whole mailbox rather than one folder. Worth knowing before asking for it.

---

## The templates

Each workflow has its own sequence, one template file per stage, named for what it is:

| Workflow | File | Stage | Mode |
|---|---|---|---|
| In-bound Leads | `inbound_lead_first_contact.md` | 1 | adaptive — AI personalises from the research |
| In-bound Leads | `inbound_lead_first_followup.md` | 2 | verbatim — approved copy, sent exactly as written |
| In-bound Leads | `inbound_lead_second_followup.md` | 3 | verbatim |
| Cold Call | `coldcall_first_contact.md` | 1 | adaptive |
| Cold Call | `coldcall_first_followup.md` | 2 | verbatim |
| Cold Call | `coldcall_second_followup.md` | 3 | verbatim |
| Cold Call | `coldcall_third_followup.md` | 4 | verbatim |

**Adaptive vs verbatim** (`sequence.verbatim_stages` in `config.json`): the first-contact
template for each workflow is a skeleton — approved paragraphs plus one bracketed slot the AI
fills from the prospect's website. The follow-up templates, refined 2026-08-18, are complete
finished copy with no personalisation slot at all, so they're marked verbatim: the AI is told
to reproduce them word for word, substituting only the greeting name (and `{{company}}` where
it appears), rather than adapt them. See `is_verbatim_stage()` in `prospect_drafter.py`.

**`generic-templates/`** holds the pre-refinement versions of the follow-up templates — the
ones that used AI-adapted personalisation instead of fixed copy — kept for reference. They
aren't wired into `config.json` and the app never reads them.

**Subject lines** are fixed per workflow (`sequence.fixed_subject`), not AI-chosen, for every
stage in that workflow — whatever subject a template file suggests is ignored.

Edit a template on GitHub (click the file, click the pencil, commit) and the app picks the
change up within a minute. No redeploy.

Change the 7-day gap in `config.json` under `sequence.wait_days`.

## The "I tried calling" sentence

Tooka's first email contains: *"I tried calling earlier to connect directly but wasn't
successful."* That is a **factual claim about something that happened**, not marketing copy.
Sent to someone nobody phoned, it is untrue, and it is the sort of thing a prospect can catch
out — several of them have no phone number in the sheet at all.

So it is off by default and controlled by a tick box above the draft button: **"I tried
phoning these people first"**. Unticked, the AI is instructed in the strongest terms to omit
the sentence and never imply a call happened. Ticked, it goes in.

The tick box only appears on the first email, since the follow-ups don't mention calling.
Change the default in `config.json` under `sender.mention_prior_call`.

## How the app tracks who's due

Both workflows are **status-driven** (`sequence.gate: "status"` in `config.json`). The sheet's
own `Status` dropdown is the single source of truth for where a row is in the sequence, and
`sequence.status_flow` maps it:

| Status | Due | Written back as |
|---|---|---|
| blank | email 1 | `First Contact` |
| `First Contact` | email 2 | `Follow-up 1 Sent` |
| `Follow-up 1 Sent` | email 3 | `Follow-up 2 Sent` |

Both sheets carry the same ten-value dropdown. `Replied – *`, `No Response`, `Not a Fit`,
`Moved to Active` and `Follow-up 3 Sent (Final)` are listed in `stop_statuses` and take a row
out permanently. `Follow-up 2 Sent` simply matches no stage, so it isn't offered again either.

Per workflow, `writeback` names the three cells to update:

| Workflow | Sheet | Status | First contact | Last contact |
|---|---|---|---|---|
| `internal_lead` | `Inbound Leads` | C | D | E |
| `cold_call` | `Cold Database in Work` | A | D | E |

A `Touches` column is no longer used by either workflow, though the touches gate is still in
the code (`sequence.gate: "touches"`) for any sheet that wants it.

Adding a fourth email is a config change, not a code change: add a `status_flow` entry `"4"`
going from `Follow-up 2 Sent` to `Follow-up 3 Sent (Final)`, a `templates` entry, a `labels`
entry, and drop `follow-up 3 sent` from `stop_statuses`.

Dates are read day-first (`01/08/2026` is 1 August), matching UK convention.

**Write-back is surgical, not a re-save.** `xlsx_patch.py` edits only the target cells inside
the one sheet's XML and copies every other part of the workbook through byte-identically. This
matters: a plain openpyxl load/save of this master report silently drops the nine dropdowns on
`WLCC Active Leads`, because they're stored as x14 extension validations that openpyxl doesn't
support. Never swap this for `openpyxl.save()`.

## The reply check

When Outlook is connected, the app can search her mailbox for a message from each prospect
since their last contact date and drop anyone who replied. It uses `Mail.Read`, which the
`Mail.ReadWrite` permission in `IT-REQUEST.md` already covers — no change needed to what you
ask IT for.

Until that's live, the manual route works: she types `Replied` in the Status column.

---

## Costs and limits

- Streamlit Community Cloud: free
- Gemini free tier: free, limits shown in your AI Studio dashboard
- The app sleeps after inactivity and wakes on first visit — expect a 30-second wait if she
  hasn't used it in a while. That's normal, not a fault.

---

## If something breaks

The app's dashboard on share.streamlit.io has a **Manage app** panel with live logs. Most
failures are one of three things: a missing secret, a column name changed in the spreadsheet, or
the Gemini free tier rate-limiting a large batch. The logs say which.
