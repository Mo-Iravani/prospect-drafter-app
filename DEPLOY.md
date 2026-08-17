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
   config.json
   template.md
   template_followup_1.md
   template_followup_2.md
   requirements.txt
   sample_prospects.xlsx
   .gitignore
   .streamlit/config.toml
   ```

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
drafting there's a **Save back to OneDrive** button that writes the updated `Touches` /
`First Contact Date` / `Last Contact Date` columns straight back to that same file — no
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

## The three templates

The app runs a three-email sequence, one template per stage:

| File | Stage | When |
|---|---|---|
| `template.md` | First email | Never contacted |
| `template_followup_1.md` | First follow-up | 7+ days after email 1, no reply |
| `template_followup_2.md` | Second follow-up | 7+ days after email 2, no reply |

**Status of each file:**

- `template_followup_1.md` — **Tooka's real approved copy**, from the Word document. The three
  WLCC paragraphs are reproduced faithfully; the AI only inserts one personalised sentence
  after the greeting, and omits even that if it found nothing real to say.
- `template_followup_2.md` — **adapted**, not supplied. Cut back to a short, graceful final
  note. Replace it if Tooka has genuine third-email copy.
- `template.md` — **Tooka's real approved copy**, from the LandWey example. The research
  paragraph is the heart of it, with explicit rules for degrading gracefully when the website
  gives the AI nothing to work with.

Edit them on GitHub (click the file, click the pencil, commit) and the app picks the change up
within a minute. No redeploy.

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

Three columns it owns, created automatically if your sheet doesn't have them:

- **`Touches`** — how many emails that person has had (0, 1, 2, 3)
- **`First Contact Date`** / **`Last Contact Date`** — set when drafts are created

The **`Status`** column stays yours. Anything starting with `Replied`, `Skip`, `Do not
contact`, `Unsubscribed`, `Bounced`, `Customer`, `Won` or `Lost` takes that person out of the
sequence permanently.

Sheets written by the earlier version — where Status read `Drafted 17 Aug 2026` and there was
no Touches column — are handled: those rows count as one touch, so nobody gets the first email
twice.

Dates are read day-first (`01/08/2026` is 1 August), matching UK convention.

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
