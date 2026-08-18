# How to use the Prospect Drafter

It reads each prospect's website and writes a personalised email for them. You read every one
before anything happens. **It never sends anything** — you always press Send yourself.

---

## Pick your workflow first

At the top of the page there are two buttons. Choose before you do anything else, because
everything below changes with it.

**Internal lead** — the one you've been using. Leads that have already had some contact with
WLCC. The app works out who's due from a `Touches` count.

**Cold Call** — cold outreach from the WLCC master report. The app reads the
**Cold Database in Work** sheet, you pick which **WLCC Fit Scores** to work, and it tracks
progress in the **Status** column (column A) using the same dropdown values already in the
sheet. The emails are written cold: they never claim the person contacted WLCC first.

Both workflows read the same master report, just different sheets — so you upload the one
file either way.

If you switch workflow after writing drafts, the drafts are cleared. That's deliberate — it
stops a Cold Call batch being recorded into an Internal lead sheet by mistake.

---

## Opening it

Click the link in the invitation email and sign in with your work email address. Bookmark the page.

If it takes half a minute to load, that's normal — it goes to sleep when nobody's using it.

---

## 1. Upload your list

Click **Browse files** and choose your prospect spreadsheet. The one you already keep.

You'll see three numbers: how many rows it found, how many are ready, and how many have no
website. Anyone already marked as done in the Status column is skipped automatically.

Your file isn't changed. The app only reads it.

### The Active sheet

Under the upload box the app shows the **Active sheet** — the one it is reading. It picks this
for you from the workflow you chose:

| Workflow | Sheet it reads |
| --- | --- |
| Internal lead | `Inbound Leads` |
| Cold Call | `Cold Database in Work` |

Normally you can just glance at it and carry on. You only need to touch the dropdown if you
keep that data on a differently named sheet, and if you do change it the app warns you, because
the column names on the sheet you pick have to match what the workflow expects. If the sheet the
workflow wants isn't in the file at all, it says so rather than reading the wrong thing quietly.

Whichever sheet is shown there is the sheet that gets written back to in step 6. The sheet
marked *active in Excel* is simply the tab that was open when the file was last saved.

**On Cold Call**, after that you get a **WLCC Fit Score** box: tick the scores you want to work,
and you can tick more than one. 5 is the best fit. Rows with no score yet are left out, and the
app tells you how many that was.

**If "Load from OneDrive" is available** (it may not be yet — ask Mo), you can skip the
upload/download entirely. Click **Connect to OneDrive** in the sidebar once per session, then
choose **Load from OneDrive** at the top of this step, paste a "Copy link" from your
spreadsheet in OneDrive or Excel Online, and click **Load from OneDrive**. Later, in step 6,
you'll get a **Save back to OneDrive** button that writes the tracking columns straight back
to that same file — no downloading and re-uploading needed.

---

## 2. Choose which email you're sending

Three options, and the app tells you how many people are ready for each:

- **First email** — people who haven't been contacted yet.
- **First follow-up** — people who got the first email, haven't replied, and were last
  contacted at least a week ago.
- **Second follow-up** — same again, one step further along. This is the last one; nobody
  gets a fourth email.

On **Cold Call** the three stages are just called First, Second and Third email, and they're
worked out from the Status column instead:

| Status now | Stage they're due | Status becomes |
| --- | --- | --- |
| blank | First email | `First Contact` |
| `First Contact` | Second email | `Follow-up 1 Sent` |
| `Follow-up 1 Sent` | Third email | `Follow-up 2 Sent` |

Anything else in that column means the app leaves the row alone: `Replied – Positive`,
`Replied – Neutral`, `Replied – Negative`, `No Response`, `Not a Fit`, `Moved to Active`, and
`Follow-up 3 Sent (Final)`. A row marked `Interested? = NO` is skipped too, whatever its
Status says.

You don't have to work out who's due. The app does that from your spreadsheet. If a stage
shows zero, click **Why not?** and it explains each person's situation.

**Before drafting a follow-up**, if Outlook is connected, click **Check who has replied
first**. The app looks through your mailbox and removes anyone who has already written
back — so nobody gets a "following up" email after they've replied. If Outlook isn't
connected, mark those people `Replied` in the Status column yourself and they'll be skipped.

---

## 3. Write the drafts

Choose how many to do, then click the **Write the drafts** button.

**Start with 2 or 3 the first few times.** Read them, see whether the tone is right, then do the
rest. Each one takes a few seconds because the app is reading their website first.

**The "I tried phoning these people first" tick box** (first email only) adds the sentence
about having tried to call. Only tick it if you actually rang them — everyone in that batch.
Left unticked, the sentence is left out completely.

On **Cold Call** the box starts unticked, because Tooka's approved cold email doesn't mention a
call. Tick it only for a batch you have actually rung, and it adds one short line saying you
tried to reach them. The app also passes the `Why WLCC?` note and the `Call Result` note from
the sheet to the AI, so a good `Why WLCC?` line makes a noticeably better email.

The Cold Call first email follows Tooka's approved sample: "I hope this email finds you well",
then one line on what specifically impressed you about the firm, then the WLCC "select number
of firms" paragraph, then the ask. The first email signs off **Best Regards**; emails two and
three sign off **Kind regards**. That difference is deliberate.

---

## 4. Read them

Every draft opens up so you can read it. You can:

- **Change any wording** — click into the subject or the message and type. It's just text.
- **Untick "Include this one"** for any you don't want.

Two things to look at:

- Drafts marked **⚠️ needs a look** are ones where the website couldn't be read. They fall back
  to the standard template, so they're the generic ones. Read those first.
- The small **Angle:** line under each name tells you what detail the app picked up on. If it
  says something odd or wrong, fix the email or untick it.

Trust your judgement over the app's. If a sentence sounds off, it is off.

---

## 5. Put them in Outlook

Three tabs. Use whichever is available:

**Straight into Outlook** — click Connect, you'll see a short code, sign in once, and all your
approved drafts appear in your Outlook Drafts folder. Open each, check, send.

**One at a time** — click a name and a new Outlook message opens with the recipient, subject and
message already filled in. Check it and press Send. Then come back for the next one.

**Download files** — only if you've been asked to use it.

---

## 6. Update your spreadsheet

If you loaded your list from OneDrive, click **Save back to OneDrive** — it writes straight to
your file, no download needed. (If it says the file is locked, close it in Excel first and
click the button again.) Otherwise, click **Download updated spreadsheet** and save it over
your own copy — this button is always there too, even when OneDrive is connected, as a backup.

**This is the step that makes follow-ups work.** Without it, the app can't tell who's due a
follow-up next week — and you risk sending someone the same first email twice.

On **Internal lead** it records how many emails each person has had and the date, in three
columns: `Touches`, `First Contact Date` and `Last Contact Date`. If you don't send a draft
after all, set that person's `Touches` back down by one, or set their Status to `Skip`.

On **Cold Call** it moves each row's **Status** on one step and stamps the two date columns.
Before you save, open **Which rows change** and check the list — it names every row number and
company it's about to touch. If you don't send a draft after all, set that row's Status back to
what it was.

The Cold Call save is deliberately careful with the master report: it edits only those cells on
the Cold Database sheet and leaves the rest of the workbook exactly as it was, so the dropdowns
and the colour rules on the other sheets survive. Even so, keep your own copy the first time
you use it.

---

## A few things worth knowing

**Nothing is sent automatically.** Every single email waits in Outlook until you press Send.

**Read every draft.** The app is good but not perfect. It's been told never to make things up,
and to fall back to the generic template when it can't find real information — but you're the
one whose name is on the email.

**If a draft mentions something that isn't true about that company**, untick it and tell Mo.
That shouldn't happen and he'll want to know.

**Status words the app respects.** On Internal lead, type any of these into the Status column
and that person is left alone permanently: `Replied`, `Skip`, `Do not contact`, `Unsubscribed`,
`Bounced`, `Customer`, `Won`, `Lost`. On Cold Call, use the dropdown already in column A — see
the table in step 2. Either way, mark someone as replied the moment they answer.

**If you get an error**, take a screenshot and send it to Mo. Nothing you do in the app can
break anything or send anything by accident.

**Your prospects' details** go to Google's AI service to write the emails. Don't paste anything
into the app that you wouldn't be comfortable sending outside the company.
