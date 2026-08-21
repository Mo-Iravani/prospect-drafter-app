# How to use the Prospect Drafter

Two of its three workflows read each prospect's website and write a personalised email for
them. The third sends one email you wrote yourself to a whole list. Either way you read every
one before anything happens, and **it never sends anything** — you always press Send yourself.

---

## Pick your workflow first

At the top of the page there are three buttons. Choose before you do anything else, because
everything below changes with it.

**In-bound Leads** — people who came to WLCC. The app reads the **Inbound Leads** sheet and
tracks progress in its **Status** column (column C).

**Cold Call** — cold outreach. The app reads the **Cold Database in Work** sheet, you pick
which **WLCC Fit Scores** to work first, and it tracks progress in that sheet's **Status**
column (column A). The emails are written cold: they never claim the person contacted WLCC
first.

Those two read the same master report, just different sheets — so you upload the one file
either way, and both work identically: the Status dropdown is what tells the app which email
is due, and the Status dropdown is what it updates afterwards. The only real differences are
which sheet they read, and that Cold Call asks for a Fit Score first.

**Batch Email** — one email you have already written, sent to a list. You supply the wording,
the spreadsheet supplies the names, references and anything else that changes. There is no AI
in this one at all: it goes out exactly as you typed it. **This one has its own instructions
at the bottom of this page** — the steps below it are for the other two.

If you switch workflow after writing drafts, the drafts are cleared. That's deliberate — it
stops one workflow's batch being recorded into the other's sheet by mistake.

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

| Workflow | Sheet it reads | Status column |
| --- | --- | --- |
| In-bound Leads | `Inbound Leads` | C |
| Cold Call | `Cold Database in Work` | A |

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

Both workflows work this out from the Status column. On **Cold Call** the three stages are
called First, Second and Third email; on **In-bound Leads** they keep the names above. Either
way the rule is the same:

| Status now | Stage they're due | Status becomes |
| --- | --- | --- |
| blank | First email | `First Contact` |
| `First Contact` | Second email | `Follow-up 1 Sent` |
| `Follow-up 1 Sent` | Third email | `Follow-up 2 Sent` |

Anything else in that column means the app leaves the row alone: `Replied – Positive`,
`Replied – Neutral`, `Replied – Negative`, `No Response`, `Not a Fit`, `Moved to Active`, and
`Follow-up 3 Sent (Final)`. Rows already at `Follow-up 2 Sent` have had all three emails, so
they aren't offered again either. On Cold Call, a row marked `Interested? = NO` is skipped too,
whatever its Status says.

A blank Status is what makes a row due a *first* email. So if every row on a sheet already has
a status, "First email" will show zero ready — that's correct, not a fault. New rows you add
with a blank Status will appear there.

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

**Straight into Outlook** — all your approved drafts appear in your Outlook **Drafts** folder.
Open each, check, send. This is the route to use whenever it's available: there's no limit on
how long the email is, and nothing to download.

There are two ways it can work, and the app picks whichever applies without being asked:

- **Outlook on this computer.** If you're running the app on your own PC with Outlook
  installed, it hands the drafts straight to Outlook. No sign-in, no setup at all. Outlook
  starts up if it isn't already running. (Needs classic Outlook — the new Outlook for Windows
  can't do this.)
- **Outlook online.** On the shared web version there's no Outlook on the server, so it signs
  in instead: click Connect, you'll see a short code, sign in once. This needs the one-off IT
  setup; until that's done this tab says so.

**One at a time** — click a name and a new Outlook message opens with the recipient, subject and
message already filled in. Check it and press Send. Then come back for the next one.

If a draft says **save as a draft, then open it** instead of opening straight away, that email
is simply too long to travel inside a web address — nothing is wrong with it. The whole message
has to fit in the address for the one-click version to work, and longer emails don't. This
happens more with Batch Email, since a written-out email is usually longer than the AI's
120-word ones.

When that happens, **use Straight into Outlook instead** — no length limit, and the drafts go
to Drafts without downloading anything. The button here saves the draft as a *file* you'd then
have to open, which is only worth doing if you actually want the file.

**Download files** — only if you've been asked to use it.

---

## 6. Update your spreadsheet

If you loaded your list from OneDrive, click **Save back to OneDrive** — it writes straight to
your file, no download needed. (If it says the file is locked, close it in Excel first and
click the button again.) Otherwise, click **Download updated spreadsheet** and save it over
your own copy — this button is always there too, even when OneDrive is connected, as a backup.

**This is the step that makes follow-ups work.** Without it, the app can't tell who's due a
follow-up next week — and you risk sending someone the same first email twice.

In **both workflows** it moves each row's **Status** on one step and stamps the two date
columns. Before you save, open **Which rows change** and check the list — it names every row
number and company it's about to touch. If you don't send a draft after all, set that row's
Status back to what it was.

The save is deliberately careful with the master report: it edits only those few cells on the
one sheet you were working, and leaves the rest of the workbook exactly as it was, so the
dropdowns and the colour rules on the other sheets survive. Even so, keep your own copy the
first time you use it.

---

# Batch Email

A different job from the other two: **one email, already written and approved, sent to a list**
where only a few words change per person. No AI, no website reading. A hundred drafts appear
instantly and say exactly what you typed.

## 1. The email you're sending

Type or paste it, or upload a Word file. Wherever a word changes per person, put the
spreadsheet's column name in double braces:

```
Dear {{First Name}},

Your membership reference is {{Ref No.}} and {{Company Name}} is listed for renewal
on {{Renewal Date}}.

Kind regards,
Tooka
```

Capitals and spacing don't have to match your spreadsheet exactly — `{{first name}}` finds a
column called "First Name", and `{{Ref No}}` finds one called "Ref No.".

**The subject line is a separate box, and it takes placeholders too** — from the same columns.
`Renewal — ref {{Ref No.}}` is fine.

**A blank cell stops that row.** If someone's Company Name cell is empty, that row is blocked
rather than sent as "Dear Amelia, and is listed for renewal". Where a bit genuinely doesn't
apply to everybody — a job title some people don't have — put a question mark on the end,
`{{Job Title?}}`, and a blank cell is then allowed: it comes out as nothing and the line
closes up.

### Links

A box of plain text has nowhere to keep a link's address. So if you paste a hyperlink in from
Word or Outlook, the words arrive and **the address is silently gone** — which is why a pasted
link looks like it stopped working. Write it down instead, either way round:

```
Have a look at our [member directory](https://example.com/directory).

Or just the address on its own: https://example.com/directory
```

Both become a proper clickable link in the finished email. The address can have a placeholder
in it too, so everyone gets their own link:

```
Renew at https://portal.example.com/renew?ref={{Ref No.}}
```

Under the email box the app tells you how many links it found and where each one points, so you
can see at a glance that it picked yours up. If it says **No links found** and your email is
meant to have one, that's the thing to fix before going any further.

**Uploading from Word keeps your hyperlinks** — they come through written in the
`[words](address)` form above, so you can see and correct exactly where each one goes.

**Everything else about Word formatting does not survive:** bold, colour, fonts and images are
dropped, and tables come out as plain lines. The wording appears in a box for you to read and
fix before going on. If the layout matters, paste it in and lay it out with blank lines
instead.

**Your sign-off:** if the email you paste already ends with "Kind regards, Tooka", the app
notices and leaves the sidebar signature off, so nobody gets it twice. The tick box is there
if it guesses wrong.

## 2. Your list

Upload the spreadsheet — one row per recipient, up to 100 at a time. Pick the sheet, and tell
it which row holds the column names if it isn't row 1. The file is only read, never changed.

## 3. Which column is which

It matches your placeholders to your columns by name and shows you what it worked out. Anything
it couldn't match gets a dropdown — nothing can be drafted until every placeholder has a
column. You also confirm which column holds the **email addresses**; that is the one thing it
can't work around.

## 4. Check the list

Three numbers: **Ready**, **Blocked**, **Duplicate addresses**.

Blocked rows are listed by row number with what's wrong — a blank cell, a missing address, an
address that isn't valid. Fix them in the spreadsheet and upload it again. They are named
rather than silently dropped, because a batch going out thirteen short without anyone noticing
is the thing worth preventing.

Duplicates are a warning, not a block: the same address twice means that person gets two
emails. Sometimes that's a shared inbox and it's fine — your call.

## 5. Read one first

Flick through a few. This is what catches what no check can: a column that reads oddly
mid-sentence, a reference in the wrong format, a date that came out American.

## 6, 7, 8. Make them, read them, send them

Make the drafts, and you get a table of all of them. Leave anyone out by naming them in
**Leave anyone out?**, or open **Change one of them by hand** to edit a single draft — edits
there stay on that one draft; change the email itself in step 1 to change them all.

Then the same three routes into Outlook as the other workflows. For a hundred, use **Straight
into Outlook** — they land in your Drafts folder ready to read and send, with no limit on how
long the email is. Batch emails are usually longer than the AI-written ones, so on **One at a
time** expect more of them to come as a draft file to open rather than a link that opens by
itself; see step 5 of the other workflows above for what that means.

**Last check:** if any draft still has a `{{...}}` gap in it — usually because a subject line
was edited by hand — it is held back and named, and cannot go to Outlook.

## Keep a record

Download the log at the bottom. One line per row of your list, including the blocked ones, with
what happened to each. Worth keeping for its own sake, and next week it is the only way to tell
who already had theirs. Batch Email does not write to your spreadsheet.

---

## A few things worth knowing

**Nothing is sent automatically.** Every single email waits in Outlook until you press Send.

**Read every draft.** The app is good but not perfect. It's been told never to make things up,
and to fall back to the generic template when it can't find real information — but you're the
one whose name is on the email.

**If a draft mentions something that isn't true about that company**, untick it and tell Mo.
That shouldn't happen and he'll want to know.

**Status words the app respects.** Use the dropdown already in the Status column — the same
list on both sheets. See the table in step 2 for which value means what. Mark someone as
replied the moment they answer, and the app will leave them alone from then on.

**If you get an error**, take a screenshot and send it to Mo. Nothing you do in the app can
break anything or send anything by accident.

**Your prospects' details** go to Google's AI service to write the emails. Don't paste anything
into the app that you wouldn't be comfortable sending outside the company. **Batch Email is the
exception** — it never calls the AI, so nothing on a batch list leaves the app except into your
own Outlook drafts.
