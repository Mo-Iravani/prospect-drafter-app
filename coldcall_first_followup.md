COLD CALL — FIRST FOLLOW-UP. Sent about a week after the first email, to people who did not
reply.

Tooka's approved copy for this stage, refined 2026-08-18. This is complete, finished copy —
there is no personalised line to insert this time. Send it exactly as written, with only the
greeting name filled in. This stage is marked verbatim in config.json
(workflows.cold_call.sequence.verbatim_stages), so the AI is instructed to reproduce it word
for word rather than adapt it.

---

APPROVED COPY — send exactly as written:

Dear {{first_name}},

I hope you are well.

I wanted to briefly reconnect regarding my previous email, as we are currently expanding World
Luxury Chamber of Commerce (WLCC) relationships within the private banking and wealth
management sector globally and I believed your firm could be a particularly strong fit.

Should this be of interest, I would be very happy to arrange a short introductory conversation
at a convenient time.

Kind regards,
Tooka

---

Rules for this stage:

- This is approved, final copy. Reproduce it in full — no personalised sentence, no shortening,
  no rewording.
- Keep the greeting as "Dear [first name]," — not "Hi".
- The sign-off stays as written.
- Subject line: fixed by the app (sequence.fixed_subject in config.json — "Invitation: Joining
  the World Luxury Chamber of Commerce"). Whatever is written here is ignored, for every stage
  in this workflow.
