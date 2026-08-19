COLD CALL — SECOND FOLLOW-UP. Sent about a week after the first follow-up, to people who still
did not reply.

Tooka's approved copy for this stage, refined 2026-08-18. This is complete, finished copy —
there is no personalised line to insert this time. Send it exactly as written, with only the
greeting name and {{company}} filled in. This stage is marked verbatim in config.json
(workflows.cold_call.sequence.verbatim_stages), so the AI is instructed to reproduce it word
for word rather than adapt it.

FLAGGED FOR TOOKA: the copy as supplied named one specific company ("Abu Dhabi Global Market's
(ADGM)") in the opening line, which reads correctly for that one firm and wrong for every other
prospect on the list. The line below substitutes {{company}} in that spot so it sends correctly
to everyone. If ADGM was meant to be the actual subject of this email rather than a stand-in,
say so and this goes back to being fixed text for that one send instead of a template.

---

APPROVED COPY — send exactly as written:

Dear {{first_name}},

I wanted to follow up as I believe there could be a strong alignment between {{company}}'s
international wealth advisory focus and the type of relationships being developed within the
World Luxury Chamber of Commerce.

Many of our members operate in sectors where trust, long-term relationships, and international
positioning are more valuable than traditional networking. This is especially relevant for
firms working with high-net-worth and internationally mobile clients.

WLCC brings together luxury decision-makers across sectors such as real estate, hospitality,
private client services, and investment through invitation-only leadership events, curated
introductions, and thought leadership initiatives.

In addition, platforms such as Luxury People Magazine and our executive networking forums help
members strengthen their visibility and positioning within the global luxury community.

Given your cross-border expertise and international client focus, I thought this could be
relevant for your team.

Would you be open to a short introductory conversation next week?

Kind regards,
Tooka

---

Rules for this stage:

- This is approved, final copy. Reproduce it in full — no personalised sentence, no shortening,
  no rewording, other than filling in {{first_name}} and {{company}}.
- Keep the greeting as "Dear [first name]," — not "Hi".
- The sign-off stays as written.
- Subject line: fixed by the app (sequence.fixed_subject in config.json — "Invitation: Joining
  the World Luxury Chamber of Commerce"). Whatever is written here is ignored, for every stage
  in this workflow.
