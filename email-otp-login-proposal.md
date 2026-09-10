# Proposal: Email + OTP Login (Allowlisted Addresses)

**Status:** Parked for later. A simpler hardcoded username/password login is being built
first (see the branch discussion in conversation — no separate doc for that one, it's
small enough to track directly). This document captures the OTP approach so it isn't
lost, to pick up whenever it's prioritized.

**Goal:** Replace (or sit alongside) a hardcoded login with a real per-login OTP sent
to one of 2-3 specific, approved email addresses — closer to genuine "only these
named people" access than a single shared password.

---

## What this is, architecturally

The app (`docs/index.html` + `docs/app.js` + Pyodide) is a static site with no backend.
This plan adds a login gate **entirely within that static site** — no server, no
Cloudflare, no hosting change — using EmailJS to send the OTP email directly from the
browser.

**Important limitation, carried over from the earlier feasibility review**
(`login-page-feasibility-review.md`): this gates the login screen for the 2-3
approved addresses. It does **not** hide the underlying `.py`/data files from someone
who bypasses the JS and requests them directly — that only closes with a hosting-level
gate (Cloudflare Access) or a real backend. This plan solves "only these people should
get past the front door," not "the source code is invisible to anyone who looks."

## Code changes

1. **`docs/index.html`**
   - Add the EmailJS SDK `<script>` tag (CDN), loaded before `app.js`.
   - Add a `login-overlay` block near the top of `<body>`, styled with the existing
     CSS variables already defined in this file (same pattern as the current
     `scrutiny-overlay`) — two steps: enter email → enter OTP.

2. **New file `docs/auth.js`**
   - `ALLOWED_EMAILS` — the 2-3 approved addresses, hardcoded, checked case-insensitively.
   - On submit: check the email against the allowlist → generate a 6-digit code →
     call `emailjs.send(...)` with the service/template/public key → show the
     OTP-entry step.
   - Verify the entered code against the generated one (kept in memory +
     `sessionStorage`, ~5-minute expiry, a small attempt limit).
   - On success: store an auth flag (e.g. `localStorage` + timestamp, so it doesn't
     ask again for ~12h on the same browser) and fire a custom `auth:ok` event.
   - A resend cooldown (~30s).

3. **`docs/app.js`**
   - The last two lines of the file (`updateRunbar(); initRuntime();`) currently run
     immediately on load and boot Pyodide/the worker. Wrap those in a `startApp()`
     function, called only after `auth:ok` fires (or immediately on load if a
     still-valid auth flag already exists). Nothing else in `app.js` changes.

No changes needed to `.github/workflows/deploy-pages.yml` — same static deploy works
on any branch.

## External configuration (not code — done by the repo owner)

1. Create a free EmailJS account (emailjs.com).
2. Connect an **Email Service** in the dashboard (e.g. link Gmail via their OAuth
   flow — no app password ever goes into the code).
3. Create an **Email Template** with placeholders like `{{to_email}}` and `{{otp_code}}`.
4. Copy the **Service ID**, **Template ID**, and **Public Key** into `docs/auth.js`
   (the public key is meant to be public — not a secret like an SMTP password).
5. Restrict **allowed origins** in EmailJS account settings to the GitHub Pages
   domain, so the public key can't be used to send mail from another site.
6. Decide the final 2-3 allowlisted email addresses.

## Cost

- EmailJS free tier: 200 sends/month, 2 templates — comfortable for 3 users logging
  in occasionally.
- No Cloudflare, no serverless hosting, no per-message SMS cost (SMS OTP was ruled
  out earlier — see feasibility doc — due to India's mandatory TRAI DLT registration
  and per-message cost).

## Open decisions for whenever this is picked back up

- Final list of 2-3 allowlisted emails.
- How long a login should persist on a given browser before re-prompting for OTP
  (proposed default: ~12h via `localStorage` + timestamp).
- Whether this replaces the hardcoded login entirely, or is offered as a second,
  stronger option later.
