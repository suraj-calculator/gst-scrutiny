# Login Page for Code Privacy — Feasibility Review

**Question asked:** Can adding a password/OTP login page (per `streamlit-private-otp-deploy-guide.md`)
keep the code private and safe on GitHub?

**Short answer:** Not as currently architected. The guide assumes a different app (a live
Streamlit server on Render). This repo has already moved to a **static, client-side,
no-backend site** (Pyodide, deployed to GitHub Pages). A login page bolted onto that
static site cannot hide the source or the data — and right now, nothing is hidden anyway.

---

## 1. Verified current state (not guesses)

Checked live, just now:

| Fact | Evidence |
|---|---|
| GitHub repo `suraj-calculator/gst-scrutiny` is **public** | Unauthenticated `curl` to the repo page returns `200` |
| `docs/` auto-deploys to GitHub Pages on every push to `main` | `.github/workflows/deploy-pages.yml` |
| The Pages site is live and public | `https://suraj-calculator.github.io/gst-scrutiny/` → `200` |
| The actual scrutiny engine source is directly fetchable, no auth | `.../py/core/gst_core.py` → `200`, returns real Python source |
| The app runs entirely in the browser via Pyodide (no Python backend) | `docs/index.html`, `docs/worker.js`, `docs/app.js`, `docs/py/*` — this matches README's "Status: being turned into a website — client-side, Pyodide, no backend" |
| `docs/sample_data/` and `builds/*` contain real-looking GSTINs and full financial ledgers | e.g. `docs/sample_data/gstr2b/062023_05AAGCA2491N1ZG_...xlsx`, `builds/05AACFT2702L1ZD_FY2025-26/GST_MASTER_...xlsx` |

So today: anyone in the world can already read the repo on GitHub, browse the live site,
and pull every `.py` file and every sample workbook by URL. **Worth confirming whether
the data in `docs/sample_data/` and `builds/` is synthetic test data or a real client's
GST filings** — if it's real, it's currently exposed publicly regardless of any login
page decision.

## 2. Why the guide's approach doesn't transfer

`streamlit-private-otp-deploy-guide.md` describes a **Streamlit app running as a live
Python process on Render**, reading a private GitHub repo. In that model, the login
gate works because:

- The Python source **never leaves the server** — the browser only ever receives
  rendered HTML/websocket updates, not the `.py` files.
- The password/OTP check runs server-side, *before* any app logic executes.
- "Private repo" + "server never ships source to the client" together mean the code
  genuinely stays hidden from an unauthenticated visitor.

This repo's actual deployment is the opposite of that model:

- It's a **static site** (GitHub Pages). There is no server process per request —
  Pages just serves files.
- The "backend" is Pyodide running **inside the visitor's own browser**, which means
  every `.py` file the engine needs must be downloaded to that browser as plain text,
  by definition, for the tool to work at all.
- A password screen built in client-side JS/HTML only hides a UI element. It cannot
  stop a request to `.../py/core/gst_core.py` — view-source, devtools Network tab, or
  a plain `curl` all bypass it trivially, because nothing on the server side is
  checking who's asking.

Net: implementing the guide's login-screen code on top of the current site would be
**security theatre** — it would look protected, and would not be.

## 3. Two separate problems, being conflated

1. **"Is my GitHub repo private?"** — a repo Settings toggle, unrelated to any login
   page. Free to do right now. Does **not** by itself stop the *Pages site* from being
   public — GitHub Pages sites are published publicly even from private repos unless
   the account is on GitHub Pro/Team/Enterprise (which adds a "private Pages
   visibility" option).
2. **"Can I gate who uses the deployed app?"** — this is what a login page actually
   controls. On a real backend, gating the app also gates the code (nothing ships
   until you're authed). On a static/Pyodide site, gating the *UI* does not gate the
   *files* — they're already sitting on the CDN.

## 4. Options, in order of fit

**A. Fix the immediate exposure (do regardless of anything else)**
Flip the repo to Private. Confirm whether `docs/sample_data/` / `builds/*` is real
client data; if so, treat it as already leaked (public git history can be cloned by
anyone who saw it before you flip visibility) and consider purging history / rotating
whatever's sensitive.

**B. Keep the static/Pyodide architecture, gate access at the hosting edge**
Move hosting from GitHub Pages (no access-control option on Free) to something that
can put a real auth check *in front of* the static files — e.g. Cloudflare Pages +
Cloudflare Access with email one-time-PIN (free for small teams, each of your 3 users
gets their own email + their own OTP, closer to real per-user auth than the guide's
single shared inbox). This is the closest thing to "free, private, and matches what's
already built" — but it's a hosting migration, not a code change to `app.py`.

**C. Revert to a real backend (what the guide actually describes)**
Rebuild/re-point `main gst tool/` as a server-side Streamlit app on Render per the
guide. This genuinely hides the code and works as documented — but it means reversing
the "client-side, Pyodide, no backend" direction the repo has already been built
toward (docs/index.html, worker.js, docs/py/*). That's a real architectural rollback,
not a small addition.

**D. Do nothing about hosting, add the login screen anyway**
Only makes sense as a speed bump against casual/accidental visitors (e.g., someone
finding the URL by accident). Does not satisfy "code private and safe" against anyone
who opens devtools.

## 5. Secondary issues in the guide itself (only relevant if you go with C)

- One shared `APP_PASSWORD` + OTP sent to one fixed `RECEIVER_EMAIL`: doesn't give 3
  users independent logins — all 3 need access to the same inbox to get the OTP, or
  you're back to one person unlocking it for everyone.
- No lockout/rate-limit on password or OTP attempts.
- No session expiry — once authenticated, stays authenticated until the server
  restarts (Render free tier sleeps, so this resets somewhat naturally).
- Password compared with plain `==` (not constant-time) — negligible practical risk
  at this scale, just noting it.
- These are all fine for a small internal tool but worth knowing they're informal,
  not audited-grade auth.

## Recommendation

Do **A** immediately (it costs nothing and closes real, currently-live exposure).
Then decide between **B** (stay static, add a real edge gate — matches the direction
you've already invested in) and **C** (go back to a live backend — matches the guide
verbatim, but is a bigger rework). I'd lean **B** given how much of `docs/` is already
built out, but that's your call to make, not something to decide by default.

No code changes have been made — this is analysis only, per your request.
