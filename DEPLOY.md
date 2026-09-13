# Putting this app online under your own name

Plain-language version. You need three things, and you already have two of them.

| Thing | What it is | Where yours lives |
|---|---|---|
| The code | This repo | GitHub: `aabirsarkar-sde/imr-2` |
| The database | Where every parsed reading is stored | Supabase (free) — the `DATABASE_URL` |
| The host | The computer that runs the app 24/7 and gives it a web address | This is the part you're setting up |

The host is the only piece that changes. Nothing about the app, the data, or the
database moves.

---

## Why move off Streamlit Community Cloud

Two things there cannot be removed by code, and both matter if you're selling this:

1. The address is `something.streamlit.app` — the framework's name is in your
   product's URL.
2. Community Cloud shows visitors a "Hosted with Streamlit" badge, and its terms
   forbid removing or hiding it. It is also a free *community* tier, not a place
   to run software you charge money for.

Everything else — the toolbar icons, the menu, the favicon, the tab title — is
already handled in the code (`.streamlit/config.toml` and the CSS block in
`main()`, plus `brand_served_page()`). Move the host and no trace is left.

---

## The free option: Render

Free, no sleep-free guarantee (see the catch below), and the address is
`your-name.onrender.com` — no framework branding anywhere.

### Steps

1. Go to **render.com** and sign up with your GitHub account.
2. Click **New → Blueprint**.
3. Pick the `imr-2` repository. Render finds `render.yaml` in this repo and fills
   in everything by itself — what to install, how to start the app.
4. It will ask you for three values. Paste them in:
   - `DATABASE_URL` — the Supabase connection string. It is the `url` line in
     your local `.streamlit/secrets.toml`, the one starting `postgresql://`.
   - `GEMINI_API_KEY` — only if you use the Scan IMR page. Leave blank otherwise.
   - `GROQ_API_KEY` — same, leave blank if unused.
5. Click **Apply**. First build takes about 5 minutes.
6. You get a link like `https://ro-membrane-health.onrender.com`. That is the
   link you send to a customer.

### The catch, stated plainly

A free Render service **goes to sleep after 15 minutes with nobody on it**, and
the next visitor waits about a minute for it to wake up. For a tool people open a
few times a month that is survivable. For a paying customer it looks bad — the
fix is Render's paid tier (about $7/month) which never sleeps. Switch when
someone is actually paying, not before.

### Updating the app later

Push to `main` on GitHub. Render rebuilds and redeploys on its own. That's it.

---

## If you want your own web address

`imr.yourcompany.com` instead of `...onrender.com`. This is the one part that
genuinely cannot be free — a domain costs roughly ₹800–1,200 a year from
Namecheap, GoDaddy or Cloudflare. Once you own one, Render's **Settings → Custom
Domain** walks you through it and issues the HTTPS certificate for free.

---

## Other hosts, and why not

- **Railway** — no real free tier any more; a trial credit, then about $5/month.
- **Google Cloud Run** — has a genuinely generous free allowance and never
  sleeps, but it requires a credit card on file and Docker, which is a lot of new
  words for the benefit.
- **Hugging Face Spaces** — free and needs no card, but your app sits on
  `huggingface.co/spaces/...` inside their site chrome. That is the same problem
  you're leaving Community Cloud to escape.

Render is the shortest path to "free, mine, and unbranded".

---

## One rule, forever

`.streamlit/secrets.toml` holds your database password. It is in `.gitignore` and
must stay there. On a host, secrets go in the host's own environment-variable
settings (which is what `sync: false` in `render.yaml` arranges) — never in the
repo.

---

## Licensing, since you're selling

Streamlit is Apache-2.0 licensed. You may sell software built with it, you owe no
fee, and you are not required to display its name anywhere. The only obligation
is to keep the licence notices inside the framework's own files — which is why
`brand_served_page()` rewrites the page title and favicon but deliberately leaves
the Apache header comment in `index.html` alone.
