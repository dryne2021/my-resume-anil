# Dryne Agency — Resume Tailor

Private one-user portal: upload a sample Word resume, paste a job description, download a tailored .docx in the same format.

No database. Nothing is stored on the server:
- Login is checked against the `OWNER_EMAIL` / `OWNER_PASSWORD` environment variables and kept in a signed cookie.
- Your sample resume stays in your browser and is sent with each "Generate" request, then discarded.
- The generated resume is returned straight to your browser for download.

## Deploy on Vercel

1. Push this folder to a GitHub repo and import it in Vercel (default settings are fine).
2. In Vercel → Settings → Environment Variables, add:
   - `OPENAI_API_KEY` — from platform.openai.com
   - `OWNER_EMAIL` — the email you log in with
   - `OWNER_PASSWORD` — the password you log in with
   - `SESSION_SECRET` — any long random string
3. Redeploy.

## Run locally

```sh
cp .env.example .env   # fill in values
npm install
npm run dev
```
