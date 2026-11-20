# Makale Editör 2 — Orthopedic Journal Chief Editor AI

A small React + Vite web app that gives editor-style feedback on an orthopedic manuscript. You paste the title, abstract and body text; the app sends it to the Gemini API and returns structured editorial comments.

## Run locally

Prerequisites: Node.js

1. Install dependencies: `npm install`
2. Create `.env.local` and set `GEMINI_API_KEY` to your own Gemini API key. Never commit this file.
3. Start the app: `npm run dev`

## Notes

- The repository contains no patient data, and the app needs none. Paste only de-identified text.
- The output is advisory. It does not replace editorial or peer review.

## License

MIT. See [LICENSE](LICENSE).
