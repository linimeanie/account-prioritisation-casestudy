# Weekly account prioritisation (emnify GTM Engineer case study)

A browser-only prototype that turns 500 accounts and their buying signals into each SDR's best accounts for the week, and explains every decision.

**Runs entirely in the browser.** The Excel files are read locally and never uploaded anywhere. There is no backend.

## How to use
1. Drop `accounts.xlsx`, `signals.xlsx` and `sdrs.xlsx` (any order, any file name: files are recognised by their columns).
2. Optionally open **Leadership settings** to boost or exclude segments, pin or exclude accounts, change territories, capacity, signal weights or the time window. You can also upload a `settings.xlsx`.
3. Click **Build this week's lists**. Change any setting afterwards and the lists re-rank immediately.

## What you see
- **SDR lists:** one column per rep, best account first, with why now, fit and a suggested opening.
- **How 500 became 80:** where every account went and why.
- **How the score works:** intent × fit × CRM stage × leadership boost, with the weights in use.
- **Data fixes:** every correction made to the messy input data.
- **Monday messages:** what each rep receives on Monday morning.
- **Account detail:** click any account for its full score breakdown, and pin or exclude it.

## Files
- `n8n/`: the same logic as an n8n workflow (import file, diagram and how to run it).
- `logic.js`: cleaning, scoring and routing. Same logic as the "Score and route" node in the n8n workflow (checked to produce identical lists, scores and data issues on the case data).
- `defaults.js`: the default leadership settings.
- `app.js`, `index.html`, `styles.css`: the interface.

All companies and people in the case data are fictional.
