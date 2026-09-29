# Materials vs Invoice

Shared, GitHub-hosted version of the Materials vs Invoice tool. Compares ServiceTitan job material costs and gross margin against invoice totals, and cross-checks tech write-ups against billed line items to catch under-billing.

## How it works

- **Reads** happen straight from `data/jobs.json` and `data/invoice_checks.json` in this repo, served publicly via GitHub Pages. No login required to view.
- **Writes** (uploading a new Jobs report, saving an invoice check, a Pass/Fail override, clearing data) go through a Cloudflare Worker that holds a GitHub token server-side and commits the updated JSON files back to this repo. The browser never touches GitHub directly.

## Setup still needed: the Worker relay

The site currently loads and displays data, but the upload/override buttons won't do anything until the Worker is deployed and `WORKER_URL` is set at the top of `index.html`'s script.

The Worker needs a fine-grained GitHub PAT (Contents: Read and write, scoped to this repo) and should expose these endpoints, all POST, JSON body, committing to this repo on each call:

- `/api/materials-jobs` — `{ jobs: [...] }` — upsert each job into `data/jobs.json` by `jobNumber` (merge, don't replace the whole file)
- `/api/materials-invoice-checks` — `{ records: [...] }` — upsert each record into `data/invoice_checks.json` by `id`
- `/api/materials-invoice-status` — `{ id, manualOverride }` — update one record's `manualOverride` field (`"pass"`, `"fail"`, or `null`)
- `/api/materials-invoice-delete` — `{ id }` — remove one record from `data/invoice_checks.json`
- `/api/materials-jobs-clear` — `{}` — empty `data/jobs.json` back to `[]`

This mirrors the pattern already used for the Labor Productivity dashboard's Worker (`young-sound-f090.mattsbaker1980.workers.dev`) — either add new routes to that same Worker, or deploy a new one, then paste its URL into `WORKER_URL` in `index.html` and push.

## Data files

- `data/jobs.json` — array of job records parsed from the ServiceTitan "Jobs" custom report export
- `data/invoice_checks.json` — array of invoice discrepancy-check records, from bulk import or manual paste

Both start as empty arrays (`[]`) and are only ever written to by the Worker, never edited by hand.

## Job Review (`review.html`)

A per-job review page built from four ServiceTitan scheduled reports, read as real `.xlsx` files from Gmail
(an Outlook rule redirects `noreply@onservicetitan.com` to Gmail; Gmail's RAW message format keeps attachments byte-exact).

| Report | Used for |
|---|---|
| Materials vs Invoice | revenue, material / PO / total costs, margin, booking notes |
| Daily Job Detail | hours worked / paid / sold, estimate count, campaign, warranty/recall |
| Invoice Line Items with Description | tech write-up (invoice summary), billed items, PO + material cost per invoice |
| All Estimates – Daily (Creation Date, all statuses) | every option presented, with summary and who built it |

Daily flow (scheduled Claude task):
1. `sync/gmail_extract.js` — decode Gmail RAW messages, write the `.xlsx` attachments.
2. `sync/review_sync.js` — parse + merge into `data/review/YYYY-MM.json` (one record per job, bucketed by completion month), compute rule flags.
3. `sync/review_ai.js queue` → Claude reviews write-up vs invoice vs options per `sync/REVIEW_RUBRIC.md` → `review_ai.js apply`.

Rule flags: no options / one option, estimate with no description, PO costs on a $0 invoice, material cost with nothing billed,
billed with no write-up, negative margin, very short / unusually long visits. Claude flags: work done but not billed,
billed but not described, options gap, write-up issues.
