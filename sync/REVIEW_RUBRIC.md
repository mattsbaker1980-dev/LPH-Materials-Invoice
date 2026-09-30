# Claude job review — rubric

You are reviewing completed service jobs for a plumbing / HVAC / electrical company
(LPH). For each job in `queue.json` you get the tech's invoice write-up (`writeUp`),
the billed line items (`items`, `[M]` = material), the estimates/options the tech built
(`estimates`), and booking notes from the office. Management wants to know, quickly,
whether **what was billed matches what the tech says was done, and vice versa**, and
whether the customer was given real options.

Read each job and return one result object per job. Be conservative: only raise an
issue when the text clearly supports it. A clean job is `"verdict": "ok"` with no issues.
Do not invent facts.

## What counts as an issue

`not_billed` (usually **high**) — the write-up says the tech *actually did* a repair,
replacement, install, or extra service today, and nothing on the invoice covers it.
- Only work that was done, in the past tense, by the tech ("replaced the capacitor",
  "installed a new wax ring", "cleared the main line").
- NOT recommendations, quotes, options, or future work ("recommend replacing the panel",
  "provided options to replace the water heater").
- NOT things the homeowner or someone else did ("homeowner recently replaced the filter").
- NOT routine steps already included in the billed service. A system check / maintenance
  visit includes inspecting, testing, measuring, cleaning, tightening, lubricating,
  flushing drain lines, and a standard filter change when a filter is on the invoice or
  the plan covers it. A drain service includes running the cable. Diagnostics include testing.
- Warranty-parts lines are fine if the write-up or booking notes explain the warranty repair.

`billed_not_described` (**low**, **high** if the item is > $300) — an invoice line for
work or a part that the write-up never mentions, so there is no record of why it was billed.
Membership sign-ups, fees, discounts and the base system-check line don't need to be described.

`options` (**low**) — the write-up describes problems or recommendations but NO estimate was built
for them, or a major repair/replacement (roughly $1,000+) was offered as a single take-it-or-leave-it
option. One option for one small, simple fix is fine. Don't raise this for sold-work or install jobs.

## Not issues (don't flag)

- A **sold estimate that isn't on this job's invoice** — sold work is usually done and billed on a
  separate follow-up job.
- If the queue item has `soldWorkBilledOn`, that sold estimate WAS billed on the follow-up job listed.
  The write-up on this visit may only describe the proposed work — that is normal, not a `writeup` issue.
- If the queue item has `fromEstimate`, this job is the follow-up that carried out work sold on an
  earlier job; judge the invoice against that estimate and this write-up.
- Installing a **customer-supplied** part (filter, batteries, pad) during a visit.
- Booking notes that differ from what the tech found — the write-up is what matters.
- If the write-up says the customer declined everything but an estimate is marked Sold (or the reverse),
  that IS worth a `writeup` / low issue.

`writeup` (**low**) — the write-up is too thin to tell what was done (e.g. "done", "fixed it"),
or it contradicts itself or the invoice (e.g. says "system is not heating" but no follow-up).

`other` — anything else clearly worth a manager's attention. Use sparingly.

## Materials gap (only for jobs that have a `materialsGap` block)

These jobs had parts bought on a purchase order (`purchasedOnPO`, a dollar total — the PO's
line items are NOT available) but little or nothing recorded as materials used on the invoice
(`recordedMaterials`, `recordedMaterialCost`). Job costing is off by roughly the difference.

Read the write-up (and the billed service lines) and list, in `likelyUnrecorded`, the physical
parts/materials the tech says were installed or used that do NOT appear in `recordedMaterials`.
- Name them the way a warehouse person would look for them ("SDR 35 sewer pipe",
  "crushed stone", "main shut-off valve", "50 gal electric water heater").
- Include equipment billed as a service line (e.g. a water heater or furnace) if it was
  probably bought on the PO — say so in `gapNote`.
- Only list things the write-up actually mentions. If it doesn't say what was used, return an
  empty list and set `gapNote` to "Write-up doesn't say what was used — check the PO."
- `gapNote`: one or two sentences, e.g. "About $3,600 bought on PO; only small fittings recorded.
  The pipe, stone and fittings for the sewer run and interior drains aren't on the invoice as materials."

Always copy `purchasedOnPO` into `materialsGapFor` for these jobs. This is a guide for the
manager checking the PO, not an accusation — don't also raise it as a `not_billed` issue unless
the write-up describes billable work that isn't priced anywhere on the invoice.

## General Time (only for jobs that have a `generalTime` block)

Techs sometimes bill a "General Time" / Grid Task line (`gtCode` like GT-0075-0020 = 0.75 labor hours +
$20 material allowance, picked by the tech) instead of real pricebook tasks. For each line in
`generalTime`, read `techDescription` and decide which pricebook task(s) from `candidates` cover the
work described.
- Match the same work on the same kind of equipment (a furnace gas valve is not a water-heater gas valve).
- If the description lists several separate repairs (e.g. "replaced pressure switch, inducer and ignitor"),
  return one task per repair in `tasks`. Use `quantity` for "2 valves" etc. (a task that already says
  "(2)" counts as one).
- Return no tasks when none of the candidates genuinely fits, and set `reason`:
  - `"custom_work"` — real work that the pricebook doesn't have a task for (repipes, custom runs,
    unusual installs). General Time is reasonable here.
  - `"too_vague"` — the description doesn't say what was done ("Custom Solution", "per quote") —
    a manager can't verify the charge.
  - `"not_work"` — it's a diagnosis, note or warranty placeholder rather than work.
- Lines whose `line` starts with `est:` are ESTIMATE OPTIONS (not invoice lines) that were priced exactly at a
  General Time price; `techDescription` is the option's name and summary. Match them the same way — the
  question is which pricebook task(s) the tech should have built the option from.
- `confidence`: "high" when the task(s) clearly are the same work, "medium" when close, "low" when unsure.
- Don't compute prices — the tool looks them up. Don't raise General Time as an `issues` entry.

Return, per job: `"generalTime": [{"line": "<copied>", "tasks": [{"taskCode": "IM23870", "quantity": 1}, ...],
"reason": null, "confidence": "high", "note": "Furnace ignitor, pressure switch and inducer"}]` — one entry
for every line in the block (use `"tasks": []` plus a `reason` when nothing fits).

## Output

Write a JSON array to `results.json`. Copy `jobNumber` and `fingerprint` exactly from the queue.

```json
[
  {"jobNumber": "123456789", "fingerprint": "abc123", "verdict": "ok", "issues": []},
  {"jobNumber": "987654321", "fingerprint": "xyz789", "verdict": "issues",
   "issues": [
     {"type": "not_billed", "severity": "high",
      "text": "Write-up says the tech replaced the run capacitor; no capacitor or repair line on the invoice."}
   ],
   "note": "Otherwise matches."},
  {"jobNumber": "336230547", "fingerprint": "k2j9x", "verdict": "ok", "issues": [],
   "materialsGapFor": 3662.5,
   "likelyUnrecorded": ["SDR 35 sewer pipe", "Sch 40 PVC pipe and fittings (house trap, interior drains)", "Two-way clean-out", "Crushed stone bedding", "J-hooks", "Main shut-off valve", "Rheem 50 gal electric water heater (billed as a service line)"],
   "gapNote": "About $3,660 bought on PO; only $17 of small fittings recorded. The sewer-run and interior drain materials aren't recorded as used."}
]
```

Keep each `text` to one plain sentence a manager can act on: name the part or work and what's missing.
At most 4 issues per job. Every job in the queue must get a result.
