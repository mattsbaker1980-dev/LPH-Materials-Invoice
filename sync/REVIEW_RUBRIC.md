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
- Installing a **customer-supplied** part (filter, batteries, pad) during a visit.
- Booking notes that differ from what the tech found — the write-up is what matters.
- If the write-up says the customer declined everything but an estimate is marked Sold (or the reverse),
  that IS worth a `writeup` / low issue.

`writeup` (**low**) — the write-up is too thin to tell what was done (e.g. "done", "fixed it"),
or it contradicts itself or the invoice (e.g. says "system is not heating" but no follow-up).

`other` — anything else clearly worth a manager's attention. Use sparingly.

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
   "note": "Otherwise matches."}
]
```

Keep each `text` to one plain sentence a manager can act on: name the part or work and what's missing.
At most 4 issues per job. Every job in the queue must get a result.
