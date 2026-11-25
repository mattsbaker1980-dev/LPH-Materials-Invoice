# Reading supplier invoices for the Purchase Orders page

Used by the "LPH Purchase Orders — Supplier Invoices" scheduled task. Goal: read every NEW supplier invoice PDF that
Mindy saves to her OneDrive `Invoices` folder, record what it says, and let `po_review.js` match it to a job.

## 1. What's already recorded

`node po_review.js known` prints `{count, highestPO, files:[...]}`: the file names already read. Skip those.

## 2. List the folder

Microsoft 365 connector, `read_resource` on
`file:///b!NIcnIy3E-0aVmelPn1D8ed2XAEE8RmtNiendkeKxsxMOaqxQqUKeS5stPCK_fLGO/Invoices`
(200 entries per page; if the last line says to re-call with startPage 2, do so). Only the LOOSE files at the top level
matter. The vendor subfolders (Hajoca, APR, …) and "2025 Invoices" hold older, already-filed invoices: never open them.

File names: `<Vendor> <supplier invoice #>  <our PO #>.pdf`, e.g. `Hajoca 39831117  95279.pdf`. Some end in a word
instead of a PO number (`Stock`, `Consumables`, `Shop`, `TRK94`, `Derrick Tools`, `7777`): those are stock/shop
orders. Read them too (they're short); they go to a "Stock / shop" list.

Pick the loose `.pdf` files that are NOT in the `known` list. Skip 0-byte files (Mindy is still saving them; they'll
be there next run). Take at most **40 per run**, highest PO number first, then the stock/shop ones. If more are
waiting, say how many in the report; the next run continues.

## 3. Read each PDF

`read_resource` on the file's uri (all pages). Record one object per PDF:

```json
{"file": "Hajoca 39823188  95258.pdf", "uri": "file:///b!NIcn.../01XD3WOWM6XHAXKD4DJNCKVY62D3WPT47W",
 "vendor": "Hajoca", "invoiceNo": "S039823188.001", "po": "95258",
 "invoiceDate": "11/23/26", "orderDate": "11/23/26", "orderedBy": "TYLER", "jobRef": "",
 "lines": [{"desc": "Gerber 28-380 Ultra Flush pressure-assist tank 1.6gpf", "qty": 1, "unit": 197.90, "ext": 197.90}],
 "subtotal": 435.11, "tax": 26.10, "total": 461.21}
```

- `file` and `uri`: exactly as in the folder listing (the uri is the item's `file:///…/<itemId>` link).
- `total` = the **amount due** (after tax and freight). This is what gets matched to the job, so copy it exactly.
  `subtotal` = before tax/freight, `tax` = sales tax.
- `orderDate`: the order date if the invoice shows one, else the invoice/ship date. Hajoca/Conestoga/APR/Schaedler/
  Yale show both.
- `orderedBy`: the "Ordered by" / "Placed by" name as printed (Home Depot: the cardholder, e.g. "Mindy Ault").
- `jobRef`: whatever is printed in the invoice's **Job name / Job number / Reference / Release** field, if it isn't
  just our company name or the PO number again (e.g. "JOSH R. - OKANYA", "Goyda"). It often names the customer.
- `lines`: every product line: a short plain description (brand, model, size, what it is), quantity shipped, unit
  price, extended amount. Long runs of tiny identical fittings may be combined into one line as long as the amounts
  still add up. The script rejects an invoice whose lines don't add up to the subtotal (within 5 cents). Fix the read;
  never force the numbers.
- Credits / returns: negative amounts.
- Copy numbers exactly; never guess a missing price.
- **Scanned PDFs**: if `read_resource` says no text could be extracted (RE Michel, Johnstone and Sid Harvey send
  scans), record only `{"file": "...", "uri": "...", "webUrl": "<webUrl from the result, if given>", "unreadable": true}`.
  Do not try to OCR or guess.

## 4. Save

Write the array to `results.json`, check it parses, then `node po_review.js add results.json`. It prints
`REJECTED …` for any invoice that failed a check (fix and re-add just those) and a `DONE …` line.
If there were no new files, run `node po_review.js rematch` instead (new job costs arrive every morning, so open
invoices can match later).
