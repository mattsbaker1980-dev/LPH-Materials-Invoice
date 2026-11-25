# Reading supplier invoices for the Job Review

Used by the "LPH Job Review — Supplier Invoices" scheduled task. Goal: for each job PO that `node po_data.js todo`
lists, find the supplier's invoice PDF in Mindy's OneDrive and record what was billed.

## Where the PDFs are

Mindy's OneDrive `Invoices` folder, shared with Matt (Microsoft 365 connector):
`file:///b!NIcnIy3E-0aVmelPn1D8ed2XAEE8RmtNiendkeKxsxMOaqxQqUKeS5stPCK_fLGO/Invoices`
New invoices land loose in that folder and are later filed into vendor subfolders (Hajoca, Conestoga, HVAC, TS, …).

File names: `<Vendor> <supplier invoice #>  <our PO #>.pdf`, e.g. `Hajoca 39831117  95279.pdf`,
`HVAC 4194573.002  94717.pdf` (a `.002` invoice is a second shipment / back-order on the same PO — read it too).

## Finding the file for a PO

`sharepoint_search` with query `"<PO>.pdf"` (e.g. `95279.pdf`), limit 5. Keep only results whose **name ends with
` <PO>.pdf`** (two spaces before the PO in the name is normal). A plain number search (`95279`) misses many files —
always use `<PO>.pdf`. If nothing matches, record `{"po": "<PO>", "notFound": true}` (it is retried on later days;
the invoice may not have been saved yet).

## Reading it

`read_resource` on the result's `uri`. Record one object per PDF:

```json
{"po": "95279", "file": "Hajoca 39831117  95279.pdf", "vendor": "Hajoca", "invoiceNo": "S039831117.001",
 "invoiceDate": "11/24/26", "orderedBy": "Trent",
 "lines": [{"desc": "Liberty LE51A-2 1/2 HP submersible sewage pump 115V", "qty": 1, "unit": 520.19, "ext": 520.19}],
 "subtotal": 520.19, "tax": 31.21, "total": 551.40}
```

- `lines`: every product line — a short plain description (brand, model, size, what it is), quantity shipped,
  unit price, extended amount. Long runs of tiny identical fittings may be combined into one line
  ("Black steel nipples 1/2 and 3/4, close to 6in (22 pcs)") **as long as the amounts still add up**.
- `subtotal` = before tax and shipping; `tax`; `total` = amount due. The script checks that the lines add up to the
  subtotal and warns if not — fix the read rather than forcing numbers.
- Multi-page invoices: read all pages; the totals are on the last page.
- Home Depot / HD Supply invoices put the PO under "Purchase Order Number"; Thomas Somerville under "Cust Order No.";
  HVAC Distributors under "Customer Order Number"; Hajoca/Conestoga under "Customer P.O.". If the PO printed on the
  invoice doesn't match the PO you searched for, don't record it.
- Credits / returns: record them with negative `ext` and amounts.
- Copy numbers exactly; never guess a missing price.
