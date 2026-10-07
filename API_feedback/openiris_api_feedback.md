# OpenIRIS Public API: feedback from University of Helsinki

**Spec reviewed:** `openapi-preview.yaml`, preview-v13 (2026-09-22), plus the ERP charge sync recipe and changelog on docs.openiris.io
**From:** Harri Jäälinoja, Light Microscopy Unit (LMU), HiLIFE, University of Helsinki
**Draft date:** 2026-09-28

> **Internal note (delete before sending):** confirm whether this goes out on behalf of LMU only
> or LMU + BIU + EMBI. See also the *Open questions for us* section below.


## 0. Open questions for us at UH (delete before sending)

- What is the maximum length and format of WBS codes at UH, for LMU and for BIU? This would
  let us give a concrete number in 6.2.
- HiLIFE HUS: how should it be treated in SAP (same as internal?).
- Charge and product comments: our notebooks use the charge comment ("Comments
  (charge)", e.g. why a discount was given) and the product comment and purchase date (who
  bought a product), but the API doesn't expose them. Does SAP, or the controller, need them
  in or alongside a posting? If yes, ask OpenIRIS to add them to Charge. Note that night-time
  rates will be handled as a `discount_percent` agreed with the user, so the reason may matter.
- `wbs_override.csv`: today we can override a group's WBS locally. With the API, every such
  override has to be made in OpenIRIS instead. Is that acceptable?
- Is the €15 minimum-invoice rule and the "prepaid" handling (detected from the request title)
  something we want OpenIRIS to support, or keep on our side?

## 1. Background: how we bill today

We are core facilities at the University of Helsinki, invoicing quarterly through SAP. Today:

1. In the OpenIRIS UI we select the period's charges and add them to an invoice draft, then
   export the invoice.
2. We run a set of checks on the export: missing price type, missing WBS/remit code, missing
   PI or requester, cancellations with comments, discounts, waived charges, overlapping
   bookings, etc.
3. If a check fails we fix the data in OpenIRIS, delete the draft and start again.
4. When everything is clean, we split the invoice per group and WBS into attachments and email
   them to our controller together with a summary spreadsheet.

The goal is to replace steps 1–4 with an API integration into the university's SAP. The end
result must stay the same: each WBS is charged the correct amount, and there is a traceable
per-line breakdown behind every posting.

## 2. API draft progress and summary of proposals and open issues

These features were added in v8 and v13:

- **Cost center sync (`POST`/`PATCH /cost-centers`).** Lets us keep valid WBS codes in step with SAP.
- **Charge sign-off (`confirmed_at`).** This is a starting point, invoices can be built from confirmed charges.

The points below are what remains: 3.1 is outside the API; 3.2 and 3.3 are related to the SAP side.
3.4. and 3.5 were added after a meeting with Julia where it was concluded that the way forward is via 
the `/invoice` endpoint.

Section 4 shows our current understanding of how the billing process would work via the API. 
This is the part that we should carefully go through with the OpenIRIS team.
4.3 and 4.4 are requests that would allow us to replace the current charge confirmation 
process that relies on creating/deleting/recreating 
an invoice in OpenIRIS by a new one that queries the API for the charges in the invoicing period.

Section 5 explains the checks we run currently. Section 6 has other items found in the spec by Claude.ai.

## 3. Known priority items

### 3.1 Use provider holidays in pricing

OpenIRIS UI allows provider admins to select holidays, but to our knowledge these holidays are not used when making pricing decisions. Our policy is to apply off-hours prices on national holidays, and currently we need to apply this outside OpenIRIS. This item is not directly related to the API, but has bearing on how we can correct the charges in OpenIRIS prior to a billing run.

*Agreed with Julia 2026-10-07*:
- provider admin marks provider holidays (existing feature)
- add checkbox "include public holidays" in Off-hours price item

### 3.2 Price type on the charge

Neither `Charge` nor `ChargeLineItem` says which price type (e.g. *HiLIFE internal*,
*HiLIFE External*, *HiLIFE Commercial*) the charge was priced under. We use it to separate
internal cost recovery from invoicing external and commercial customers, which are different
documents in SAP. Requests:

- Add `price_type_id` (and name) to Charge, or at least to each line item.
- Document the response of `/price-types`, `/price-types/{id}` and `/price-types/{id}/items`
  (currently just `200 OK`).

Internal invoices and external invoices are two separate interfaces.
Internal invoices are accounting memos, not real invoices. But they still have an approval 
workflow in SAP, so they go through verification (asiatarkastus) and approval before they are posted to the accounts.
We just need some piece of data that distinguishes these two datasets. 
Based on that, the integration routes the data to the different interfaces.

### 3.3 People: verifiers and contacts

Each posting needs a verifier (in Finnish *asiatarkastaja*): the person who approves the cost
on the payer's side. The verifier belongs to the **group**, not to an individual request or WBS.
A group typically has several requests and often several WBS codes, but the same verifier.

Today the only place to collect it is a field in the service request form. That is a poor fit:
we ask the same group for the same verifier again and again, and the answers can disagree. When
the form has no verifier, we fall back to listing all request owners in the group who charged
that WBS. That list is only an approximation.

Requests:

- Add a **verifier** (user reference, or name + email; possibly more than one) to `Group`,
  maintained by the group or by facility staff in the UI. It should be readable from
  `GET /groups/{id}` and embedded on charges, or at least reachable in one step from
  `charge.group_id`.
- Add user role **verifier** that allows to check and verify charges.
- As a fallback, the request owner (`Request.user_id`, via `charge.request_id`) already covers
  what we do today. `GET /requests/{id}/form-submission` / `FormSubmission.data` would still be useful
  for other form fields, but is not the right long-term home for the verifier.
- `Group` has `contact_email` but no group heads/PIs. Please add them (e.g. a `role` on
  `GroupMember`).

### 3.4 Improve cost centers (WBS)

Cost centers should have 
- start date
- end date
- validity (is_active)

It was proposed that WBSs continue to be imported manually in Iris, but an API endpoint is added that allows
to get the list of WBSs and the set the validity parameters. This allows SAP integration to synchronize the
status of the WBSs known to Iris.

### 3.5 Improve `/invoice` endpoint

Recording individual charges in SAP is not feasible, we should record totals as we've done so far.
Using the `/charges` endpoint would mean that we would have to build the logic of combining charges 
somewhere. UH integration does not do logic, and building it by providers would be cludgy, so the best 
way forward is to improve the `/invoice` endpoint of the API.

These are the minimum required data for SAP:

*External invoices*

- SAP customer ID
- Invoice date

The following can occur n times per invoice:

- Product being invoiced, SAP ID (if there is only one product, it can be hard-coded in the integration)
- WBS to which the revenue is posted
- Quantity invoiced (can also be a constant 1, in which case the total price is given as the unit price)
- Unit of measure (can be fixed, e.g. "pcs")
- Unit price (net); SAP calculates the taxes
- Description text for the invoice line (this can be taken from SAP, but it may not be descriptive enough)
- Reference required by the customer
- Some reference number that lets you link the invoice the customer receives to the billing transaction in OpenIRIS. Does the system have a unique invoice number?
- Any additional descriptions needed, either per invoice or per line

*Internal invoices*

- Invoice date
- Header and reference information: profit center / WBS / orderer, the necessary OpenIRIS reference, etc. as description text
- Reference (the sending system's invoice number or similar)
- Invoicing (seller's) profit center
- Customer's profit center

The following can occur n times per invoice:

- Invoice line description
- Quantity invoiced (the unit can be sent as a constant, e.g. "pcs")
- Unit price invoiced (excluding tax)
- Invoice line total (quantity × unit price)
- Revenue GL account
- Revenue WBS


## 4. Workflow functions and consistency

Key functions:
- `/charges/filter[...]` (provider lists charges of billing period, runs checks)
- `/charges/{id}/confirmation`, `/charges/confirmations` (provider confirms charge(s) are good-to-go)

*TODO: add `/invoice` endpoint calls* 
- invoice is created in OpenIRIS
- invoice is marked good-to-go
- integration reads `/invoice`
- SAP gets invoice
- invoice is marked as exported to SAP

An internal invoice can have only one paying WBS per invoice. So the charges should be 
bundled by paying WBS into one invoice. You could of course send them charge by charge, 
but wouldn't that get too fragmented?

External invoices should preferably be bundled by customer/reference. By reference 
I mean the reference you have received from the external customer and that they want to appear on the invoice.

### 4.1 What happens when a charge changes after sign-off or export?

- If OpenIRIS recalculates a charge after it was confirmed (`recalculated_at` > `confirmed_at`),
  is the sign-off kept, withdrawn automatically, or flagged? We suggest withdrawing it, or at
  least exposing a `changed_since_confirmation` flag.
- After a charge is `exported`: is it locked? If not, how are corrections and credits
  represented, so the ERP side can post a reversal?
- `updated_at` is NULL on ~75% of rows. Until that is backfilled, delta sync is unreliable.
- Please publish the webhook event catalogue. At minimum we would use `charge.updated`,
  `charge.recalculated`, `charge.confirmation_withdrawn` and `charge.exported`.
- Add `confirmed_by` (user or key) to the confirmation, for audit.


### 4.3 Line items in bulk

`GET /charges/{id}/line-items` is per charge. A quarter for one facility is ~1,500 charges.
Requests:

- Support `expand=line_items` on `GET /charges`, and/or
- Include line items in `POST /charges/bulk-export` output.
- Document the possible values of `price_item_code` (regular usage, off-hours, cancellation,
  training, …).

### 4.4 Filters on charges and bulk export

Please add:

- `filter[period_start]` / `filter[period_end]` (with `gte`/`lt`), to select the charges of a
  billing period by when the usage happened. Only `created_at`, `updated_at` and
  `confirmed_at` can be filtered today.
- `filter[group_id]`, `filter[is_waived]`
- `filter[status]` with plain `eq` (only `filter[status][in]` is listed)
- The same filters on `POST /charges/bulk-export`, which today accepts only `from`/`to`/`status`.
  Most useful for us are provider, period, `confirmed` and `cost_center.provider_code`.
- `filter[id][in]` on `/users` and `/groups`, so the people and groups behind a set of charges
  can be fetched in one call instead of one call each (see 5.5).



## 5. Current checks: can the API support them?

Before each billing run our tooling checks the exported invoice and writes each finding to a
spreadsheet (`InvoiceNN__<check>.xlsx`), which we review and fix in OpenIRIS. With an API
integration the same checks would run on the charges before sign-off. The table shows whether
the draft API provides the data for each one.

**API support:** ✅ yes · 🟡 partly, or only with extra lookups · ❌ no · ➖ no longer needed
(see the notes below the tables)

### 5.1 Checks we review every billing run

"Found in one quarter" shows what each check caught on one real quarterly invoice from one of
our facilities (~1,550 lines). "—" means not counted here.

| Check (output file) | What it finds | API support | How / what is missing | Found in one quarter |
|---|---|---|---|---|
| `price_type_missing` | Charges with no price type | ❌ | No price type on Charge or line item (3.2) | 0 |
| `wbs_multiple_price_types` | A WBS charged under more than one price type | ❌ | Same as above | 1 WBS, 52 rows: bookings priced *HiLIFE internal*, a product priced *Default* (products always get *Default*, whatever the group's price type) |
| `request_id_missing` | Charges not linked to a request | ✅ | `charge.request_id` | 35 rows, 1 group |
| `group_or_wbs_missing` | Charges with no group or no WBS | ✅ | `charge.group_id`, `charge.cost_center.provider_code` | 0 |
| `remit_code_missing` | WBS with no remit code (H-code) | 🟡 | `remit_code` is on `/cost-centers`, not on the charge; one lookup per cost center (6.2) | 0 |
| `cancellation_reasons` | Cancellation charges with a booking comment | 🟡 | `Booking.comments` ✅; identifying a cancellation needs documented `price_item_code` values (4.3) | — |
| `pi_email_missing` | Groups with no PI email | 🟡 | `Group.contact_email` only; no group heads (4.4) | 77 rows, 1 group |
| `products` | Product purchases, listed for review | 🟡 | `source = product` ✅; product comment and purchase date not exposed | — |
| `overlapping_bookings` | Overlapping bookings on the same instrument | ✅ | `GET /bookings` with `resource_id`, `start`, `end`, `is_cancelled` | — |
| `prepaid` | Charges on prepaid requests | 🟡 | Only by matching "prepaid" in `Request.name`; a structured flag would be better | — |
| `totals_by_group_and_wbs_with_verifiers` | Totals per group, remit code and WBS, with verifier | 🟡 | Totals ✅ by grouping charges; remit code needs a lookup; verifier missing (3.3) | — |

### 5.2 Supporting checks and listings (debug output)

| Check (output file) | What it finds | API support | How / what is missing |
|---|---|---|---|
| `requester_missing` | Charges with no requester (the request owner) | ✅ | `charge.request_id` → `GET /requests/{id}` → `Request.user_id`; one call per request, or `expand` if supported. Request *participants* are not requesters: they share the owner's access and WBS but don't own the request |
| `verifiers`, `error_multiple_verifiers_per_WBS` | Verifier per WBS, and conflicting answers from request forms | ❌ | Form data not yet implemented; we'd rather have a verifier on the group (3.3) |
| `cancellations` | All cancellation charges | 🟡 | Needs documented `price_item_code` values (4.3) |
| `discount`, `discount_factor` | Charges with a discount | ✅ | `discount_percent`; the reason (charge comment) is not exposed |
| `waived` | Waived charges | ✅ | `is_waived`, `status = waived` |
| `trainings` | Training charges | ✅ | `ChargeLineItem.is_training` (line items per charge, 4.3) |
| `cost_center_code_int` | WBS code format (numeric at LMU) | ✅ | `cost_center.provider_code` |
| `totals_by_group_and_wbs_*`, `totals_by_resource_*` | Totals for reconciliation | 🟡 | Grouping by resource needs `booking_id` → booking → `resource_id`; not on Charge directly |
| Staff and test bookings (`BIU_bookings`, `LMU_bookings`) | Facility staff and test groups/instruments, excluded from billing | ✅ | Filter by `group_id` / `resource_id` |

### 5.3 Checks that will no longer be needed

These support our local price-list recalculation, which we are phasing out. Night-time rates
will be a discount agreed with the user, and cancellation prices are defined per price type in
OpenIRIS.

- `needs_split_2`, `needs_split_3`, `night1_and_night2*`, `regular_price_during_holidays`:
  bookings spanning prime / off-hours / night bands
- `duration_changed`, `charge_changed`: differences between OpenIRIS and our recalculation
- `discount_with_split`: a discount on only some of a booking's split lines. This is probably
  moot if `discount_percent` applies to the whole charge; please confirm.


### 5.5 How many API calls would the checks take?

With the sign-off workflow, we would run the checks, fix problems in OpenIRIS, and run the
checks again until everything is clean. So what matters is the cost of **one full check run**,
repeated several times per billing period.

We estimated this for one real quarterly invoice from one of our facilities: about 1,360
charges (1,359 bookings, 178 of them split over several lines, plus 1 product) covering 126
requests, 62 groups, 68 WBS codes, about 165 people and 16 instruments. We assumed one charge
per booking, with split lines as line items; if split lines are separate charges, the
per-charge rows below grow to about 1,550.

With the current spec, one check run takes about **1,700 read calls**, almost all of them
per-charge line-item lookups and per-person/per-group lookups. That is roughly 3 minutes of the
default 600 calls/minute limit for every run. Recording the SAP postings afterwards adds another
~1,360 calls, one per charge. This is workable for one quarterly run of a small facility, but
it scales linearly with the size of the facility and the number of check-and-fix rounds.

With the changes suggested in this document, the same run takes **about a dozen calls**, and
recording the postings takes 3.

| What | Calls, current spec | Calls, with suggested changes | Change needed |
|---|---|---|---|
| Charges | 3 | 3 | — (500 per page) |
| Line items | ~1,360 | 0 | `expand=line_items` on `GET /charges` (3.3) |
| Bookings (title, comments, instrument, overlap check) | ~3 | ~3 | — (listed by period) |
| Requests (owner, title, "prepaid") | 2 | 2 | — (list all of the facility's ~900 requests) |
| People (user and requester names) | ~165 | ~1 | `filter[id][in]` on `/users` (3.4) |
| Groups | ~62 | ~1 | `filter[id][in]` on `/groups` (3.4) |
| Cost centers (remit code) | 1–2 | 0 | `remit_code` and `name` embedded in `charge.cost_center` (6.2) |
| Instruments | 1 | 1 | — |
| Verifier from request forms (once implemented) | ~126 | 0 | Verifier on the group (3.5) |
| **Check run, total** | **~1,700** | **~10–15** | |
| Sign-off | 3 | 3 | — (bulk, 500 per call) |
| Recording SAP postings | ~1,360 | 3 | Bulk export acknowledgement (4.2) |

## 6. Problems in the current spec

### 6.1 Write operations are declared under the read-only `billing:read` scope

These endpoints change billing data but only require `billing:read`:

- `POST /charges/{id}/transitions` (→ `invoiced` / `waived` / `cancelled`)
- `PATCH /invoices/{id}` and `DELETE /invoices/{id}` (void)
- `POST /invoices/{id}/exports`

*Why it matters:* a key issued only to read charges into SAP could waive charges or void
invoices. We suggest moving these to `billing:write`.

(`POST /charges/{id}/exports` is also under `billing:read`. That may be intentional, since the
recipe says a read key records exports, but it would be good to state it explicitly.)

### 6.2 Cost center codes: format and length

`provider_code` / `/cost-centers.code` is limited to **≤10 alphanumeric** characters. The spec
says this is the "SAP cost center length limit", yet describes the field as "SAP WBS elements,
cost centers". These are different SAP objects: WBS elements are longer and often contain
separators. The spec's own example, `provider_code: "MDC-IMG-01"`, would fail the
`^[A-Za-z0-9]+$` pattern.

At our university the two facilities differ: one uses purely numeric WBS codes, the other also
uses codes with other characters. Requests:

- Relax the constraint to match SAP WBS elements, or keep cost center and WBS element as
  separate fields.
- Clarify `external_account_assignment` on Charge: who writes it, and is it meant for the WBS
  element?
- Embed `name` and `remit_code` in the charge's `cost_center` object (today only
  `provider_cost_center_id` is there, so every charge needs a second lookup).

### 6.3 Local timestamps

- The example shows `invoice_date_local: 2026-03-10T01:00:00Z`: a local time with a `Z`
  (UTC) suffix. Please use a real offset (`+02:00`) or a zone-less local time, and document
  which.
- `ChargeLineItem.period_start` / `period_end` have no `_local` variants. Price bands are
  defined in local time, so these would be useful too.

### 6.4 Other inconsistencies

- `Idempotency-Key`: max 64 characters and cached 24 h in the yaml, versus 128 characters and
  "survives restarts" in the docs.
- The `ChargesList` example uses `cost_center.organization` / `.group`, while the schema
  defines `organization_code` / `group_code`.
- `/invoices` description says the Invoice entity lacks `subtotal`, `tax_total`, … but the
  `Invoice` schema already has them.
- `POST /charges/{id}/exports` responses: the yaml documents only `201` (recorded). The ERP
  charge sync recipe also describes `200` (this SAP document was already recorded for the
  charge, so a retry is safe) and `409` (the charge is waived and should not have been
  posted). The dedup protection that the whole ERP workflow relies on is visible only in the
  recipe. An integrator generating a client from the yaml would treat `200` and `409` as
  unexpected, so both belong in the spec.
- `POST /charges/{id}/exports` request body: the yaml makes `exported_at` required, but the
  recipe's example request omits it. One of the two needs correcting, so that the documented
  example doesn't fail validation.


---

