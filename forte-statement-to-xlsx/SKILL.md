---
name: forte-statement-to-xlsx
description: Convert Forte (ForteForex-related) KZT statement PDFs into an Excel file with transaction rows and a summary sheet where income/withdrawals/net exclude cancelled ForteForex orders. Use when the user provides a Forte statement PDF in PDF format and needs calculations in XLSX, especially when cancelled orders must be excluded by same-day matching of refunds (order cancelled + commission cancelled) to the initial debit.
---

# Forte Statement to XLSX

## When to use this skill
Use when the user asks to convert a Forte/ForteForex statement PDF (KZT) into an Excel file so they can calculate income and withdrawals, and the result must correctly exclude **cancelled ForteForex orders**.

## Core logic (what makes the net correct)
On each day in the statement, match cancelled ForteForex bundles as follows:

1. Identify the **principal refund** rows (positive KZT) containing `Снятие заявки по ForteForex`.
2. Identify the **commission refund** rows (positive KZT) containing `Отмена комиссии по сделке ForteForex`.
3. For each refund pair on the same day, compute the expected initial debit:
   - `expected_initial = -(principal_refund_amount + commission_refund_amount)`
4. Mark as cancelled the **initial debit** row (negative KZT) on that same day whose amount matches `expected_initial` (amount-to-values matching).

Then exclude all rows marked as cancelled bundles from the “clean” income/withdrawal/net totals.

## Workflow
1. Run the end-to-end converter script from this skill folder:
   - Example:
     - `PYTHONPATH=<skill_folder> python3 <skill_folder>/scripts/forte_statement_pdf_to_xlsx.py "<pdf_path>" -o "<out_xlsx_path>" --initial-balance-kzt <initial_balance_if_known>`
2. The script generates:
   - `transactions` sheet with a `is_cancelled_forteforex_bundle` flag.
   - `summary` sheet with:
     - income/withdrawals/net for KZT (all)
     - income/withdrawals/net excluding cancelled ForteForex bundles
     - initial balance input
     - expected final balance and final balance extracted from the PDF
3. Verify:
   - `Expected final balance` equals `Final balance from statement` (given the user’s initial balance).

## Notes / limitations
- This works best on PDFs like your Forte/ForteForex statements where the order lives within a single day.
- If the Excel shows a “broken formula”, ensure the file is recalculated in Excel; the script uses bounded ranges rather than whole-column references.

