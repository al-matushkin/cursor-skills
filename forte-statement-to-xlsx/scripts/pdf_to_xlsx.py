#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional


DATE_RE = re.compile(r"^(?P<date>\d{2}\.\d{2}\.\d{4})(?P<rest>.*)$")
AMOUNT_RE = re.compile(
    r"(?P<amount>[+-]?\d+(?:\.\d+)?)\s*(?P<ccy>[A-Z]{3})\b", re.IGNORECASE
)
ORIG_AMOUNT_RE = re.compile(
    r"\(\s*(?P<amount>[+-]?\d+(?:\.\d+)?)\s*(?P<ccy>[A-Z]{3})\s*\)"
)


def _clean_line(s: str) -> str:
    s = s.replace("\u00a0", " ")
    return re.sub(r"\s+", " ", s).strip()


def _is_noise_line(s: str) -> bool:
    if not s:
        return True
    if s.startswith("-- ") and " of " in s:
        return True
    if s.lower().startswith("сформировано в интернет банкинге".lower()):
        return True
    if s.lower().startswith("реквизиты:".lower()):
        return True
    if s.lower().startswith("контактные данные:".lower()):
        return True
    if s.lower().startswith("выписка по карточному счету".lower()):
        return True
    if s.lower().startswith("дата выписки:".lower()):
        return True
    if s.lower().startswith("детализация выписки".lower()):
        return True
    if s.lower() == "дата сумма описание детализация":
        return True
    return False


@dataclass
class Txn:
    date: str
    amount_kzt: Optional[float]
    ccy: Optional[str]
    orig_amount: Optional[float]
    orig_ccy: Optional[str]
    description: str
    details: str


def _parse_amount(text: str) -> tuple[Optional[float], Optional[str]]:
    m = AMOUNT_RE.search(text)
    if not m:
        return None, None
    return float(m.group("amount")), m.group("ccy").upper()


def _parse_orig_amount(text: str) -> tuple[Optional[float], Optional[str]]:
    m = ORIG_AMOUNT_RE.search(text)
    if not m:
        return None, None
    return float(m.group("amount")), m.group("ccy").upper()


def iter_pdf_lines(pdf_path: Path) -> Iterable[str]:
    import pdfplumber  # type: ignore

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for raw in text.splitlines():
                yield _clean_line(raw)


def parse_transactions(lines: Iterable[str]) -> list[Txn]:
    txns: list[Txn] = []
    cur: Optional[Txn] = None
    pending_lines: list[str] = []
    in_table = False

    def flush():
        nonlocal cur
        if not cur:
            return
        cur.description = _clean_line(cur.description)
        cur.details = _clean_line(cur.details)
        txns.append(cur)
        cur = None

    for line in lines:
        # Don't treat header/account metadata as txn content.
        if not in_table:
            if line.lower().startswith("дата сумма описание детализация".lower()):
                in_table = True
            continue

        # Once we hit the end markers, stop consuming table content.
        if line.lower().startswith("сформировано в интернет банкинге".lower()):
            in_table = False
            pending_lines = []
            flush()
            continue

        if _is_noise_line(line):
            continue

        m = DATE_RE.match(line)
        if m:
            flush()
            date = m.group("date")
            rest = _clean_line(m.group("rest"))

            # pdfplumber may output some columns first (amount/details) and the date later.
            # Carry any "pending" pre-date lines into this transaction.
            carried = " ".join(pending_lines).strip()
            pending_lines = []

            amt, ccy = _parse_amount(rest)
            oamt, occy = _parse_orig_amount(rest)
            desc = rest
            det = ""

            if carried:
                # Prefer the first amount found in carried text if the date line doesn't include it.
                if amt is None or ccy is None:
                    amt2, ccy2 = _parse_amount(carried)
                    if amt2 is not None:
                        amt, ccy = amt2, ccy2
                if oamt is None or occy is None:
                    oamt2, occy2 = _parse_orig_amount(carried)
                    if oamt2 is not None:
                        oamt, occy = oamt2, occy2

                # Whatever remains is useful context (merchant / place / etc.)
                det = carried

            cur = Txn(
                date=date,
                amount_kzt=amt,
                ccy=ccy,
                orig_amount=oamt,
                orig_ccy=occy,
                description=desc,
                details=det,
            )
            continue

        # Sometimes an amount/details line for the *next* transaction appears before its date.
        # If we're currently inside a txn and see such a line, flush and carry it forward.
        if cur is not None:
            looks_like_next = False
            if line.lower().startswith("снятие наличных".lower()):
                looks_like_next = True
            elif AMOUNT_RE.search(line) and not line.startswith("("):
                # Common continuation prefixes we *don't* want to split on.
                low = line.lower()
                if not (low.startswith("денег") or low.startswith("k,") or "mcc:" in low):
                    looks_like_next = True

            if looks_like_next:
                flush()
                pending_lines = [line]
                continue

        # If we see plausible txn text before a date (common with column-ordered extraction),
        # stash it so the next date line can consume it.
        if cur is None:
            if AMOUNT_RE.search(line) or line.lower().startswith("снятие наличных".lower()):
                pending_lines.append(line)
            continue

        # Continuation line: append to current txn
        if cur:
            # Some PDFs wrap "Снятие наличных денег" across lines; keep it in description.
            # If we haven't captured amount yet, try to extract it from continuation.
            if cur.amount_kzt is None or cur.ccy is None:
                amt, ccy = _parse_amount(line)
                if amt is not None:
                    cur.amount_kzt, cur.ccy = amt, ccy

            # Original amount (e.g. foreign currency) can appear on continuation lines
            # even when the main amount is already known.
            oamt, occy = _parse_orig_amount(line)
            if oamt is not None:
                cur.orig_amount, cur.orig_ccy = oamt, occy

            if cur.details:
                cur.details += " " + line
            else:
                cur.details = line
        # If we get content outside txn blocks, ignore it (headers, account info, etc.)

    flush()
    return txns


def to_dataframe(txns: list[Txn]):
    import pandas as pd  # type: ignore

    rows = []
    for t in txns:
        # Strip the leading amount from description if present; keep full text in details.
        desc = t.description
        if desc:
            desc = AMOUNT_RE.sub("", desc, count=1).strip(" -\t")
        rows.append(
            {
                "date": t.date,
                "amount": t.amount_kzt,
                "currency": (t.ccy or "").upper() if t.ccy else None,
                "orig_amount": t.orig_amount,
                "orig_currency": (t.orig_ccy or "").upper() if t.orig_ccy else None,
                "description": desc or None,
                "details": t.details or None,
            }
        )
    df = pd.DataFrame(rows)
    if not df.empty:
        df["date"] = pd.to_datetime(df["date"], format="%d.%m.%Y", errors="coerce")
    return df


def fix_misattached_merchants(txns: list[Txn]) -> list[Txn]:
    """
    pdfplumber sometimes emits merchant-only lines out of order.
    Apply a conservative fix for known cases in this statement format.
    """

    def is_purchase(t: Txn) -> bool:
        return "покупка" in t.description.strip().lower()

    def is_debit_transfer(t: Txn) -> bool:
        return "списание" in t.description.strip().lower()

    from datetime import datetime

    def parse_dt(s: str) -> Optional[datetime]:
        try:
            return datetime.strptime(s, "%d.%m.%Y")
        except Exception:
            return None

    purchases = []
    for idx, t in enumerate(txns):
        if is_purchase(t):
            purchases.append((idx, parse_dt(t.date), t))

    for i, t in enumerate(txns):
        if not t.details:
            continue
        if "CURSOR," in t.details and is_debit_transfer(t) and not is_purchase(t):
            # Move the merchant line to the nearest purchase row (by date) lacking it.
            merchant = t.details
            t_dt = parse_dt(t.date)
            best_idx: Optional[int] = None
            best_dist = 999999
            for p_idx, p_dt, p in purchases:
                if "CURSOR," in (p.details or ""):
                    continue
                if t_dt and p_dt:
                    dist = abs((p_dt - t_dt).days)
                else:
                    dist = abs(p_idx - i)
                if dist < best_dist:
                    best_dist = dist
                    best_idx = p_idx

            # Only apply if it's very close (avoid surprising moves).
            if best_idx is not None and best_dist <= 2:
                txns[best_idx].details = _clean_line(
                    ((txns[best_idx].details or "") + " " + merchant).strip()
                )
                t.details = ""

    return txns


def main() -> int:
    ap = argparse.ArgumentParser(description="Extract transactions table from PDF into XLSX.")
    ap.add_argument("pdf", type=Path, help="Input PDF path")
    ap.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Output XLSX path (default: next to PDF with .xlsx extension)",
    )
    args = ap.parse_args()

    pdf_path: Path = args.pdf
    out_path: Path = args.output or pdf_path.with_suffix(".xlsx")

    lines = iter_pdf_lines(pdf_path)
    txns = parse_transactions(lines)
    txns = fix_misattached_merchants(txns)
    df = to_dataframe(txns)

    with __import__("pandas").ExcelWriter(out_path, engine="openpyxl") as xw:  # type: ignore
        df.to_excel(xw, index=False, sheet_name="transactions")

    print(f"Wrote {len(df)} rows to {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

