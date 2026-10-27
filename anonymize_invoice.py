"""Anonymize personal/identifying data in OpenIRIS invoice and provider-requests
.xlsx exports (the INVOICE_FILE and REQUESTS_FILE from billing_checks_and_fixes.ipynb).

Usage:
    python anonymize_invoice.py \
        --invoice-in Invoice58.xlsx --invoice-out Invoice58_anon.xlsx \
        --requests-in "Light Microscopy Unit-provider-requests.xlsx" \
        --requests-out LMU_requests_anon.xlsx

Either pair can be omitted, but running both together in one invocation is
what makes the anonymization consistent *across* the two files (e.g. the
invoice's "Requester" and the requests file's "Requester name" get the same
fake value when they refer to the same real person) - it's all driven by one
shared value->fake-value map built up during the run.

Every column not explicitly listed below, and the invoice's 2-row summary
header (if present), is copied through unchanged.
"""

import argparse
import json
import re
from pathlib import Path

import pandas as pd

MULTI_VALUE_DELIMITER = ";"
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

# (column name, role, is_multi_value) - role is one of "email", "name", "code", "text"
INVOICE_COLUMNS = [
    ("User email", "email", False),
    ("Group head email(s)", "email", True),
    ("Requester email", "email", False),
    ("User name", "name", False),
    ("Group head(s) text", "name", True),
    ("Group admin(s)", "name", True),
    ("Submitter", "name", False),
    ("Submitted by", "name", False),
    ("Requester", "name", False),
    ("Cost center code", "code", False),
    ("Group", "text", False),
    ("Cost center name", "text", False),
    ("Billing address", "text", False),
    ("Remit code", "text", False),
    ("Request Title", "text", False),
    ("Request Alt ID", "text", False),
    ("Request comments", "text", False),
    ("Booking title", "text", False),
    ("Booking comments", "text", False),
]

REQUESTS_COLUMNS = [
    ("Requester email", "email", False),
    ("Group head emails", "email", True),
    ("Submitter (email)", "email", False),
    ("Requester name", "name", False),
    ("Group heads text", "name", True),
    ("Group admins", "name", True),
    ("Submitter (name)", "name", False),
    ("Cost center code", "code", False),
    ("Request title", "text", False),
    ("Request Alt ID", "text", False),
    ("Group", "text", False),
    ("Cost center name", "text", False),
    ("Billing address", "text", False),
    ("Remit code", "text", False),
]

# not anonymized consistently - just wiped, since we don't need this info at all
REQUESTS_ERASE_COLUMNS = ["Participants", "Form"]

# "Form (JSON)" holds a JSON list of dicts; only the verifier field is needed
# downstream (see find_verifier() in billing_checks_and_fixes.ipynb) - every
# other field in it gets redacted, the verifier field is left intact.
REQUESTS_JSON_COLUMN = "Form (JSON)"
JSON_VERIFIER_KEY = "Verifier (asiatarkastaja)"


class Anonymizer:
    """Maps original values to fake values, consistently, across any number
    of columns and files processed with the same instance."""

    def __init__(self):
        self._map = {}
        self._counters = {}

    def _next_id(self, kind):
        n = self._counters.get(kind, 0) + 1
        self._counters[kind] = n
        return n

    def _fake_value(self, original, role, column):
        key = original.strip()
        if key in self._map:
            return self._map[key]

        if role == "email" or EMAIL_RE.match(key):
            fake = f"person{self._next_id('email')}@example.com"
        elif role == "name":
            fake = f"Person {self._next_id('name')}"
        elif role == "code":
            fake = str(900000 + self._next_id("code"))
        else:
            fake = f"{column} {self._next_id(column)}"

        self._map[key] = fake
        return fake

    def anonymize_cell(self, value, role, column, multi_value=False):
        if pd.isna(value):
            return value
        text = str(value)
        if text.strip() == "":
            return value

        if multi_value and MULTI_VALUE_DELIMITER in text:
            parts = [p.strip() for p in text.split(MULTI_VALUE_DELIMITER)]
            fake_parts = [self._fake_value(p, role, column) if p else p for p in parts]
            return f"{MULTI_VALUE_DELIMITER} ".join(fake_parts)

        return self._fake_value(text, role, column)


def erase_cell(value):
    if pd.isna(value):
        return value
    if str(value).strip() == "":
        return value
    return "x"


def redact_form_json(value):
    if pd.isna(value):
        return value
    text = str(value)
    if text.strip() == "":
        return value

    try:
        obj = json.loads(text.replace("\n", " "))
    except json.JSONDecodeError:
        return "x"

    def redact_item(item):
        if isinstance(item, dict):
            return {k: (v if k == JSON_VERIFIER_KEY else "x") for k, v in item.items()}
        return "x"

    if isinstance(obj, list):
        redacted = [redact_item(item) for item in obj]
    else:
        redacted = redact_item(obj)

    return json.dumps(redacted, ensure_ascii=False)


def apply_anonymization(df, columns_config, anonymizer):
    for column, role, multi_value in columns_config:
        if column not in df.columns:
            continue
        df[column] = df[column].apply(
            lambda v, role=role, column=column, multi_value=multi_value: anonymizer.anonymize_cell(
                v, role, column, multi_value
            )
        )


def anonymize_invoice(input_file, output_file, anonymizer=None):
    if anonymizer is None:
        anonymizer = Anonymizer()

    input_file = Path(input_file)
    output_file = Path(output_file)

    # invoices exported from IRIS have a 2-row summary (labels + values)
    # before the real column headers and data; detect it the same way
    # billing_checks_and_fixes.ipynb does.
    probe = pd.read_excel(input_file, nrows=1)
    has_summary = "Created by" in probe.columns

    if has_summary:
        summary = pd.read_excel(input_file, header=None, nrows=2)
        df = pd.read_excel(input_file, skiprows=[0, 1])
    else:
        summary = None
        df = pd.read_excel(input_file)

    apply_anonymization(df, INVOICE_COLUMNS, anonymizer)

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        if summary is not None:
            summary.to_excel(writer, index=False, header=False, startrow=0)
            df.to_excel(writer, index=False, startrow=2)
        else:
            df.to_excel(writer, index=False)

    return anonymizer


def anonymize_requests(input_file, output_file, anonymizer=None):
    if anonymizer is None:
        anonymizer = Anonymizer()

    input_file = Path(input_file)
    output_file = Path(output_file)

    df = pd.read_excel(input_file)

    apply_anonymization(df, REQUESTS_COLUMNS, anonymizer)

    for column in REQUESTS_ERASE_COLUMNS:
        if column in df.columns:
            df[column] = df[column].apply(erase_cell)

    if REQUESTS_JSON_COLUMN in df.columns:
        df[REQUESTS_JSON_COLUMN] = df[REQUESTS_JSON_COLUMN].apply(redact_form_json)

    df.to_excel(output_file, index=False)

    return anonymizer


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument("--invoice-in", help="original invoice .xlsx")
    parser.add_argument("--invoice-out", help="path to write the anonymized invoice to")
    parser.add_argument("--requests-in", help="original provider-requests .xlsx")
    parser.add_argument("--requests-out", help="path to write the anonymized requests file to")
    args = parser.parse_args()

    if not args.invoice_in and not args.requests_in:
        parser.error("provide --invoice-in and/or --requests-in")
    if args.invoice_in and not args.invoice_out:
        parser.error("--invoice-out is required with --invoice-in")
    if args.requests_in and not args.requests_out:
        parser.error("--requests-out is required with --requests-in")

    anonymizer = Anonymizer()

    if args.invoice_in:
        anonymize_invoice(args.invoice_in, args.invoice_out, anonymizer)
        print(f"Wrote {args.invoice_out}")

    if args.requests_in:
        anonymize_requests(args.requests_in, args.requests_out, anonymizer)
        print(f"Wrote {args.requests_out}")

    print(f"{len(anonymizer._map)} distinct values anonymized in total.")
