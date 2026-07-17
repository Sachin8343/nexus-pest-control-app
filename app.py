"""
NEXUS PEST CONTROL - PORTAL BACKEND (Python / Flask)
-----------------------------------------------------
This is a 1:1 port of the original Google Apps Script backend (Code.gs) to a
standalone Python app so the whole project (frontend + backend) can live in
one GitHub repo and be edited/deployed like normal code, instead of through
the Apps Script web editor.

Storage stays on Google Sheets (jobs/settings) and Google Drive (photos) -
same data, same free storage - just accessed here via the Google Sheets API
and Google Drive API (through a service account) instead of being natively
bound the way an Apps Script project is.

Email is sent via Gmail SMTP with an "app password" for the same
nexuspestcontrolservice@gmail.com account that used to send through MailApp.

SETUP: see README.md for the exact one-time setup steps (Google Cloud
service account, sharing the Sheet + Drive folder with it, Gmail app
password, and deploying this repo to Render).

Required environment variables (set these in your host, never commit them):
  GOOGLE_SERVICE_ACCOUNT_JSON  - full JSON key for the service account (as a single-line string)
  GOOGLE_SHEET_ID              - the spreadsheet ID of "Nexus Pest Control Portal"
  DRIVE_FOLDER_ID              - ID of the Drive folder to store job photos in
  GMAIL_ADDRESS                - nexuspestcontrolservice@gmail.com
  GMAIL_APP_PASSWORD           - 16-character Gmail App Password (not your normal password)
"""

import base64
import gc as gc_module  # aliased: this module already defines a gc() function (the Sheets client getter)
import json
import os
import re
import smtplib
import uuid
from datetime import datetime, timezone
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from zoneinfo import ZoneInfo

from flask import Flask, Response, jsonify, request, send_from_directory
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import gspread
import io

# ---------------- Config ----------------
GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "")
DRIVE_FOLDER_ID = os.environ.get("DRIVE_FOLDER_ID", "")
GMAIL_ADDRESS = os.environ.get("GMAIL_ADDRESS", "nexuspestcontrolservice@gmail.com")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

_creds = None
_gc = None
_drive = None


def _load_credentials():
    global _creds
    if _creds:
        return _creds
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "")
    if not raw:
        raise RuntimeError(
            "GOOGLE_SERVICE_ACCOUNT_JSON env var is not set. See README.md setup steps."
        )
    info = json.loads(raw)
    _creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return _creds


def gc():
    """Authenticated gspread client (Sheets)."""
    global _gc
    if _gc is None:
        _gc = gspread.authorize(_load_credentials())
    return _gc


def drive():
    """Authenticated Drive API client."""
    global _drive
    if _drive is None:
        _drive = build("drive", "v3", credentials=_load_credentials())
    return _drive


def spreadsheet():
    if not GOOGLE_SHEET_ID:
        raise RuntimeError("GOOGLE_SHEET_ID env var is not set. See README.md setup steps.")
    return gc().open_by_key(GOOGLE_SHEET_ID)


# ---------------- Schema ----------------
JOBS_SHEET = "Jobs"
PHOTOS_SHEET = "Photos"
SETTINGS_SHEET = "Settings"

JOB_HEADERS = [
    "id", "createdAt", "status", "custName", "custPhone", "custEmail", "custAddress",
    "serviceTypes", "notes", "lineItemsJSON", "completedNotes", "completedAt", "scheduledDate",
    "quoteSentAt", "reportSentAt", "paymentAmount", "paymentMethod", "paymentReceivedAt", "receiptSentAt",
    "serviceReportJSON", "reportPdfFileId",
]
PHOTO_HEADERS = ["photoId", "jobId", "fileId", "name", "uploadedAt"]

SETTINGS_DEFAULTS = {
    "companyName": "Nexus Pest Control",
    "fromEmail": "nexuspestcontrolservice@gmail.com",
    "phone": "(519) 000-0000",
    "address": "Owen Sound, ON, Canada",
    "licenseNote": "Licensed Exterminator — Ontario Pesticides Act (License # on file)",
    "taxRate": 13,
    "quoteValidityDays": 30,
    "warrantyDays": 30,
    "tripFee": 50,
}


# ---------------- Sheet helpers ----------------
def get_or_create_sheet(name, headers):
    ss = spreadsheet()
    try:
        sh = ss.worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        sh = ss.add_worksheet(title=name, rows=200, cols=max(10, len(headers)))
        sh.append_row(headers)
        sh.freeze(rows=1)
    ensure_headers(sh, headers)
    return sh


def ensure_headers(sh, headers):
    """Migration helper: pad the header row if a sheet is missing newly-added columns."""
    existing = sh.row_values(1)
    if len(existing) < len(headers):
        sh.update(range_name="A1", values=[headers])


def all_rows(sh):
    """Returns all data rows (excluding header) as plain string lists."""
    values = sh.get_all_values()
    return values[1:] if values else []


def row_to_object(headers, row):
    o = {}
    for i, h in enumerate(headers):
        o[h] = row[i] if i < len(row) else ""
    return o


def find_row_index_by_id(sh, id_col_idx, item_id):
    """Returns 1-based sheet row number, or -1 if not found."""
    values = sh.get_all_values()
    for i in range(1, len(values)):
        if len(values[i]) > id_col_idx and str(values[i][id_col_idx]) == str(item_id):
            return i + 1
    return -1


# ---------------- Settings ----------------
def get_settings():
    sh = get_or_create_sheet(SETTINGS_SHEET, ["key", "value"])
    rows = all_rows(sh)
    m = {r[0]: (r[1] if len(r) > 1 else "") for r in rows if r}
    settings = {}
    for k, default in SETTINGS_DEFAULTS.items():
        v = m.get(k, "")
        settings[k] = v if v not in (None, "") else default
    # keep numeric fields numeric
    for k in ("taxRate", "quoteValidityDays", "warrantyDays", "tripFee"):
        try:
            settings[k] = float(settings[k]) if "." in str(settings[k]) else int(settings[k])
        except (ValueError, TypeError):
            settings[k] = SETTINGS_DEFAULTS[k]
    if not m:
        save_settings(settings)  # seed sheet on first run
    return settings


def save_settings(settings):
    sh = get_or_create_sheet(SETTINGS_SHEET, ["key", "value"])
    sh.clear()
    rows = [["key", "value"]] + [[k, settings.get(k, SETTINGS_DEFAULTS[k])] for k in SETTINGS_DEFAULTS]
    sh.update(range_name="A1", values=rows)
    return get_settings()


# ---------------- Photos (Drive) ----------------
def photo_folder_id():
    if not DRIVE_FOLDER_ID:
        raise RuntimeError("DRIVE_FOLDER_ID env var is not set. See README.md setup steps.")
    return DRIVE_FOLDER_ID


def add_photo(job_id, filename, base64_data, mime_type):
    import base64

    file_bytes = base64.b64decode(base64_data)
    file_metadata = {"name": filename or "photo.jpg", "parents": [photo_folder_id()]}
    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type or "image/jpeg", resumable=False)
    file = drive().files().create(body=file_metadata, media_body=media, fields="id").execute()
    file_id = file["id"]
    drive().permissions().create(
        fileId=file_id, body={"type": "anyone", "role": "reader"}
    ).execute()

    sh = get_or_create_sheet(PHOTOS_SHEET, PHOTO_HEADERS)
    photo_id = str(uuid.uuid4())
    sh.append_row([photo_id, job_id, file_id, filename or "photo.jpg", now_iso()])
    return {
        "photoId": photo_id,
        "fileId": file_id,
        "thumb": f"https://drive.google.com/thumbnail?id={file_id}&sz=w300",
    }


def remove_photo(photo_id):
    sh = get_or_create_sheet(PHOTOS_SHEET, PHOTO_HEADERS)
    row_idx = find_row_index_by_id(sh, 0, photo_id)
    if row_idx > 0:
        file_id = sh.cell(row_idx, 3).value
        try:
            drive().files().delete(fileId=file_id).execute()
        except Exception:
            pass
        sh.delete_rows(row_idx)
    return True


def get_photos_for_job(job_id):
    sh = get_or_create_sheet(PHOTOS_SHEET, PHOTO_HEADERS)
    out = []
    for r in all_rows(sh):
        if len(r) > 1 and str(r[1]) == str(job_id):
            out.append({
                "photoId": r[0], "jobId": r[1], "fileId": r[2],
                "name": r[3] if len(r) > 3 else "", "uploadedAt": r[4] if len(r) > 4 else "",
                "thumb": f"https://drive.google.com/thumbnail?id={r[2]}&sz=w300",
            })
    return out


def get_photos_by_job():
    """Reads the Photos sheet ONCE and groups rows by jobId in memory. get_jobs()
    used to call get_photos_for_job() per job, which re-read the entire Photos
    sheet from scratch for every single job (N+1 Sheets API reads) — with more
    than a handful of jobs this tripped Google's per-minute read quota (429
    'Quota exceeded' errors). This does the same grouping with one read."""
    sh = get_or_create_sheet(PHOTOS_SHEET, PHOTO_HEADERS)
    by_job = {}
    for r in all_rows(sh):
        if len(r) > 1 and r[1]:
            by_job.setdefault(str(r[1]), []).append({
                "photoId": r[0], "jobId": r[1], "fileId": r[2],
                "name": r[3] if len(r) > 3 else "", "uploadedAt": r[4] if len(r) > 4 else "",
                "thumb": f"https://drive.google.com/thumbnail?id={r[2]}&sz=w300",
            })
    return by_job


# ---------------- Jobs ----------------
def now_iso():
    return datetime.now(timezone.utc).isoformat()


def get_jobs():
    sh = get_or_create_sheet(JOBS_SHEET, JOB_HEADERS)
    rows = all_rows(sh)
    photos_by_job = get_photos_by_job()  # one Sheets read, reused for every job below
    jobs = []
    for r in rows:
        if not r or not r[0]:
            continue
        o = row_to_object(JOB_HEADERS, r)
        o["serviceTypes"] = o["serviceTypes"].split("|") if o.get("serviceTypes") else []
        o["lineItems"] = json.loads(o["lineItemsJSON"]) if o.get("lineItemsJSON") else []
        o["customer"] = {
            "name": o.get("custName", ""), "phone": o.get("custPhone", ""),
            "email": o.get("custEmail", ""), "address": o.get("custAddress", ""),
        }
        o["payment"] = (
            {"amount": o["paymentAmount"], "method": o["paymentMethod"], "receivedAt": o["paymentReceivedAt"]}
            if o.get("paymentAmount") else None
        )
        o["serviceReport"] = json.loads(o["serviceReportJSON"]) if o.get("serviceReportJSON") else None
        o["photos"] = photos_by_job.get(str(o["id"]), [])
        # The signed PDF isn't stored anywhere (Drive rejects uploads from a
        # service account into a personal-account folder — see
        # generate_and_store_report_pdf) — it's rebuilt on request instead.
        # Show the link once there's a saved report with at least the
        # customer's signature on it.
        _r = o["serviceReport"] or {}
        o["reportPdfUrl"] = (
            f"/report-pdf/{o['id']}"
            if (_r.get("clientAcknowledgment") or {}).get("signature") else None
        )
        jobs.append(o)
    jobs.sort(key=lambda j: j.get("createdAt") or "", reverse=True)
    return jobs


def create_job(job):
    sh = get_or_create_sheet(JOBS_SHEET, JOB_HEADERS)
    job_id = str(uuid.uuid4())
    created_at = now_iso()
    customer = job.get("customer", {})
    sh.append_row([
        job_id, created_at, "Quoted",
        customer.get("name", ""), customer.get("phone", ""), customer.get("email", ""), customer.get("address", ""),
        "|".join(job.get("serviceTypes", [])), job.get("notes", ""),
        json.dumps(job.get("lineItems", [])),
        "", "", "", "", "", "", "", "", "", "", "",
    ])
    return job_id


def set_job_field(job_id, col_name, value):
    sh = get_or_create_sheet(JOBS_SHEET, JOB_HEADERS)
    row_idx = find_row_index_by_id(sh, 0, job_id)
    if row_idx < 1:
        raise ValueError("Job not found")
    col = JOB_HEADERS.index(col_name) + 1
    sh.update_cell(row_idx, col, value)


def mark_scheduled(job_id, date_str):
    set_job_field(job_id, "status", "Scheduled")
    set_job_field(job_id, "scheduledDate", date_str)
    return True


def mark_completed(job_id):
    set_job_field(job_id, "status", "Completed")
    set_job_field(job_id, "completedAt", datetime.now(timezone.utc).date().isoformat())
    return True


def save_service_report(job_id, report):
    report = report or {}
    set_job_field(job_id, "serviceReportJSON", json.dumps(report))
    set_job_field(job_id, "completedNotes", report.get("generalComments", ""))
    return True


def record_payment(job_id, amount, method):
    set_job_field(job_id, "paymentAmount", amount)
    set_job_field(job_id, "paymentMethod", method)
    set_job_field(job_id, "paymentReceivedAt", now_iso())
    set_job_field(job_id, "status", "Closed")
    try:
        send_receipt_email(job_id)
    except Exception:
        pass
    return True


def get_job_by_id(job_id):
    for j in get_jobs():
        if j["id"] == job_id:
            return j
    raise ValueError("Job not found")


# ---------------- Money / formatting ----------------
def money(n):
    try:
        n = float(n or 0)
    except (ValueError, TypeError):
        n = 0
    return "${:,.2f}".format(n)


def calc_totals(line_items, tax_rate):
    subtotal = sum((float(li.get("qty") or 0) * float(li.get("price") or 0)) for li in (line_items or []))
    tax = subtotal * (float(tax_rate or 0) / 100)
    return {"subtotal": subtotal, "tax": tax, "total": subtotal + tax}


def esc(s):
    if s is None:
        return ""
    return (
        str(s)
        .replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        .replace('"', "&quot;").replace("'", "&#39;")
    )


def fmt_date(iso):
    if not iso:
        return "—"
    try:
        d = datetime.fromisoformat(str(iso).replace("Z", "+00:00"))
    except ValueError:
        try:
            d = datetime.strptime(str(iso)[:10], "%Y-%m-%d")
        except ValueError:
            return str(iso)
    d = d.astimezone(ZoneInfo("America/Toronto")) if d.tzinfo else d
    return d.strftime("%b %-d, %Y") if os.name != "nt" else d.strftime("%b %#d, %Y")


# ---------------- Terms & Conditions ----------------
def default_terms(s):
    return [
        ("Quote Validity", f"This quote is valid for {s['quoteValidityDays']} days from the date of issue. Final pricing may be adjusted following an on-site inspection if conditions differ from those described by the customer."),
        ("Service Guarantee", f"If target pest activity persists between scheduled treatments within {s['warrantyDays']} days of service, {s['companyName']} will return to re-treat the affected area at no additional charge. This guarantee covers control of the pests listed on this quote; it does not guarantee that pests will never return to the property."),
        ("Licensing & Products", f"{s['licenseNote']}. All products used are registered with Health Canada's Pest Management Regulatory Agency (PMRA) and are applied strictly according to label directions."),
        ("Access & Preparation", "Customer is responsible for providing safe, clear access to all treatment areas and for completing any preparation instructions provided by the technician. Incomplete preparation may reduce treatment effectiveness or require an additional service charge."),
        ("Cancellation & Missed Appointments", f"Appointments may be rescheduled or cancelled free of charge with at least 24 hours' notice. Cancellations with less than 24 hours' notice, or a missed appointment due to lack of access, may be subject to a trip charge of {money(s['tripFee'])}."),
        ("Payment Terms", "Payment is due upon completion of service unless other arrangements have been agreed to in writing. We accept cash, debit, credit card, and e-transfer. Outstanding balances after 15 days may be subject to a late fee."),
        ("Health & Safety", "Keep children, pets, and food items away from treated areas until products have fully dried or as instructed by the technician. Please notify us in advance of any chemical sensitivities, allergies, pregnancy, or pets/aquariums on the property."),
        ("Liability", f"{s['companyName']} is not responsible for pre-existing damage caused by pest activity prior to treatment, or for damage resulting from a failure to follow preparation or access instructions. Our liability for direct damages is limited to the amount paid for the applicable service."),
        ("Photos & Records", "Before/after photos may be taken during service to document treatment and included in your service report for your records."),
        ("Acceptance", "Approval of this quote (by reply, signature, or payment of a deposit) constitutes acceptance of the terms above."),
    ]


def email_wrapper(body_html, s):
    return (
        '<div style="font-family:Arial,Helvetica,sans-serif;max-width:640px;margin:0 auto;color:#1b2430;">'
        '<div style="background:#0c2647;padding:18px 22px;border-radius:10px 10px 0 0;">'
        '<div style="color:#fff;font-size:20px;font-weight:800;letter-spacing:0.5px;">NEXUS <span style="color:#3fcf76;">PEST CONTROL</span></div>'
        '<div style="color:#bcd0ea;font-size:11px;letter-spacing:2px;margin-top:2px;">PROFESSIONAL PEST MANAGEMENT</div></div>'
        '<div style="border:1px solid #e2e6ec;border-top:none;border-radius:0 0 10px 10px;padding:22px;">' + body_html + '</div>'
        '<div style="font-size:11px;color:#8992a0;padding:14px 6px;">' + esc(s['companyName']) + ' • ' + esc(s['address']) + ' • ' + esc(s['phone']) + ' • ' + esc(s['fromEmail']) + '</div></div>'
    )


# ---------------- Email builders ----------------
def quote_email_html(job, s):
    totals = calc_totals(job["lineItems"], s["taxRate"])
    rows = "".join(
        '<tr><td style="padding:7px 6px;border-bottom:1px solid #eee;">' + esc(li.get("desc")) + '</td>'
        '<td style="padding:7px 6px;border-bottom:1px solid #eee;text-align:center;">' + str(li.get("qty")) + '</td>'
        '<td style="padding:7px 6px;border-bottom:1px solid #eee;text-align:right;">' + money(li.get("price")) + '</td>'
        '<td style="padding:7px 6px;border-bottom:1px solid #eee;text-align:right;">' + money(float(li.get("qty") or 0) * float(li.get("price") or 0)) + '</td></tr>'
        for li in job["lineItems"]
    )
    terms = "".join(
        f'<p style="margin:8px 0;"><strong style="color:#0c2647;">{esc(t)}.</strong> <span style="color:#5b6572;">{esc(b)}</span></p>'
        for t, b in default_terms(s)
    )
    body = (
        f'<p>Hi {esc(job["customer"]["name"])},</p>'
        f'<p>Thank you for the opportunity to quote your pest control service. Here are the details for <strong>{esc(job["customer"]["address"])}</strong>:</p>'
        f'<p><strong>Service:</strong> {esc(", ".join(job["serviceTypes"]))}</p>'
        + (f'<p><strong>Scope of Work:</strong> {esc(job.get("notes"))}</p>' if job.get("notes") else "")
        + '<table style="width:100%;border-collapse:collapse;margin:14px 0;font-size:13px;"><thead><tr style="text-align:left;color:#66707d;font-size:11px;text-transform:uppercase;">'
        '<th style="padding:6px;border-bottom:2px solid #0c2647;">Description</th><th style="padding:6px;border-bottom:2px solid #0c2647;text-align:center;">Qty</th>'
        '<th style="padding:6px;border-bottom:2px solid #0c2647;text-align:right;">Unit Price</th><th style="padding:6px;border-bottom:2px solid #0c2647;text-align:right;">Amount</th></tr></thead><tbody>' + rows + '</tbody></table>'
        '<div style="max-width:260px;margin-left:auto;font-size:13.5px;">'
        f'<div style="display:flex;justify-content:space-between;padding:3px 0;"><span>Subtotal</span><strong>{money(totals["subtotal"])}</strong></div>'
        f'<div style="display:flex;justify-content:space-between;padding:3px 0;"><span>HST ({s["taxRate"]}%)</span><strong>{money(totals["tax"])}</strong></div>'
        f'<div style="display:flex;justify-content:space-between;padding:6px 0;border-top:1px solid #ddd;font-size:16px;color:#0c2647;"><span>Total</span><strong>{money(totals["total"])}</strong></div></div>'
        f'<p style="margin-top:16px;">This quote is valid for {s["quoteValidityDays"]} days. Reply to this email or reach us at <a href="mailto:{esc(s["fromEmail"])}" style="color:#166e35;">{esc(s["fromEmail"])}</a> / {esc(s["phone"])} to schedule your service or ask questions.</p>'
        '<hr style="border:none;border-top:1px solid #e2e6ec;margin:18px 0;">'
        f'<div style="font-size:11px;color:#66707d;line-height:1.6;"><strong style="color:#0c2647;">Terms &amp; Conditions</strong><br>{terms}</div>'
    )
    return email_wrapper(body, s)


def row_(label, value):
    if value in (None, ""):
        return ""
    return (
        '<tr><td style="padding:5px 8px 5px 0;color:#66707d;white-space:nowrap;vertical-align:top;">' + esc(label) + '</td>'
        '<td style="padding:5px 0;">' + esc(value) + '</td></tr>'
    )


def checked_list(obj, labels):
    if not obj:
        return ""
    picked = [labels[k] for k in labels if obj.get(k)]
    text = ", ".join(picked)
    if obj.get("other") and obj.get("otherText"):
        text = (text + ", " if text else "") + obj["otherText"]
    return text


def section_heading(title):
    return f'<h4 style="color:#0c2647;font-size:13px;margin:18px 0 6px;border-bottom:1px solid #e2e6ec;padding-bottom:4px;">{esc(title)}</h4>'


def report_email_html(job, s, cid_names):
    photos_html = "".join(
        f'<img src="cid:{cid}" style="width:150px;height:150px;object-fit:cover;border-radius:8px;margin:4px;border:1px solid #e2e6ec;">'
        for cid in (cid_names or [])
    )
    r = job.get("serviceReport") or {}
    table_open = '<table style="width:100%;font-size:13px;border-collapse:collapse;">'
    table_close = '</table>'

    service_details = (
        table_open
        + row_("Property Type", r.get("propertyType"))
        + row_("Date of Service", fmt_date(r["dateOfService"]) if r.get("dateOfService") else "")
        + row_("Time In / Out", " – ".join(x for x in [r.get("timeIn"), r.get("timeOut")] if x))
        + row_("Technician", r.get("technicianName"))
        + row_("Technician License / Cert. No.", r.get("technicianLicense"))
        + row_("Type of Service", r.get("serviceTypeOther") if r.get("serviceType") == "Other" else r.get("serviceType"))
        + table_close
    )

    pest_section = (
        table_open
        + row_("Pest Identified", r.get("pestType"))
        + row_("Level of Infestation", r.get("infestationLevel"))
        + row_("Areas Affected", r.get("areasAffected"))
        + table_close
    )

    evidence = checked_list(r.get("evidence"), {
        "livePests": "Live pests", "deadPests": "Dead pests", "droppings": "Droppings",
        "nesting": "Nesting materials", "damage": "Damage to property", "tracks": "Tracks or smear marks",
    })
    conditions = checked_list(r.get("contributingConditions"), {
        "food": "Food sources", "water": "Water/moisture", "clutter": "Clutter",
        "cracks": "Cracks or gaps", "openings": "Open doors/windows", "sanitation": "Poor sanitation",
    })
    observation = (
        table_open + row_("Evidence of Pest Activity", evidence) + row_("Contributing Conditions", conditions) + table_close
    )

    products = [p for p in (r.get("products") or []) if p.get("name")]
    products_rows = "".join(
        '<tr>'
        f'<td style="padding:5px;border-bottom:1px solid #eee;">{esc(p.get("name"))}</td>'
        f'<td style="padding:5px;border-bottom:1px solid #eee;">{esc(p.get("activeIngredient"))}</td>'
        f'<td style="padding:5px;border-bottom:1px solid #eee;">{esc(p.get("pcpNumber"))}</td>'
        f'<td style="padding:5px;border-bottom:1px solid #eee;">{esc(p.get("amountUsed"))}</td>'
        f'<td style="padding:5px;border-bottom:1px solid #eee;">{esc(p.get("location"))}</td>'
        '</tr>'
        for p in products
    )
    products_table = (
        '<table style="width:100%;font-size:12px;border-collapse:collapse;margin-top:6px;">'
        '<thead><tr style="text-align:left;color:#66707d;font-size:10.5px;text-transform:uppercase;">'
        '<th style="padding:5px;border-bottom:2px solid #0c2647;">Product</th><th style="padding:5px;border-bottom:2px solid #0c2647;">Active Ingredient</th>'
        '<th style="padding:5px;border-bottom:2px solid #0c2647;">PCP #</th><th style="padding:5px;border-bottom:2px solid #0c2647;">Amount</th>'
        f'<th style="padding:5px;border-bottom:2px solid #0c2647;">Location</th></tr></thead><tbody>{products_rows}</tbody></table>'
    ) if products_rows else ""

    app_method = checked_list(r.get("applicationMethod"), {
        "spray": "Spray", "gelBait": "Gel Bait", "dust": "Dust", "baitStation": "Bait Station", "trap": "Trap", "fumigation": "Fumigation",
    })
    treatment = (
        table_open + row_("Treatment Method", r.get("treatmentMethod")) + table_close + products_table + table_open
        + row_("Target Pest", r.get("targetPest"))
        + row_("Application Equipment", r.get("applicationEquipment"))
        + row_("Application Method", app_method)
        + row_("Areas Treated", r.get("areasTreated"))
        + table_close
    )

    follow_up = (
        table_open
        + row_("Recommendations", r.get("recommendations"))
        + row_("Additional Treatment Required", r.get("followUpRequired"))
        + row_("Recommended Follow-Up Date", fmt_date(r["followUpDate"]) if r.get("followUpDate") else "")
        + row_("General Comments", r.get("generalComments") or job.get("completedNotes"))
        + table_close
    )

    tech_decl = r.get("technicianDeclaration") or {}
    client_ack = r.get("clientAcknowledgment") or {}
    declarations = (
        table_open
        + row_("Technician", " — ".join(x for x in [tech_decl.get("name"), fmt_date(tech_decl["date"]) if tech_decl.get("date") else ""] if x))
        + row_("Client Acknowledgment", " — ".join(x for x in [client_ack.get("name"), fmt_date(client_ack["date"]) if client_ack.get("date") else ""] if x))
        + table_close
    )

    has_report = bool(job.get("serviceReport"))

    body = (
        f'<p>Hi {esc(job["customer"]["name"])},</p>'
        f'<p>Your pest control service at <strong>{esc(job["customer"]["address"])}</strong> is complete. Full service report below:</p>'
        + section_heading("Service Details") + service_details
        + (section_heading("Pest Identified") + pest_section if has_report else "")
        + (section_heading("Observations") + observation if has_report else "")
        + (section_heading("Treatment Performed") + treatment if has_report else "")
        + (section_heading("Recommendations & Follow-Up") + follow_up if has_report else "")
        + ("" if has_report else '<p><strong>Notes:</strong><br>' + esc(job.get("completedNotes") or "—").replace("\n", "<br>") + '</p>')
        + (section_heading("Photographs") + f'<div>{photos_html}</div>' if photos_html else "")
        + (section_heading("Declaration & Acknowledgment") + declarations if has_report else "")
        + f'<p style="margin-top:16px;">Reminder: if you continue to see pest activity within {s["warrantyDays"]} days, contact us for a free re-treatment under our service guarantee.</p>'
        + f'<p>Thank you for choosing {esc(s["companyName"])}!</p>'
    )
    return email_wrapper(body, s)


# ---------------- PDF service report (with signatures) ----------------
def _fetch_job_photo_bytes(job, limit=6):
    """Fetch raw photo bytes from Drive ONCE. Both the emailed inline images
    and the PDF reuse this same list instead of each re-downloading from Drive
    — on Render's free 512MB instance, doubling up on Drive round-trips and
    in-memory copies was a real contributor to OOM crashes."""
    out = []
    for p in (job.get("photos") or [])[:limit]:
        try:
            out.append(drive().files().get_media(fileId=p["fileId"]).execute())
        except Exception:
            pass
    return out


def _downscaled_data_uri(img_bytes, max_dim=260, quality=55):
    """Re-encodes a photo at PDF-display size before embedding it. The photos
    are already client-compressed to ~700px for the dashboard/email, but
    reportlab decodes the full image just to draw it at ~130px in the PDF —
    downscaling server-side first cuts that decode memory substantially."""
    try:
        from PIL import Image

        img = Image.open(io.BytesIO(img_bytes))
        img.thumbnail((max_dim, max_dim))
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")
        out = io.BytesIO()
        img.save(out, format="JPEG", quality=quality)
        small_bytes = out.getvalue()
        del img, out
        return "data:image/jpeg;base64," + base64.b64encode(small_bytes).decode("ascii")
    except Exception:
        # Fall back to the original bytes if Pillow can't decode it for any reason.
        return "data:image/jpeg;base64," + base64.b64encode(img_bytes).decode("ascii")


def report_pdf_html(job, s, photo_data_uris, cust_sig_data_uri, tech_sig_data_uri):
    """Builds a print-ready HTML document (plain tables/divs only — xhtml2pdf's
    renderer supports a limited CSS subset, no flexbox/box-shadow) for the
    signed service report PDF."""
    r = job.get("serviceReport") or {}
    table_open = '<table style="width:100%;font-size:11px;border-collapse:collapse;">'
    table_close = "</table>"

    service_details = (
        table_open
        + row_("Property Type", r.get("propertyType"))
        + row_("Date of Service", fmt_date(r["dateOfService"]) if r.get("dateOfService") else "")
        + row_("Time In / Out", " – ".join(x for x in [r.get("timeIn"), r.get("timeOut")] if x))
        + row_("Technician", r.get("technicianName"))
        + row_("Technician License / Cert. No.", r.get("technicianLicense"))
        + row_("Type of Service", r.get("serviceTypeOther") if r.get("serviceType") == "Other" else r.get("serviceType"))
        + table_close
    )
    pest_section = (
        table_open
        + row_("Pest Identified", r.get("pestType"))
        + row_("Level of Infestation", r.get("infestationLevel"))
        + row_("Areas Affected", r.get("areasAffected"))
        + table_close
    )
    evidence = checked_list(r.get("evidence"), {
        "livePests": "Live pests", "deadPests": "Dead pests", "droppings": "Droppings",
        "nesting": "Nesting materials", "damage": "Damage to property", "tracks": "Tracks or smear marks",
    })
    conditions = checked_list(r.get("contributingConditions"), {
        "food": "Food sources", "water": "Water/moisture", "clutter": "Clutter",
        "cracks": "Cracks or gaps", "openings": "Open doors/windows", "sanitation": "Poor sanitation",
    })
    observation = table_open + row_("Evidence of Pest Activity", evidence) + row_("Contributing Conditions", conditions) + table_close

    products = [p for p in (r.get("products") or []) if p.get("name")]
    products_rows = "".join(
        "<tr>"
        f'<td style="padding:4px;border-bottom:1px solid #eee;">{esc(p.get("name"))}</td>'
        f'<td style="padding:4px;border-bottom:1px solid #eee;">{esc(p.get("activeIngredient"))}</td>'
        f'<td style="padding:4px;border-bottom:1px solid #eee;">{esc(p.get("pcpNumber"))}</td>'
        f'<td style="padding:4px;border-bottom:1px solid #eee;">{esc(p.get("amountUsed"))}</td>'
        f'<td style="padding:4px;border-bottom:1px solid #eee;">{esc(p.get("location"))}</td>'
        "</tr>"
        for p in products
    )
    products_table = (
        '<table style="width:100%;font-size:10px;border-collapse:collapse;margin-top:5px;">'
        '<tr><td style="padding:4px;font-weight:bold;border-bottom:1px solid #0c2647;">Product</td>'
        '<td style="padding:4px;font-weight:bold;border-bottom:1px solid #0c2647;">Active Ingredient</td>'
        '<td style="padding:4px;font-weight:bold;border-bottom:1px solid #0c2647;">PCP #</td>'
        '<td style="padding:4px;font-weight:bold;border-bottom:1px solid #0c2647;">Amount</td>'
        f'<td style="padding:4px;font-weight:bold;border-bottom:1px solid #0c2647;">Location</td></tr>{products_rows}</table>'
    ) if products_rows else ""

    app_method = checked_list(r.get("applicationMethod"), {
        "spray": "Spray", "gelBait": "Gel Bait", "dust": "Dust", "baitStation": "Bait Station", "trap": "Trap", "fumigation": "Fumigation",
    })
    treatment = (
        table_open + row_("Treatment Method", r.get("treatmentMethod")) + table_close + products_table + table_open
        + row_("Target Pest", r.get("targetPest"))
        + row_("Application Equipment", r.get("applicationEquipment"))
        + row_("Application Method", app_method)
        + row_("Areas Treated", r.get("areasTreated"))
        + table_close
    )
    follow_up = (
        table_open
        + row_("Recommendations", r.get("recommendations"))
        + row_("Additional Treatment Required", r.get("followUpRequired"))
        + row_("Recommended Follow-Up Date", fmt_date(r["followUpDate"]) if r.get("followUpDate") else "")
        + row_("General Comments", r.get("generalComments") or job.get("completedNotes"))
        + table_close
    )

    has_report = bool(job.get("serviceReport"))
    tech_decl = r.get("technicianDeclaration") or {}
    client_ack = r.get("clientAcknowledgment") or {}

    photos_html = "".join(
        f'<img src="{uri}" style="width:130px;height:130px;margin:4px;border:1px solid #ccc;">'
        for uri in (photo_data_uris or [])
    )

    def sig_cell(title, sig_uri, name, date):
        img = (
            f'<img src="{sig_uri}" style="height:55px;max-width:220px;border-bottom:1px solid #333;">'
            if sig_uri else '<div style="height:55px;border-bottom:1px solid #333;"></div>'
        )
        return (
            '<td style="width:260px;padding:6px;vertical-align:top;">'
            f'<div style="font-size:10px;color:#66707d;margin-bottom:4px;">{esc(title)}</div>'
            + img
            + f'<div style="font-size:10px;margin-top:4px;">{esc(name or "")}{" — " + fmt_date(date) if date else ""}</div>'
            "</td>"
        )

    sig_block = (
        '<table style="width:100%;margin-top:10px;"><tr>'
        + sig_cell("Client Acknowledgment (signed first)", cust_sig_data_uri, client_ack.get("name") or job["customer"]["name"], client_ack.get("date"))
        + sig_cell("Technician Declaration (signed second)", tech_sig_data_uri, tech_decl.get("name") or r.get("technicianName"), tech_decl.get("date"))
        + "</tr></table>"
    )

    body = (
        '<div style="background:#0c2647;padding:14px 18px;">'
        '<div style="color:#ffffff;font-size:18px;font-weight:bold;">NEXUS PEST CONTROL</div>'
        '<div style="color:#bcd0ea;font-size:9px;letter-spacing:2px;">PROFESSIONAL PEST MANAGEMENT &mdash; SERVICE REPORT</div></div>'
        '<div style="padding:16px;">'
        f'<p style="font-size:12px;">Customer: <strong>{esc(job["customer"]["name"])}</strong><br>'
        f'Address: {esc(job["customer"]["address"])}<br>'
        f'Phone: {esc(job["customer"]["phone"])} &nbsp; Email: {esc(job["customer"]["email"])}</p>'
        + section_heading("Service Details") + service_details
        + (section_heading("Pest Identified") + pest_section if has_report else "")
        + (section_heading("Observations") + observation if has_report else "")
        + (section_heading("Treatment Performed") + treatment if has_report else "")
        + (section_heading("Recommendations &amp; Follow-Up") + follow_up if has_report else "")
        + (section_heading("Photographs") + f"<div>{photos_html}</div>" if photos_html else "")
        + section_heading("Signatures") + sig_block
        + f'<p style="font-size:9px;color:#8992a0;margin-top:16px;">{esc(s["companyName"])} &bull; {esc(s["address"])} &bull; {esc(s["phone"])} &bull; {esc(s["fromEmail"])}<br>'
        f'Reminder: if pest activity continues within {s["warrantyDays"]} days of service, contact us for a free re-treatment under our service guarantee.</p>'
        "</div>"
    )
    return f'<html><head><meta charset="utf-8"></head><body style="font-family:Helvetica,Arial,sans-serif;color:#1b2430;">{body}</body></html>'


def html_to_pdf_bytes(html):
    from xhtml2pdf import pisa

    buf = io.BytesIO()
    result = pisa.CreatePDF(io.StringIO(html), dest=buf)
    if result.err:
        raise RuntimeError("PDF generation failed")
    return buf.getvalue()


def generate_and_store_report_pdf(job_id, photo_bytes=None, job=None, settings=None):
    """Builds the signed PDF and returns the raw bytes.

    NOTE: this used to also upload the PDF to Drive so there'd be a permanent
    link, but Drive rejected every upload with a 403
    "Service Accounts do not have storage quota" error. That's not a
    permissions/sharing problem — it's a hard Google Drive limitation for any
    service account writing into a *personal* (non-Workspace) Google account's
    storage: service accounts get 0 storage quota there, and the Shared
    Drives / domain-wide-delegation workarounds both require a paid Google
    Workspace account, which this project isn't on. So instead of storing the
    PDF anywhere, we just regenerate it on demand — the report is fully
    reproducible from the job's saved serviceReport + photos, so a stored copy
    isn't actually necessary. See the "/report-pdf/<job_id>" route below.

    job/settings let a caller that already has them (complete_job_with_report)
    pass them straight through — every Sheets read counts against Google's
    per-minute read quota, and re-fetching the same job + settings 2-3 times
    within one "Finish Report & Sign Off" click was enough to trip it."""
    job = job if job is not None else get_job_by_id(job_id)
    s = settings if settings is not None else get_settings()
    r = job.get("serviceReport") or {}
    cust_sig = (r.get("clientAcknowledgment") or {}).get("signature") or ""
    tech_sig = (r.get("technicianDeclaration") or {}).get("signature") or ""

    if photo_bytes is None:
        photo_bytes = _fetch_job_photo_bytes(job)
    # The PDF only needs a handful of small preview images — keep this
    # tighter than the email's inline photo count to limit decode memory.
    photo_uris = [_downscaled_data_uri(b) for b in photo_bytes[:4]]

    html = report_pdf_html(job, s, photo_uris, cust_sig, tech_sig)
    del photo_uris
    pdf_bytes = html_to_pdf_bytes(html)
    del html
    gc_module.collect()

    return {"url": f"/report-pdf/{job_id}", "bytes": pdf_bytes}


def complete_job_with_report(job_id, report, send_to_customer):
    """Saves the signed service report, marks the job completed, generates the
    signed PDF, and (if the technician chose to) emails it to the customer.
    Fetches job photos from Drive exactly once and reuses them for both the
    PDF and the email — this endpoint does several memory-heavy things
    back-to-back (PDF render + email attachment) so it frees each big buffer
    as soon as it's no longer needed rather than letting them all pile up."""
    save_service_report(job_id, report)
    mark_completed(job_id)
    job = get_job_by_id(job_id)
    s = get_settings()
    photo_bytes = _fetch_job_photo_bytes(job)
    pdf_info = generate_and_store_report_pdf(job_id, photo_bytes=photo_bytes, job=job, settings=s)
    sent = False
    if send_to_customer:
        send_report_email(job_id, photo_bytes=photo_bytes, pdf_bytes=pdf_info["bytes"], job=job, settings=s)
        sent = True
    result = {"pdfUrl": pdf_info["url"], "sent": sent}
    del photo_bytes, pdf_info
    gc_module.collect()
    return result


def receipt_email_html(job, s):
    body = (
        f'<p>Hi {esc(job["customer"]["name"])},</p>'
        f'<p>This confirms payment has been received for your service at <strong>{esc(job["customer"]["address"])}</strong>. Your case is now closed.</p>'
        '<table style="width:100%;font-size:13.5px;margin:14px 0;">'
        f'<tr><td style="padding:5px 0;color:#66707d;">Service</td><td style="padding:5px 0;text-align:right;">{esc(", ".join(job["serviceTypes"]))}</td></tr>'
        f'<tr><td style="padding:5px 0;color:#66707d;">Amount Paid</td><td style="padding:5px 0;text-align:right;font-weight:700;">{money(job["payment"]["amount"])}</td></tr>'
        f'<tr><td style="padding:5px 0;color:#66707d;">Payment Method</td><td style="padding:5px 0;text-align:right;">{esc(job["payment"]["method"])}</td></tr>'
        f'<tr><td style="padding:5px 0;color:#66707d;">Date</td><td style="padding:5px 0;text-align:right;">{fmt_date(job["payment"]["receivedAt"])}</td></tr></table>'
        f'<p>Thank you for your business! If pest activity returns within {s["warrantyDays"]} days, you\'re covered under our service guarantee — just reply to this email.</p>'
    )
    return email_wrapper(body, s)


# ---------------- Email sending (Gmail SMTP) ----------------
def send_email(to, subject, html_body, text_body, cc=None, inline_images=None, attachments=None):
    if not GMAIL_APP_PASSWORD:
        raise RuntimeError("GMAIL_APP_PASSWORD env var is not set. See README.md setup steps.")
    s = get_settings()
    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"] = formataddr((s["companyName"], GMAIL_ADDRESS))
    msg["To"] = to
    if cc:
        msg["Cc"] = cc
    msg["Reply-To"] = s["fromEmail"]

    body_part = MIMEMultipart("related")
    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(text_body or " ", "plain"))
    alt.attach(MIMEText(html_body, "html"))
    body_part.attach(alt)

    for cid, img_bytes in (inline_images or {}).items():
        img = MIMEImage(img_bytes)
        img.add_header("Content-ID", f"<{cid}>")
        img.add_header("Content-Disposition", "inline", filename=f"{cid}.jpg")
        body_part.attach(img)

    msg.attach(body_part)

    for fname, file_bytes, mimetype in (attachments or []):
        subtype = (mimetype.split("/")[-1] if mimetype else "pdf")
        part = MIMEApplication(file_bytes, _subtype=subtype)
        part.add_header("Content-Disposition", "attachment", filename=fname)
        msg.attach(part)

    recipients = [to] + ([cc] if cc else [])
    # Explicit timeout matters here: smtplib.SMTP() with no timeout can hang
    # indefinitely on a slow/unreliable outbound connection (observed on
    # Render's free tier), and gunicorn's default 30s worker timeout would
    # then SIGKILL the whole worker (logged misleadingly as "out of memory?")
    # instead of this raising a normal, catchable exception.
    with smtplib.SMTP("smtp.gmail.com", 587, timeout=20) as server:
        server.starttls()
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_ADDRESS, recipients, msg.as_string())


# ---------------- Send actions ----------------
def send_quote_email(job_id):
    job = get_job_by_id(job_id)
    s = get_settings()
    if not job["customer"]["email"]:
        raise ValueError("No customer email on file.")
    html = quote_email_html(job, s)
    totals = calc_totals(job["lineItems"], s["taxRate"])
    send_email(
        to=job["customer"]["email"], cc=s["fromEmail"],
        subject=f'Your Quote from {s["companyName"]} — {", ".join(job["serviceTypes"])}',
        html_body=html, text_body=f'View this quote in HTML. Total: {money(totals["total"])}',
    )
    set_job_field(job_id, "quoteSentAt", now_iso())
    return True


def send_report_email(job_id, photo_bytes=None, pdf_bytes=None, job=None, settings=None):
    """photo_bytes / pdf_bytes / job / settings let complete_job_with_report
    pass along data it already fetched this request, instead of this function
    re-downloading the same photos/PDF/job/settings from Drive and Sheets
    again. That duplication was costing enough Sheets API reads per click to
    trip Google's per-minute read quota (429 errors). Manual re-sends (no
    caller-supplied values) still work — they just fetch fresh."""
    job = job if job is not None else get_job_by_id(job_id)
    s = settings if settings is not None else get_settings()
    if not job["customer"]["email"]:
        raise ValueError("No customer email on file.")

    if photo_bytes is None:
        photo_bytes = _fetch_job_photo_bytes(job)
    inline_images = {}
    cid_names = []
    for i, data in enumerate(photo_bytes[:6]):
        cid = f"photo{i}"
        inline_images[cid] = data
        cid_names.append(cid)
    html = report_email_html(job, s, cid_names)

    # Make sure a signed PDF exists and attach it — this is the record the
    # customer keeps (with both signatures embedded). The PDF isn't stored
    # anywhere (see generate_and_store_report_pdf's docstring), so a resend
    # just regenerates it fresh from the saved serviceReport data.
    attachments = []
    if pdf_bytes is None:
        try:
            pdf_info = generate_and_store_report_pdf(job_id, photo_bytes=photo_bytes, job=job, settings=s)
            pdf_bytes = pdf_info["bytes"]
        except Exception:
            pdf_bytes = None
    if pdf_bytes:
        attachments.append(("Service Report.pdf", pdf_bytes, "application/pdf"))

    send_email(
        to=job["customer"]["email"], cc=s["fromEmail"],
        subject=f'Service Report — {s["companyName"]} — {job["customer"]["address"]}',
        html_body=html, text_body=f'Service complete at {job["customer"]["address"]}. Notes: {job.get("completedNotes", "")}',
        inline_images=inline_images, attachments=attachments,
    )
    set_job_field(job_id, "reportSentAt", now_iso())
    del inline_images, attachments, html
    gc_module.collect()
    return True


def send_receipt_email(job_id):
    job = get_job_by_id(job_id)
    s = get_settings()
    if not job["customer"]["email"] or not job.get("payment"):
        return False
    html = receipt_email_html(job, s)
    send_email(
        to=job["customer"]["email"], cc=s["fromEmail"],
        subject=f'Payment Receipt — {s["companyName"]}',
        html_body=html, text_body=f'Receipt for {money(job["payment"]["amount"])}',
    )
    set_job_field(job_id, "receiptSentAt", now_iso())
    return True


# ---------------- Bootstrap ----------------
def get_bootstrap():
    return {"settings": get_settings(), "jobs": get_jobs()}


# ---------------- API dispatch table ----------------
API_FUNCTIONS = {
    "getBootstrap": lambda: get_bootstrap(),
    "getSettings": lambda: get_settings(),
    "saveSettings": lambda settings: save_settings(settings),
    "getJobs": lambda: get_jobs(),
    "createJob": lambda job: create_job(job),
    "markScheduled": lambda job_id, date_str: mark_scheduled(job_id, date_str),
    "markCompleted": lambda job_id: mark_completed(job_id),
    "saveServiceReport": lambda job_id, report: save_service_report(job_id, report),
    "completeJobWithReport": lambda job_id, report, send_to_customer: complete_job_with_report(job_id, report, send_to_customer),
    "recordPayment": lambda job_id, amount, method: record_payment(job_id, amount, method),
    "sendQuoteEmail": lambda job_id: send_quote_email(job_id),
    "sendReportEmail": lambda job_id: send_report_email(job_id),
    "sendReceiptEmail": lambda job_id: send_receipt_email(job_id),
    "addPhoto": lambda job_id, filename, b64, mime: add_photo(job_id, filename, b64, mime),
    "removePhoto": lambda photo_id: remove_photo(photo_id),
}


def handle_api(action, payload_args):
    result, error = None, None
    try:
        fn = API_FUNCTIONS.get(action)
        if not fn:
            raise ValueError(f"Unknown action: {action}")
        result = fn(*(payload_args or []))
    except Exception as err:
        error = str(err)
    return {"result": result, "error": error}


# ---------------- Flask app ----------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__, static_folder=None)


@app.route("/")
def index():
    return send_from_directory(BASE_DIR, "index.html")


@app.route("/api", methods=["GET", "POST"])
def api():
    if request.method == "GET":
        action = request.args.get("action")
        payload_json = request.args.get("payload")
        args = json.loads(payload_json) if payload_json else []
    else:
        body = request.get_json(silent=True) or {}
        action = body.get("action")
        args = body.get("payload") or []
    return jsonify(handle_api(action, args))


@app.route("/report-pdf/<job_id>")
def report_pdf(job_id):
    """Regenerates the signed service-report PDF fresh and streams it back.
    There's no stored copy to fetch (see generate_and_store_report_pdf) —
    everything needed (service report answers, both signatures, photos) is
    already saved on the job, so it's cheap to rebuild on every request."""
    try:
        job = get_job_by_id(job_id)
        if not job:
            return ("Job not found.", 404)
        pdf_info = generate_and_store_report_pdf(job_id, job=job, settings=get_settings())
    except Exception as err:
        return (f"Could not generate the report PDF: {err}", 500)
    safe_name = re.sub(r"[^A-Za-z0-9 _-]", "", job["customer"].get("name") or "Customer")
    filename = f"Service Report - {safe_name}.pdf"
    return Response(
        pdf_info["bytes"],
        mimetype="application/pdf",
        headers={"Content-Disposition": f'inline; filename="{filename}"'},
    )


@app.route("/healthz")
def healthz():
    return jsonify({"ok": True})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
