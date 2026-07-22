import os
import io
import re
import time
import uuid
import json
import base64
import threading
from datetime import datetime, timezone

import requests
from flask import Flask, jsonify, request
from flask_cors import CORS
from apscheduler.schedulers.background import BackgroundScheduler
import sendgrid
from sendgrid.helpers.mail import (
    Mail, Attachment, FileContent, FileName, FileType, Disposition,
)
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

app = Flask(__name__)
CORS(app)

# ─── eBay credentials ─────────────────────────────────────────────────────────
EBAY_CLIENT_ID     = os.environ.get("EBAY_CLIENT_ID", "")
EBAY_CLIENT_SECRET = os.environ.get("EBAY_CLIENT_SECRET", "")
EBAY_ENV           = os.environ.get("EBAY_ENV", "production")

# ─── Email / SendGrid ─────────────────────────────────────────────────────────
SENDGRID_API_KEY = os.environ.get("SENDGRID_API_KEY", "")
ALERT_EMAIL_TO   = os.environ.get("ALERT_EMAIL_TO",   "marcus@apd.co.uk")
ALERT_EMAIL_FROM = os.environ.get("ALERT_EMAIL_FROM", "alerts@apd.co.uk")

# ─── Watchlist persistence ────────────────────────────────────────────────────
WATCHLIST_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "watchlist.json")

# ─── In-memory state ──────────────────────────────────────────────────────────
_token_cache    = {"token": None, "expires_at": 0}
_watchlist_lock = threading.Lock()
_scan_status    = {
    "last_ran":      None,
    "items_checked": 0,
    "changes_found": 0,
    "email_sent":    False,
    "scanning":      False,
}


# ═══════════════════════════════════════════════════════════════════════════════
# eBay auth
# ═══════════════════════════════════════════════════════════════════════════════

def get_token():
    now = time.time()
    if _token_cache["token"] and now < _token_cache["expires_at"] - 60:
        return _token_cache["token"]

    url = (
        "https://api.sandbox.ebay.com/identity/v1/oauth2/token"
        if EBAY_ENV == "sandbox"
        else "https://api.ebay.com/identity/v1/oauth2/token"
    )
    credentials = base64.b64encode(
        f"{EBAY_CLIENT_ID}:{EBAY_CLIENT_SECRET}".encode()
    ).decode()
    resp = requests.post(
        url,
        headers={
            "Authorization": f"Basic {credentials}",
            "Content-Type":  "application/x-www-form-urlencoded",
        },
        data={
            "grant_type": "client_credentials",
            "scope":      "https://api.ebay.com/oauth/api_scope",
        },
        timeout=10,
    )
    resp.raise_for_status()
    data = resp.json()
    _token_cache["token"]      = data["access_token"]
    _token_cache["expires_at"] = now + data.get("expires_in", 7200)
    return _token_cache["token"]


# ═══════════════════════════════════════════════════════════════════════════════
# Watchlist helpers
# ═══════════════════════════════════════════════════════════════════════════════

def load_watchlist():
    if not os.path.exists(WATCHLIST_FILE):
        return []
    with open(WATCHLIST_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def save_watchlist(items):
    with open(WATCHLIST_FILE, "w", encoding="utf-8") as f:
        json.dump(items, f, indent=2)


# ═══════════════════════════════════════════════════════════════════════════════
# Pricing logic
# ═══════════════════════════════════════════════════════════════════════════════

def extract_part_number(search_term):
    """Extract a part number from a search term string."""
    m = re.search(r'\b\d{4,}\b', search_term)
    if m:
        return m.group(0)
    m = re.search(r'\b[A-Za-z]{1,5}[-]?\d{3,}\b', search_term)
    if m:
        return m.group(0).upper()
    return ""


def get_confidence(listing_count):
    if listing_count >= 10:
        return "High"
    if listing_count >= 5:
        return "Medium"
    return "Low"


def calculate_action(my_price, cost_price, min_margin_percent, competitor_lowest, listing_count):
    """Return (action, suggested_price, reason, confidence)."""
    min_floor  = cost_price * (1 + min_margin_percent / 100)
    confidence = get_confidence(listing_count)

    if competitor_lowest < my_price * 0.85:
        pct      = (my_price - competitor_lowest) / my_price * 100
        suggested = max(round(competitor_lowest * 1.02, 2), round(min_floor, 2))
        return (
            "Lower",
            suggested,
            f"Competitor is {pct:.1f}% cheaper than your price",
            confidence,
        )

    if my_price < competitor_lowest * 0.85:
        pct      = (competitor_lowest - my_price) / competitor_lowest * 100
        suggested = max(round(competitor_lowest * 0.95, 2), round(min_floor, 2))
        return (
            "Raise",
            suggested,
            f"You are {pct:.1f}% cheaper than nearest competitor",
            confidence,
        )

    return "Hold", round(my_price, 2), "Price within 15% of nearest competitor", confidence


def _ebay_base_url():
    return "https://api.sandbox.ebay.com" if EBAY_ENV == "sandbox" else "https://api.ebay.com"


def scan_watchlist_item(item):
    """Search eBay for item, return pricing dict. Raises on failure."""
    token = get_token()
    resp  = requests.get(
        f"{_ebay_base_url()}/buy/browse/v1/item_summary/search",
        headers={
            "Authorization":           f"Bearer {token}",
            "X-EBAY-C-MARKETPLACE-ID": item.get("market", "EBAY_GB"),
            "Content-Type":            "application/json",
        },
        params={"q": item["search_term"], "limit": 10},
        timeout=15,
    )
    resp.raise_for_status()

    prices = []
    for s in resp.json().get("itemSummaries", []):
        try:
            p = float(s.get("price", {}).get("value", 0))
            if p > 0:
                prices.append(p)
        except (ValueError, TypeError):
            pass

    if not prices:
        raise ValueError("No valid prices in eBay results")

    competitor_lowest = round(min(prices), 2)
    market_average    = round(sum(prices) / len(prices), 2)
    listing_count     = len(prices)

    action, suggested_price, reason, confidence = calculate_action(
        item["my_price"],
        item.get("cost_price", 0),
        item.get("min_margin_percent", 20),
        competitor_lowest,
        listing_count,
    )

    return {
        "item":              item,
        "competitor_lowest": competitor_lowest,
        "market_average":    market_average,
        "listing_count":     listing_count,
        "action":            action,
        "suggested_price":   suggested_price,
        "reason":            reason,
        "confidence":        confidence,
        "part_number":       extract_part_number(item["search_term"]),
    }


# ═══════════════════════════════════════════════════════════════════════════════
# Scan runner (called by routes and scheduler)
# ═══════════════════════════════════════════════════════════════════════════════

def run_scan():
    global _scan_status
    if _scan_status["scanning"]:
        return {"error": "Scan already in progress"}, 409

    _scan_status["scanning"] = True
    lower_items   = []
    raise_items   = []
    all_results   = []
    errors        = []
    price_changes = 0

    try:
        with _watchlist_lock:
            items = load_watchlist()

        active = [it for it in items if it.get("active", True)]

        for item in active:
            try:
                result = scan_watchlist_item(item)
            except Exception as e:
                errors.append({
                    "id":        item["id"],
                    "part_name": item["part_name"],
                    "error":     str(e),
                })
                continue

            now_str    = datetime.now(timezone.utc).isoformat()
            new_price  = result["competitor_lowest"]
            last_price = item.get("last_price")
            threshold  = item.get("alert_threshold_percent", 2) / 100

            price_changed = (
                last_price is not None
                and abs(new_price - last_price) / max(last_price, 0.01) > threshold
            )
            if price_changed:
                price_changes += 1

            item["last_price"]   = new_price
            item["last_checked"] = now_str
            item.setdefault("price_history", []).append(
                {"date": now_str[:10], "price": new_price}
            )
            item["price_history"] = item["price_history"][-90:]

            result["price_changed"] = price_changed
            all_results.append(result)

            if result["action"] == "Lower":
                lower_items.append(result)
            elif result["action"] == "Raise":
                raise_items.append(result)

        with _watchlist_lock:
            save_watchlist(items)

        summary = {
            "total_active":   len(active),
            "total_scanned":  len(all_results),
            "errors":         len(errors),
            "lower_needed":   len(lower_items),
            "raise_possible": len(raise_items),
            "hold":           len(all_results) - len(lower_items) - len(raise_items),
            "price_changes":  price_changes,
        }

        email_sent, email_msg = send_alert_email(lower_items, raise_items, summary, all_results)

        _scan_status.update({
            "last_ran":      datetime.now(timezone.utc).isoformat(),
            "items_checked": len(all_results),
            "changes_found": price_changes,
            "email_sent":    email_sent,
            "scanning":      False,
        })

        return {
            "summary":    summary,
            "results": [
                {
                    "id":               r["item"]["id"],
                    "part_name":        r["item"]["part_name"],
                    "action":           r["action"],
                    "competitor_lowest": r["competitor_lowest"],
                    "suggested_price":  r["suggested_price"],
                    "confidence":       r["confidence"],
                    "price_changed":    r["price_changed"],
                }
                for r in all_results
            ],
            "errors":     errors,
            "email_sent": email_sent,
            "email_msg":  email_msg,
        }, 200

    except Exception as e:
        _scan_status["scanning"] = False
        return {"error": str(e)}, 500


# ═══════════════════════════════════════════════════════════════════════════════
# Excel generation (MAM format)
# ═══════════════════════════════════════════════════════════════════════════════

def generate_excel(scan_results):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MAM Price Update"

    headers = [
        "Part Number",
        "Description",
        "Current Price (£)",
        "Suggested Price (£)",
        "Competitor Lowest (£)",
        "Market Average (£)",
        "Confidence",
        "Action",
        "Reason",
    ]

    hdr_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    hdr_font = Font(color="FFFFFF", bold=True)

    for col, h in enumerate(headers, 1):
        cell            = ws.cell(row=1, column=col, value=h)
        cell.fill       = hdr_fill
        cell.font       = hdr_font
        cell.alignment  = Alignment(horizontal="center")

    action_fills = {
        "Lower": PatternFill(start_color="FFD7D7", end_color="FFD7D7", fill_type="solid"),
        "Raise": PatternFill(start_color="D7FFD7", end_color="D7FFD7", fill_type="solid"),
        "Hold":  PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"),
    }

    for row_idx, result in enumerate(scan_results, 2):
        item   = result["item"]
        action = result.get("action", "Hold")
        fill   = action_fills.get(action, action_fills["Hold"])

        row_data = [
            result.get("part_number") or extract_part_number(item["search_term"]),
            item["part_name"],
            item["my_price"],
            result.get("suggested_price", item["my_price"]),
            result.get("competitor_lowest"),
            result.get("market_average"),
            result.get("confidence", "Low"),
            action,
            result.get("reason", ""),
        ]

        for col_idx, val in enumerate(row_data, 1):
            cell      = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.fill = fill
            if col_idx in (3, 4, 5, 6) and isinstance(val, (int, float)):
                cell.number_format = "£#,##0.00"

    col_widths = [16, 36, 18, 20, 22, 18, 12, 8, 55]
    for i, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
# Email
# ═══════════════════════════════════════════════════════════════════════════════

def _table_rows(results):
    rows = ""
    for r in results:
        item = r["item"]
        rows += (
            f"<tr>"
            f"<td style='padding:6px 10px;border-bottom:1px solid #eee'>{item['part_name']}</td>"
            f"<td style='padding:6px 10px;border-bottom:1px solid #eee'>£{item['my_price']:.2f}</td>"
            f"<td style='padding:6px 10px;border-bottom:1px solid #eee'>£{r.get('competitor_lowest',0):.2f}</td>"
            f"<td style='padding:6px 10px;border-bottom:1px solid #eee;font-weight:bold'>"
            f"£{r.get('suggested_price', item['my_price']):.2f}</td>"
            f"<td style='padding:6px 10px;border-bottom:1px solid #eee'>{r.get('confidence','')}</td>"
            f"<td style='padding:6px 10px;border-bottom:1px solid #eee'>{r.get('reason','')}</td>"
            f"</tr>"
        )
    return rows


def _section(title, color, emoji, results):
    if not results:
        return ""
    rows = _table_rows(results)
    return f"""
    <div style="margin-bottom:32px">
      <h2 style="color:{color};margin-bottom:8px">{emoji} {title}</h2>
      <table style="width:100%;border-collapse:collapse;font-size:13px">
        <thead>
          <tr style="background:{color};color:#fff">
            <th style="padding:8px 10px;text-align:left">Part</th>
            <th style="padding:8px 10px;text-align:left">Your Price</th>
            <th style="padding:8px 10px;text-align:left">Competitor</th>
            <th style="padding:8px 10px;text-align:left">Suggested</th>
            <th style="padding:8px 10px;text-align:left">Confidence</th>
            <th style="padding:8px 10px;text-align:left">Reason</th>
          </tr>
        </thead>
        <tbody>{rows}</tbody>
      </table>
    </div>"""


def build_email_html(lower_items, raise_items, summary):
    today = datetime.now(timezone.utc).strftime("%d %B %Y")
    return f"""<!DOCTYPE html>
<html>
<body style="font-family:Arial,sans-serif;max-width:920px;margin:0 auto;padding:20px;color:#222">
  <h1 style="border-bottom:3px solid #1a5276;padding-bottom:10px;margin-bottom:24px">
    eBay Parts Finder — Daily Price Report
    <span style="display:block;font-size:14px;font-weight:normal;color:#555;margin-top:4px">{today}</span>
  </h1>

  {_section("Urgent — Reprice Lower", "#c0392b", "🔴", lower_items)}
  {_section("Opportunities — Raise Prices", "#d68910", "🟡", raise_items)}

  <div style="background:#eafaf1;border-radius:8px;padding:20px;margin-top:24px">
    <h2 style="color:#1e8449;margin-top:0">🟢 Daily Summary</h2>
    <table style="font-size:14px;border-collapse:collapse">
      <tr><td style="padding:4px 20px 4px 0;color:#555">Active items monitored</td>
          <td><strong>{summary.get('total_active', 0)}</strong></td></tr>
      <tr><td style="padding:4px 20px 4px 0;color:#555">Successfully scanned</td>
          <td><strong>{summary.get('total_scanned', 0)}</strong></td></tr>
      <tr><td style="padding:4px 20px 4px 0;color:#555">Scan errors</td>
          <td><strong>{summary.get('errors', 0)}</strong></td></tr>
      <tr><td style="padding:4px 20px 4px 0;color:#555">Price changes detected</td>
          <td><strong>{summary.get('price_changes', 0)}</strong></td></tr>
      <tr><td style="padding:4px 20px 4px 0;color:#555">Urgent reprice needed</td>
          <td><strong style="color:#c0392b">{summary.get('lower_needed', 0)}</strong></td></tr>
      <tr><td style="padding:4px 20px 4px 0;color:#555">Raise opportunities</td>
          <td><strong style="color:#d68910">{summary.get('raise_possible', 0)}</strong></td></tr>
      <tr><td style="padding:4px 20px 4px 0;color:#555">Holding (no action needed)</td>
          <td><strong style="color:#1e8449">{summary.get('hold', 0)}</strong></td></tr>
    </table>
  </div>

  <p style="margin-top:24px;font-size:12px;color:#999">
    MAM-compatible price update spreadsheet attached. Generated by eBay Parts Finder.
  </p>
</body>
</html>"""


def send_alert_email(lower_items, raise_items, summary, all_results):
    if not lower_items and not raise_items:
        return False, "No actionable changes — email not sent"
    if not SENDGRID_API_KEY:
        return False, "SENDGRID_API_KEY not configured"

    today         = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    urgent_count  = len(lower_items)
    subject       = (
        f"⚠️ {urgent_count} urgent reprice(s) — eBay Parts {today}"
        if urgent_count
        else f"eBay Parts — price raise opportunities {today}"
    )

    html          = build_email_html(lower_items, raise_items, summary)
    excel_bytes   = generate_excel(all_results)
    encoded_excel = base64.b64encode(excel_bytes).decode()

    message = Mail(
        from_email=ALERT_EMAIL_FROM,
        to_emails=ALERT_EMAIL_TO,
        subject=subject,
        html_content=html,
    )
    message.attachment = Attachment(
        FileContent(encoded_excel),
        FileName(f"mam_price_update_{today}.xlsx"),
        FileType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        Disposition("attachment"),
    )

    try:
        sg       = sendgrid.SendGridAPIClient(api_key=SENDGRID_API_KEY)
        response = sg.send(message)
        return True, f"Sent (HTTP {response.status_code})"
    except Exception as e:
        return False, str(e)


# ═══════════════════════════════════════════════════════════════════════════════
# Scheduler — daily 07:00 UTC
# ═══════════════════════════════════════════════════════════════════════════════

def _scheduled_scan():
    try:
        run_scan()
    except Exception:
        pass


_scheduler = BackgroundScheduler(daemon=True)
_scheduler.add_job(_scheduled_scan, "cron", hour=7, minute=0)
_scheduler.start()


# ═══════════════════════════════════════════════════════════════════════════════
# Routes — health + eBay search (existing)
# ═══════════════════════════════════════════════════════════════════════════════

@app.route("/health")
def health():
    return jsonify({"status": "ok"})


@app.route("/search")
def search():
    q = request.args.get("q", "").strip()
    if not q:
        return jsonify({"error": "Missing search query"}), 400

    limit     = min(int(request.args.get("limit",  50)),  200)
    sort      = request.args.get("sort",      "relevance")
    min_price = request.args.get("min",       "")
    max_price = request.args.get("max",       "")
    condition = request.args.get("condition", "")
    market    = request.args.get("market",    "EBAY_GB")

    filters = []
    if min_price:
        filters.append(f"price:[{min_price}]")
    if max_price:
        filters.append(f"price:[..{max_price}]")
    if condition:
        filters.append(f"conditions:{{{condition}}}")

    params = {"q": q, "limit": limit}
    if sort != "relevance":
        params["sort"] = sort
    if filters:
        params["filter"] = ",".join(filters)

    try:
        token = get_token()
        resp  = requests.get(
            f"{_ebay_base_url()}/buy/browse/v1/item_summary/search",
            headers={
                "Authorization":           f"Bearer {token}",
                "X-EBAY-C-MARKETPLACE-ID": market,
                "Content-Type":            "application/json",
            },
            params=params,
            timeout=15,
        )
        resp.raise_for_status()
        return jsonify(resp.json())
    except requests.HTTPError as e:
        try:
            detail = e.response.json()
        except Exception:
            detail = {"raw": e.response.text}
        return jsonify({"error": str(e), "detail": detail}), e.response.status_code
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/seller")
def seller_search():
    username = request.args.get("username", "").strip()
    if not username:
        return jsonify({"error": "Missing username parameter"}), 400

    limit  = min(int(request.args.get("limit", 200)), 200)
    market = request.args.get("market", "EBAY_GB")

    # ── Attempt 1: Browse API seller filter ─────────────────────────────────
    try:
        token = get_token()
        resp  = requests.get(
            f"{_ebay_base_url()}/buy/browse/v1/item_summary/search",
            headers={
                "Authorization":           f"Bearer {token}",
                "X-EBAY-C-MARKETPLACE-ID": market,
                "Content-Type":            "application/json",
            },
            params={"q": username, "filter": f"sellers:{{{username}}}", "limit": limit},
            timeout=15,
        )
        if resp.ok:
            items = resp.json().get("itemSummaries", [])
            if items:
                return jsonify({"itemSummaries": items, "source": "browse_filter", "total": len(items)})
    except Exception:
        pass

    # ── Attempt 2: Finding API findItemsFromSeller ───────────────────────────
    global_id_map = {
        "EBAY_GB": "EBAY-GB", "EBAY_US": "EBAY-US",
        "EBAY_DE": "EBAY-DE", "EBAY_FR": "EBAY-FR", "EBAY_AU": "EBAY-AU",
    }
    finding_base = {
        "OPERATION-NAME":       "findItemsFromSeller",
        "SERVICE-VERSION":      "1.13.0",
        "SECURITY-APPNAME":     EBAY_CLIENT_ID,
        "RESPONSE-DATA-FORMAT": "JSON",
        "GLOBAL-ID":            global_id_map.get(market, "EBAY-GB"),
        "itemFilter(0).name":   "Seller",
        "itemFilter(0).value":  username,
        "sortOrder":            "CurrentPriceHighest",
    }
    try:
        all_items = []
        pages     = 2 if limit > 100 else 1
        for page in range(1, pages + 1):
            r = requests.get(
                "https://svcs.ebay.com/services/search/FindingService/v1",
                params={**finding_base, "paginationInput.pageSize": 100, "paginationInput.pageNumber": page},
                timeout=15,
            )
            if not r.ok:
                break
            result = r.json().get("findItemsFromSellerResponse", [{}])[0]
            if result.get("ack", [""])[0] != "Success":
                break
            raw = result.get("searchResult", [{}])[0].get("item", [])
            all_items.extend(_normalize_finding(i) for i in raw)
            if len(raw) < 100:
                break
        if all_items:
            return jsonify({"itemSummaries": all_items[:limit], "source": "finding_api", "total": len(all_items)})
    except Exception:
        pass

    # ── Attempt 3: Keyword fallback ──────────────────────────────────────────
    try:
        token = get_token()
        resp  = requests.get(
            f"{_ebay_base_url()}/buy/browse/v1/item_summary/search",
            headers={
                "Authorization":           f"Bearer {token}",
                "X-EBAY-C-MARKETPLACE-ID": market,
                "Content-Type":            "application/json",
            },
            params={"q": username, "limit": limit},
            timeout=15,
        )
        resp.raise_for_status()
        all_items = resp.json().get("itemSummaries", [])
        filtered  = [
            i for i in all_items
            if (i.get("seller", {}).get("username") or "").lower() == username.lower()
        ]
        return jsonify({"itemSummaries": filtered, "source": "keyword_fallback", "total": len(filtered)})
    except requests.HTTPError as e:
        try:
            detail = e.response.json()
        except Exception:
            detail = {"raw": e.response.text}
        return jsonify({"error": str(e), "detail": detail}), e.response.status_code
    except Exception as e:
        return jsonify({"error": str(e)}), 500


def _normalize_finding(item):
    price_obj  = item.get("sellingStatus", [{}])[0].get("currentPrice", [{}])[0]
    price_val  = price_obj.get("__value__", "0")
    price_cur  = price_obj.get("@currencyId", "GBP")
    cond       = item.get("condition",  [{}])[0].get("conditionDisplayName", [""])[0]
    seller_obj = item.get("seller",     [{}])[0]
    s_name     = seller_obj.get("sellerUserName", [""])[0]
    s_fb_raw   = seller_obj.get("feedbackScore",  ["0"])[0]
    s_fb       = int(s_fb_raw) if str(s_fb_raw).isdigit() else 0
    ship_obj   = item.get("shippingInfo", [{}])[0]
    ship_cost  = ship_obj.get("shippingServiceCost", [{}])[0].get("__value__", None)
    ship_cur   = ship_obj.get("shippingServiceCost", [{}])[0].get("@currencyId", price_cur)
    title      = item.get("title",       [""])[0]
    url        = item.get("viewItemURL", ["#"])[0]
    item_id    = item.get("itemId",      [""])[0]
    gallery    = item.get("galleryURL",  [""])[0]

    normalized = {
        "itemId":     item_id,
        "title":      title,
        "price":      {"value": price_val, "currency": price_cur},
        "condition":  cond,
        "seller":     {"username": s_name, "feedbackScore": s_fb},
        "itemWebUrl": url,
    }
    if ship_cost is not None:
        normalized["shippingOptions"] = [{"shippingCost": {"value": ship_cost, "currency": ship_cur}}]
    if gallery:
        normalized["image"] = {"imageUrl": gallery}
    return normalized


# ═══════════════════════════════════════════════════════════════════════════════
# Routes — watchlist CRUD
# (scan routes registered first so Flask prefers them over /<id> matches)
# ═══════════════════════════════════════════════════════════════════════════════

@app.route("/watchlist/scan/status", methods=["GET"])
def scan_status():
    return jsonify(_scan_status)


@app.route("/watchlist/scan", methods=["POST"])
def trigger_scan():
    result, status = run_scan()
    return jsonify(result), status


@app.route("/watchlist", methods=["GET"])
def get_watchlist():
    with _watchlist_lock:
        items = load_watchlist()
    return jsonify(items)


@app.route("/watchlist", methods=["POST"])
def add_watchlist_item():
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"error": "JSON body required"}), 400

    missing = [f for f in ("part_name", "search_term", "my_price", "cost_price") if f not in data]
    if missing:
        return jsonify({"error": f"Missing fields: {', '.join(missing)}"}), 400

    item = {
        "id":                      str(uuid.uuid4()),
        "part_name":               str(data["part_name"]),
        "search_term":             str(data["search_term"]),
        "my_price":                float(data["my_price"]),
        "cost_price":              float(data["cost_price"]),
        "min_margin_percent":      float(data.get("min_margin_percent", 20)),
        "alert_threshold_percent": float(data.get("alert_threshold_percent", 2)),
        "market":                  str(data.get("market", "EBAY_GB")),
        "last_price":              None,
        "last_checked":            None,
        "price_history":           [],
        "active":                  bool(data.get("active", True)),
    }

    with _watchlist_lock:
        items = load_watchlist()
        items.append(item)
        save_watchlist(items)

    return jsonify(item), 201


@app.route("/watchlist/<item_id>", methods=["PUT"])
def update_watchlist_item(item_id):
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"error": "JSON body required"}), 400

    numeric_fields = {"my_price", "cost_price", "min_margin_percent", "alert_threshold_percent"}
    allowed_fields = numeric_fields | {"part_name", "search_term", "market", "active"}

    with _watchlist_lock:
        items = load_watchlist()
        for i, item in enumerate(items):
            if item["id"] == item_id:
                for field in allowed_fields:
                    if field in data:
                        val = data[field]
                        if field in numeric_fields:
                            val = float(val)
                        elif field == "active":
                            val = bool(val)
                        items[i][field] = val
                save_watchlist(items)
                return jsonify(items[i])

    return jsonify({"error": "Item not found"}), 404


@app.route("/watchlist/<item_id>", methods=["DELETE"])
def delete_watchlist_item(item_id):
    with _watchlist_lock:
        items    = load_watchlist()
        filtered = [it for it in items if it["id"] != item_id]
        if len(filtered) == len(items):
            return jsonify({"error": "Item not found"}), 404
        save_watchlist(filtered)
    return jsonify({"deleted": item_id})


# ═══════════════════════════════════════════════════════════════════════════════
# Routes — Taxonomy API (category search + item specifics)
# ═══════════════════════════════════════════════════════════════════════════════

_category_tree_cache = {}

REQUIREMENT_ORDER = {"REQUIRED": 0, "RECOMMENDED": 1, "OPTIONAL": 2}


def get_category_tree_id(marketplace):
    if marketplace in _category_tree_cache:
        return _category_tree_cache[marketplace]

    token = get_token()
    resp  = requests.get(
        f"{_ebay_base_url()}/commerce/taxonomy/v1/get_default_category_tree_id",
        headers={"Authorization": f"Bearer {token}"},
        params={"marketplace_id": marketplace},
        timeout=10,
    )
    resp.raise_for_status()
    tree_id = resp.json()["categoryTreeId"]
    _category_tree_cache[marketplace] = tree_id
    return tree_id


def _build_breadcrumb(category_name, ancestors):
    ordered = sorted(ancestors, key=lambda a: a.get("categoryTreeNodeLevel", 0))
    return " > ".join([a["categoryName"] for a in ordered] + [category_name])


@app.route("/api/taxonomy/search-categories", methods=["GET"])
def search_categories():
    q = request.args.get("q", "").strip()
    if not q:
        return jsonify({"error": "Missing search query"}), 400

    marketplace = request.args.get("marketplace", "EBAY_GB")

    try:
        tree_id = get_category_tree_id(marketplace)
        token   = get_token()
        resp    = requests.get(
            f"{_ebay_base_url()}/commerce/taxonomy/v1/category_tree/{tree_id}/get_category_suggestions",
            headers={"Authorization": f"Bearer {token}"},
            params={"q": q},
            timeout=15,
        )
        resp.raise_for_status()

        suggestions = []
        for s in resp.json().get("categorySuggestions", []):
            category = s.get("category", {})
            suggestions.append({
                "categoryId":   category.get("categoryId"),
                "categoryName": category.get("categoryName"),
                "breadcrumb":   _build_breadcrumb(
                    category.get("categoryName", ""),
                    s.get("categoryTreeNodeAncestors", []),
                ),
            })

        return jsonify(suggestions)
    except requests.HTTPError as e:
        try:
            detail = e.response.json()
        except Exception:
            detail = {"raw": e.response.text}
        return jsonify({"error": str(e), "detail": detail}), e.response.status_code
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/taxonomy/item-specifics", methods=["GET"])
def item_specifics():
    category_id = request.args.get("category_id", "").strip()
    if not category_id:
        return jsonify({"error": "Missing category_id parameter"}), 400

    marketplace = request.args.get("marketplace", "EBAY_GB")

    try:
        tree_id = get_category_tree_id(marketplace)
        token   = get_token()
        resp    = requests.get(
            f"{_ebay_base_url()}/commerce/taxonomy/v1/category_tree/{tree_id}/get_item_aspects_for_category",
            headers={"Authorization": f"Bearer {token}"},
            params={"category_id": category_id},
            timeout=15,
        )
        resp.raise_for_status()

        aspects = []
        for a in resp.json().get("aspects", []):
            constraint = a.get("aspectConstraint", {})
            required   = bool(constraint.get("aspectRequired"))
            usage      = constraint.get("aspectUsage", "OPTIONAL")
            level      = "REQUIRED" if required else ("RECOMMENDED" if usage == "RECOMMENDED" else "OPTIONAL")

            aspects.append({
                "name":                   a.get("localizedAspectName", ""),
                "requirementLevel":       level,
                "dataType":               constraint.get("aspectDataType", ""),
                "mode":                   constraint.get("aspectMode", ""),
                "supportsMultipleValues": constraint.get("itemToAspectCardinality") == "MULTI",
                "enabledForVariations":   bool(constraint.get("aspectEnabledForVariations")),
                "allowedValues":          [v.get("localizedValue", "") for v in a.get("aspectValues", [])],
                "searchCount":            a.get("relevanceIndicator", {}).get("searchCount", 0),
            })

        aspects.sort(key=lambda a: REQUIREMENT_ORDER.get(a["requirementLevel"], 3))

        summary = {
            "total":       len(aspects),
            "required":    sum(1 for a in aspects if a["requirementLevel"] == "REQUIRED"),
            "recommended": sum(1 for a in aspects if a["requirementLevel"] == "RECOMMENDED"),
            "optional":    sum(1 for a in aspects if a["requirementLevel"] == "OPTIONAL"),
        }

        return jsonify({"aspects": aspects, "summary": summary})
    except requests.HTTPError as e:
        try:
            detail = e.response.json()
        except Exception:
            detail = {"raw": e.response.text}
        return jsonify({"error": str(e), "detail": detail}), e.response.status_code
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
