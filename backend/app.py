import os
import time
import base64
import requests
from flask import Flask, jsonify, request
from flask_cors import CORS

app = Flask(__name__)
CORS(app)  # Allow requests from your GitHub Pages frontend

# --- eBay credentials (set these as environment variables on Railway) ---
EBAY_CLIENT_ID     = os.environ.get("EBAY_CLIENT_ID", "")
EBAY_CLIENT_SECRET = os.environ.get("EBAY_CLIENT_SECRET", "")
EBAY_ENV           = os.environ.get("EBAY_ENV", "production")  # "sandbox" or "production"

# Token cache
_token_cache = {"token": None, "expires_at": 0}


def get_token():
    """Fetch or return cached eBay OAuth token."""
    now = time.time()
    if _token_cache["token"] and now < _token_cache["expires_at"] - 60:
        return _token_cache["token"]

    if EBAY_ENV == "sandbox":
        url = "https://api.sandbox.ebay.com/identity/v1/oauth2/token"
    else:
        url = "https://api.ebay.com/identity/v1/oauth2/token"

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


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


@app.route("/search")
def search():
    """
    Proxy eBay Browse API search.
    Query params:
        q        – search term (required)
        limit    – number of results (default 50, max 200)
        sort     – relevance | price | -price | -date
        min      – minimum price filter
        max      – maximum price filter
        condition – NEW | USED | UNSPECIFIED
        market   – eBay marketplace ID (default EBAY_GB)
    """
    q = request.args.get("q", "").strip()
    if not q:
        return jsonify({"error": "Missing search query"}), 400

    limit     = min(int(request.args.get("limit",  50)),  200)
    sort      = request.args.get("sort",      "relevance")
    min_price = request.args.get("min",       "")
    max_price = request.args.get("max",       "")
    condition = request.args.get("condition", "")
    market    = request.args.get("market",    "EBAY_GB")

    # Build filter string
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

    if EBAY_ENV == "sandbox":
        base_url = "https://api.sandbox.ebay.com"
    else:
        base_url = "https://api.ebay.com"

    try:
        token = get_token()
        resp  = requests.get(
            f"{base_url}/buy/browse/v1/item_summary/search",
            headers={
                "Authorization":            f"Bearer {token}",
                "X-EBAY-C-MARKETPLACE-ID":  market,
                "Content-Type":             "application/json",
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
    """
    Fetch listings for a specific eBay seller.

    Strategy (each tried in order, stops at first success):
      1. Browse API  – filter=sellers:{username}   (requires HQ access; may return 0)
      2. Finding API – findItemsFromSeller          (App-ID only, works at all tiers)
      3. Keyword fallback – Browse API q=username, filtered client-side

    Response shape matches /search so the frontend needs no format change:
      { itemSummaries: [...], source: "browse_filter|finding_api|keyword_fallback", total: N }
    """
    username = request.args.get("username", "").strip()
    if not username:
        return jsonify({"error": "Missing username parameter"}), 400

    limit  = min(int(request.args.get("limit", 200)), 200)
    market = request.args.get("market", "EBAY_GB")

    base_url = (
        "https://api.sandbox.ebay.com" if EBAY_ENV == "sandbox"
        else "https://api.ebay.com"
    )

    # ── Attempt 1: Browse API seller filter ─────────────────
    try:
        token = get_token()
        resp  = requests.get(
            f"{base_url}/buy/browse/v1/item_summary/search",
            headers={
                "Authorization":           f"Bearer {token}",
                "X-EBAY-C-MARKETPLACE-ID": market,
                "Content-Type":            "application/json",
            },
            params={
                "q":      username,
                "filter": f"sellers:{{{username}}}",
                "limit":  limit,
            },
            timeout=15,
        )
        if resp.ok:
            items = resp.json().get("itemSummaries", [])
            if items:
                return jsonify({
                    "itemSummaries": items,
                    "source":        "browse_filter",
                    "total":         len(items),
                })
    except Exception:
        pass  # fall through to next strategy

    # ── Attempt 2: Finding API findItemsFromSeller ───────────
    global_id_map = {
        "EBAY_GB": "EBAY-GB", "EBAY_US": "EBAY-US",
        "EBAY_DE": "EBAY-DE", "EBAY_FR": "EBAY-FR", "EBAY_AU": "EBAY-AU",
    }
    global_id    = global_id_map.get(market, "EBAY-GB")
    finding_url  = "https://svcs.ebay.com/services/search/FindingService/v1"
    finding_base = {
        "OPERATION-NAME":       "findItemsFromSeller",
        "SERVICE-VERSION":      "1.13.0",
        "SECURITY-APPNAME":     EBAY_CLIENT_ID,
        "RESPONSE-DATA-FORMAT": "JSON",
        "GLOBAL-ID":            global_id,
        "itemFilter(0).name":   "Seller",
        "itemFilter(0).value":  username,
        "sortOrder":            "CurrentPriceHighest",
    }

    try:
        all_items = []
        pages     = 2 if limit > 100 else 1

        for page in range(1, pages + 1):
            params = {
                **finding_base,
                "paginationInput.pageSize":   100,
                "paginationInput.pageNumber": page,
            }
            r = requests.get(finding_url, params=params, timeout=15)
            if not r.ok:
                break
            data   = r.json()
            result = data.get("findItemsFromSellerResponse", [{}])[0]
            if result.get("ack", [""])[0] != "Success":
                break
            raw = result.get("searchResult", [{}])[0].get("item", [])
            all_items.extend(_normalize_finding(i) for i in raw)
            if len(raw) < 100:
                break  # last page

        if all_items:
            return jsonify({
                "itemSummaries": all_items[:limit],
                "source":        "finding_api",
                "total":         len(all_items),
            })
    except Exception:
        pass  # fall through to keyword fallback

    # ── Attempt 3: Keyword fallback ──────────────────────────
    try:
        token = get_token()
        resp  = requests.get(
            f"{base_url}/buy/browse/v1/item_summary/search",
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
        # Keep only results whose seller username matches
        filtered = [
            i for i in all_items
            if (i.get("seller", {}).get("username") or "").lower() == username.lower()
        ]
        return jsonify({
            "itemSummaries": filtered,
            "source":        "keyword_fallback",
            "total":         len(filtered),
        })
    except requests.HTTPError as e:
        try:
            detail = e.response.json()
        except Exception:
            detail = {"raw": e.response.text}
        return jsonify({"error": str(e), "detail": detail}), e.response.status_code
    except Exception as e:
        return jsonify({"error": str(e)}), 500


def _normalize_finding(item):
    """Convert a Finding API item dict to Browse API item_summary shape."""
    def _first(d, *keys, default=""):
        for k in keys:
            d = d.get(k, [{}])[0] if isinstance(d, dict) else {}
        return d if not isinstance(d, dict) else default

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

    title   = item.get("title",       [""])[0]
    url     = item.get("viewItemURL", ["#"])[0]
    item_id = item.get("itemId",      [""])[0]
    gallery = item.get("galleryURL",  [""])[0]

    normalized = {
        "itemId":     item_id,
        "title":      title,
        "price":      {"value": price_val, "currency": price_cur},
        "condition":  cond,
        "seller":     {"username": s_name, "feedbackScore": s_fb},
        "itemWebUrl": url,
    }
    if ship_cost is not None:
        normalized["shippingOptions"] = [
            {"shippingCost": {"value": ship_cost, "currency": ship_cur}}
        ]
    if gallery:
        normalized["image"] = {"imageUrl": gallery}
    return normalized


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
