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


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
