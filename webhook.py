from flask import Flask, request, jsonify
import json
import os
import random
import string

app = Flask(__name__)

# Gunakan direktori lokal di aplikasi render untuk penyimpanan sementara tokens.json
TOKEN_FILE = os.path.join(os.path.dirname(__file__), 'tokens.json')

def _new_token(prefix="SDMAT"):
    chars = string.ascii_uppercase + string.digits
    chars = chars.replace('O','').replace('0','').replace('I','').replace('1','')
    part1 = ''.join(random.choice(chars) for _ in range(4))
    part2 = ''.join(random.choice(chars) for _ in range(4))
    return f"{prefix}-{part1}-{part2}"

def _store_token(token, buyer):
    if not os.path.exists(TOKEN_FILE):
        data = []
    else:
        try:
            with open(TOKEN_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            data = []

    new_id = max((item.get("id", 0) for item in data), default=0) + 1
    data.append({
        "id": new_id,
        "token": token,
        "status": "active",
        "created_for": buyer
    })
    with open(TOKEN_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

@app.route("/", methods=["GET"])
def home():
    return "SmartDrill Webhook Service is Running!", 200

@app.route("/lynk/webhook", methods=["POST"])
def lynk_hook():
    try:
        payload = request.get_json(force=True) or {}
        buyer = payload.get("buyer_name", "Customer")
        email = payload.get("buyer_email", "unknown@email.com")
        
        prefix = random.choice(["SDMAT", "SDBIND", "SDVIP"])
        token = _new_token(prefix)
        _store_token(token, buyer)
        
        # Di sini bisa ditambahkan trigger pengiriman email (opsional / via Mailgun/SendGrid)
        print(f"Generated token {token} for {buyer} ({email})")
        
        return jsonify({
            "status": "success",
            "token": token,
            "buyer": buyer
        }), 200
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
