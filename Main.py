from flask import Flask, render_template, request, jsonify, send_file
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError
import pandas as pd
from datetime import datetime
import os
import re

app = Flask(__name__)

# ✅ Load secrets from environment variables
slack_token = os.environ.get("SLACK_BOT_TOKEN")
channel_id = os.environ.get("SLACK_CHANNEL_ID")
client = WebClient(token=slack_token)

# ✅ Extractor function with smart parsing
def extract_data(text, user):
    data = {
        "posted_by": user,
        "job_number": "",
        "phone_number": "",
        "pickup": "",
        "dropoff": "",
        "price": "",
        "drv_number": "",
        "reason": "",
        "notes": ""
    }

    job_match = re.search(r"\b\d{7,8}[A-Z]?\b", text)
    if job_match:
        data["job_number"] = job_match.group()

    phone_match = re.search(r"\b0\d{9,10}\b", text)
    if phone_match:
        data["phone_number"] = phone_match.group()

    price_match = re.search(r"£\s?(\d+(?:\.\d{1,2})?)", text)
    if not price_match:
        price_match = re.search(r"(\d+(?:\.\d{1,2})?)£", text)
    if price_match:
        data["price"] = price_match.group(1)

    drv_match = re.search(r"(?:Drv|driver|owed to|to|owes)\s*(\d{3})", text, re.IGNORECASE)
    if drv_match:
        data["drv_number"] = drv_match.group(1)

    loc_match = re.search(r"from\s+(.*?)\s+to\s+([^\.,\n]+)", text, re.IGNORECASE)
    if loc_match:
        data["pickup"] = loc_match.group(1).strip()
        data["dropoff"] = loc_match.group(2).strip()
    else:
        arrow_match = re.search(r"(\S+)\s*->\s*(\S+)", text)
        if arrow_match:
            data["pickup"] = arrow_match.group(1).strip()
            data["dropoff"] = arrow_match.group(2).strip()

    reason_match = re.search(r"(said.*|card wasn't working.*|forgot.*|cleaning fee.*|didn't pay.*)", text, re.IGNORECASE)
    if reason_match:
        data["reason"] = reason_match.group().strip()

    missing = []
    for key in ["job_number", "phone_number", "price"]:
        if not data[key]:
            missing.append(key)
    if missing:
        data["notes"] = "Missing: " + ", ".join(missing)

    return data

# ✅ Home route (GUI page)
@app.route("/", methods=["GET", "POST"])
def home():
    file = None
    error = None

    if request.method == "POST":
        try:
            messages = []
            response = client.conversations_history(channel=channel_id, limit=100)

            for message in response["messages"]:
                if "text" in message and "user" in message:
                    user_id = message["user"]
                    text = message["text"]
                    user_info = client.users_info(user=user_id)
                    username = user_info["user"]["real_name"]

                    extracted = extract_data(text, username)
                    messages.append(extracted)

            df = pd.DataFrame(messages)
            filename = f"jobs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            filepath = f"/tmp/{filename}"
            df.to_excel(filepath, index=False)
            file = filename

        except SlackApiError as e:
            error = str(e)

    return render_template("index.html", file=file, error=error)

# ✅ Download the latest Excel file
@app.route("/download", methods=["GET"])
def download():
    files = [f for f in os.listdir("/tmp") if f.endswith(".xlsx")]
    if files:
        latest = sorted(files)[-1]
        return send_file(f"/tmp/{latest}", as_attachment=True)
    return "No Excel files found."

# ✅ Start the app
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=3000)
