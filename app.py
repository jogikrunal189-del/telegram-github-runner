import os
import requests
from flask import Flask, request

app = Flask(__name__)

TELEGRAM_TOKEN = os.environ["TELEGRAM_TOKEN"]
GITHUB_TOKEN = os.environ["GITHUB_TOKEN"]

REPO = "jogikrunal189-del/nse-intraday-picks"
WORKFLOW = "main.yml"


def send_telegram(chat_id, message):
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"

    requests.post(
        url,
        json={
            "chat_id": chat_id,
            "text": message
        },
        timeout=20
    )


@app.route("/")
def home():
    return "NSE Telegram Runner is online!"


@app.route("/telegram", methods=["POST"])
def telegram():
    data = request.get_json(silent=True) or {}

    message = data.get("message", {})
    chat = message.get("chat", {})
    chat_id = chat.get("id")
    text = message.get("text", "").strip()

    if not chat_id:
        return "OK"

    if text == "/start":
        send_telegram(
            chat_id,
            "NSE Runner is online.\n\nUse /run to start GitHub Actions."
        )

    elif text == "/run":
        url = (
            f"https://api.github.com/repos/{REPO}"
            f"/actions/workflows/{WORKFLOW}/dispatches"
        )

        headers = {
            "Accept": "application/vnd.github+json",
            "Authorization": f"Bearer {GITHUB_TOKEN}",
            "X-GitHub-Api-Version": "2022-11-28"
        }

        response = requests.post(
            url,
            headers=headers,
            json={"ref": "main"},
            timeout=30
        )

        if response.status_code == 204:
            send_telegram(
                chat_id,
                "✅ NSE GitHub Action started successfully!"
            )
        else:
            send_telegram(
                chat_id,
                f"❌ GitHub Action failed to start.\n"
                f"Status: {response.status_code}"
            )

    else:
        send_telegram(
            chat_id,
            "Use /run to start the NSE GitHub Action."
        )

    return "OK"


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
