import os
import requests
from flask import Flask, request

app = Flask(__name__)

TELEGRAM_TOKEN = os.environ["TELEGRAM_TOKEN"]
GITHUB_TOKEN = os.environ["GITHUB_TOKEN"]

# Repository 1 — existing NSE Intraday PRO
REPO_1 = "jogikrunal189-del/nse-intraday-picks"
WORKFLOW_1 = "main.yml"

# Repository 2 — new V3 forward-test repository
# Change V3_REPO if your actual GitHub repo name is different.
REPO_2 = "jogikrunal189-del/V3"
WORKFLOW_2 = "main.yml"


def send_telegram(chat_id, message):
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"

    try:
        response = requests.post(
            url,
            json={
                "chat_id": chat_id,
                "text": message,
            },
            timeout=20,
        )
        response.raise_for_status()
    except Exception as exc:
        print(f"Telegram message error: {exc}")


def trigger_workflow(repo, workflow, label):
    url = (
        f"https://api.github.com/repos/{repo}"
        f"/actions/workflows/{workflow}/dispatches"
    )

    headers = {
        "Accept": "application/vnd.github+json",
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "X-GitHub-Api-Version": "2022-11-28",
    }

    try:
        response = requests.post(
            url,
            headers=headers,
            json={"ref": "main"},
            timeout=30,
        )

        if response.status_code == 204:
            return True, f"✅ {label} GitHub Action started."
        return False, (
            f"❌ {label} failed to start.\n"
            f"HTTP status: {response.status_code}\n"
            f"Response: {response.text[:500]}"
        )

    except Exception as exc:
        return False, f"❌ {label} request error: {exc}"


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
            "NSE Runner is online.\n\n"
            "Use /run to start both NSE scanners."
        )

    elif text == "/run":
        ok1, msg1 = trigger_workflow(
            REPO_1,
            WORKFLOW_1,
            "NSE Intraday PRO",
        )

        ok2, msg2 = trigger_workflow(
            REPO_2,
            WORKFLOW_2,
            "V3 Forward Test",
        )

        send_telegram(
            chat_id,
            "🚀 Scanner run requested\n\n"
            f"{msg1}\n\n"
            f"{msg2}"
        )

    else:
        send_telegram(
            chat_id,
            "Use /run to start both NSE scanners."
        )

    return "OK"


if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 10000)),
    )
