
import os
import requests
from flask import Flask, request

app = Flask(__name__)

# =====================================================
# ENVIRONMENT VARIABLES
# =====================================================

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "").strip()
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN", "").strip()

if not TELEGRAM_TOKEN:
    raise RuntimeError("Missing TELEGRAM_TOKEN in Render Environment")

if not GITHUB_TOKEN:
    raise RuntimeError("Missing GITHUB_TOKEN in Render Environment")


# =====================================================
# GITHUB REPOSITORY CONFIGURATION
# =====================================================

# Repository 1: NSE Intraday PRO
REPO_1 = "jogikrunal189-del/nse-intraday-picks"
WORKFLOW_1 = "main.yml"

# Repository 2: V3 Forward Test
REPO_2 = "jogikrunal189-del/V3"
WORKFLOW_2 = "main.yml"

BRANCH = "main"

GITHUB_API = "https://api.github.com"


# =====================================================
# GITHUB HEADERS
# =====================================================

GITHUB_HEADERS = {
    "Accept": "application/vnd.github+json",
    "Authorization": f"Bearer {GITHUB_TOKEN}",
    "X-GitHub-Api-Version": "2022-11-28",
}


# =====================================================
# TELEGRAM MESSAGE
# =====================================================

def send_telegram(chat_id, message):

    url = (
        f"https://api.telegram.org/"
        f"bot{TELEGRAM_TOKEN}/sendMessage"
    )

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


# =====================================================
# CHECK GITHUB AUTHENTICATION
# =====================================================

def check_github_auth():

    print("\n" + "=" * 60)
    print("CHECKING GITHUB AUTHENTICATION")
    print("=" * 60)

    repositories = [
        REPO_1,
        REPO_2,
    ]

    auth_ok = True

    for repo in repositories:

        url = f"{GITHUB_API}/repos/{repo}"

        try:

            response = requests.get(
                url,
                headers=GITHUB_HEADERS,
                timeout=20,
            )

            if response.status_code == 200:

                print(f"AUTH OK: {repo}")

            elif response.status_code == 401:

                print(
                    f"AUTH FAILED: {repo}\n"
                    "HTTP 401: Invalid or rejected GitHub token."
                )

                auth_ok = False

            elif response.status_code == 403:

                print(
                    f"ACCESS DENIED: {repo}\n"
                    f"Response: {response.text[:300]}"
                )

                auth_ok = False

            elif response.status_code == 404:

                print(
                    f"REPOSITORY NOT FOUND: {repo}\n"
                    "Check repository name and token access."
                )

                auth_ok = False

            else:

                print(
                    f"ERROR: {repo}\n"
                    f"HTTP: {response.status_code}\n"
                    f"Response: {response.text[:300]}"
                )

                auth_ok = False

        except requests.RequestException as exc:

            print(f"CONNECTION ERROR: {repo}: {exc}")

            auth_ok = False

    print("=" * 60)

    return auth_ok


# =====================================================
# TRIGGER GITHUB ACTION
# =====================================================

def trigger_workflow(repo, workflow, label):

    url = (
        f"{GITHUB_API}/repos/{repo}"
        f"/actions/workflows/{workflow}/dispatches"
    )

    payload = {
        "ref": BRANCH
    }

    try:

        response = requests.post(
            url,
            headers=GITHUB_HEADERS,
            json=payload,
            timeout=30,
        )

        # GitHub returns HTTP 204 when dispatch succeeds.
        if response.status_code == 204:

            print(f"SUCCESS: {label}")

            return True, (
                f"✅ {label} started successfully."
            )

        # Authentication error.
        elif response.status_code == 401:

            print(f"401 AUTHENTICATION ERROR: {label}")

            return False, (
                f"❌ {label} failed.\n"
                "HTTP 401: GitHub rejected the token.\n"
                "Check GITHUB_TOKEN in Render."
            )

        # Permission error.
        elif response.status_code == 403:

            print(f"403 PERMISSION ERROR: {label}")

            return False, (
                f"❌ {label} failed.\n"
                "HTTP 403: Insufficient repository permissions.\n"
                "Check Actions permission: Read and write."
            )

        # Repository/workflow not found.
        elif response.status_code == 404:

            print(f"404 NOT FOUND: {label}")

            return False, (
                f"❌ {label} failed.\n"
                "HTTP 404: Repository or workflow not found.\n"
                f"Repository: {repo}\n"
                f"Workflow: {workflow}"
            )

        else:

            print(
                f"WORKFLOW ERROR: {label}\n"
                f"HTTP: {response.status_code}\n"
                f"Response: {response.text[:300]}"
            )

            return False, (
                f"❌ {label} failed.\n"
                f"HTTP status: {response.status_code}\n"
                f"Response: {response.text[:300]}"
            )

    except requests.RequestException as exc:

        print(f"REQUEST ERROR: {label}: {exc}")

        return False, (
            f"❌ {label} connection error:\n{str(exc)}"
        )


# =====================================================
# HOME ROUTE
# =====================================================

@app.route("/")
def home():

    return (
        "NSE Telegram Runner is online!\n"
        "Both scanner workflows are configured."
    )


# =====================================================
# HEALTH CHECK
# =====================================================

@app.route("/health")
def health():

    return {
        "status": "online",
        "service": "NSE Telegram Runner",
    }, 200


# =====================================================
# TELEGRAM WEBHOOK
# =====================================================

@app.route("/telegram", methods=["POST"])
def telegram():

    data = request.get_json(silent=True) or {}

    message = data.get("message", {})

    chat = message.get("chat", {})

    chat_id = chat.get("id")

    text = message.get("text", "").strip()

    if not chat_id:
        return "OK", 200

    # -------------------------------------------------
    # START COMMAND
    # -------------------------------------------------

    if text == "/start":

        send_telegram(
            chat_id,
            "🚀 NSE Telegram Runner is online!\n\n"
            "Available command:\n"
            "/run - Start both NSE scanners."
        )

    # -------------------------------------------------
    # RUN COMMAND
    # -------------------------------------------------

    elif text == "/run":

        send_telegram(
            chat_id,
            "🚀 NSE scanner run requested.\n\n"
            "Checking GitHub authentication..."
        )

        # Test authentication first.
        if not check_github_auth():

            send_telegram(
                chat_id,
                "❌ GitHub authentication/access check failed.\n\n"
                "Please check:\n"
                "1. GITHUB_TOKEN in Render.\n"
                "2. Token repository access.\n"
                "3. Token expiration.\n"
                "4. Repository names.\n\n"
                "Check Render logs for details."
            )

            return "OK", 200

        # Trigger scanner 1.
        ok1, msg1 = trigger_workflow(
            REPO_1,
            WORKFLOW_1,
            "NSE Intraday PRO",
        )

        # Trigger scanner 2.
        ok2, msg2 = trigger_workflow(
            REPO_2,
            WORKFLOW_2,
            "V3 Forward Test",
        )

        # Final Telegram report.
        final_message = (
            "🚀 NSE SCANNER RESULTS\n\n"
            f"{msg1}\n\n"
            f"{msg2}\n\n"
            f"Successful workflows: {int(ok1) + int(ok2)}/2"
        )

        send_telegram(
            chat_id,
            final_message,
        )

    # -------------------------------------------------
    # UNKNOWN COMMAND
    # -------------------------------------------------

    else:

        send_telegram(
            chat_id,
            "Unknown command.\n\n"
            "Use /run to start both NSE scanners."
        )

    return "OK", 200


# =====================================================
# START APPLICATION
# =====================================================

if __name__ == "__main__":

    print("\n" + "=" * 60)
    print("NSE TELEGRAM GITHUB RUNNER")
    print("=" * 60)

    print("Repository 1:", REPO_1)
    print("Workflow 1:", WORKFLOW_1)

    print("Repository 2:", REPO_2)
    print("Workflow 2:", WORKFLOW_2)

    print("Branch:", BRANCH)

    # Authentication check at startup.
    check_github_auth()

    print("\nStarting Flask application...")

    app.run(
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 10000)),
    )
