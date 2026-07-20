"""
app.py

Deploy this on Azure App Service. This is the public HTTPS endpoint that
Azure Bot Service's Teams calling webhook points to, and that you pass as
callbackUri when placing the call from the private server.

FOR TESTING ONLY: it accepts Microsoft Graph's call-state notifications and
just logs them -- it does not forward anything to the private server. This
satisfies Graph's requirement for a real, reachable callbackUri while you
validate the join by watching the Teams meeting directly, or by polling from
the private server's initiate_call.py.

Before using this in anything beyond a quick test, add validation of the
bearer token Graph sends in the Authorization header of each notification
(see Microsoft's "Configure incoming call notifications" docs) -- this stub
skips that check for simplicity.
"""

import logging
import os

from flask import Flask, request, jsonify

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)


@app.route("/api/calling", methods=["POST"])
def calling_webhook():
    body = request.get_json(silent=True) or {}

    for notification in body.get("value", []):
        resource = notification.get("resourceUrl", "unknown resource")
        state = notification.get("resourceData", {}).get("state")
        app.logger.info(f"Notification for {resource}: state={state}")

    # Graph expects a fast response -- don't do slow work in this handler.
    return jsonify({"status": "received"}), 200


@app.route("/health", methods=["GET"])
def health():
    return "ok", 200


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port)
