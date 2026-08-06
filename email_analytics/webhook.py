from __future__ import annotations

import logging
from datetime import datetime, timezone
from typing import Any
from uuid import uuid4

import uvicorn
from fastapi import BackgroundTasks, FastAPI, HTTPException, Query, Request
from fastapi.responses import PlainTextResponse, Response
from pydantic import BaseModel, Field

from .config import get_settings
from .database import initialize_database, upsert_emails
from .graph_client import GraphMailClient


logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)
app = FastAPI(title="Email Analytics Graph Webhook", docs_url=None, redoc_url=None)


class DemoEmail(BaseModel):
    sender_email: str = Field(default="demo.sender@example.com", max_length=320)
    sender_name: str = Field(default="Demo Sender", max_length=200)
    subject: str = Field(default="Live trigger demo email", max_length=500)
    body_preview: str = Field(default="This is a locally simulated incoming email.", max_length=2000)
    has_attachments: bool = False


def graph_client() -> GraphMailClient:
    settings = get_settings()
    return GraphMailClient(
        settings.client_id,
        settings.tenant_id,
        settings.database_path.parent / "token_cache.bin",
        settings.store_body_preview,
    )


def save_graph_message(message_id: str) -> None:
    """Fetch the email after Graph notifies us and save it to SQLite."""
    settings = get_settings()
    try:
        message = graph_client().get_message(message_id)
        stored = upsert_emails(settings.database_path, [message])
        logger.info("Graph message %s saved (%s row affected).", message_id, stored)
    except Exception:
        logger.exception("Could not process Graph message %s. Scheduled sync will catch it later.", message_id)


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/webhooks/graph")
async def graph_notifications(
    request: Request,
    background_tasks: BackgroundTasks,
    validationToken: str | None = Query(default=None),
) -> Response:
    """Receive Graph validation and real change-notification POSTs."""
    if validationToken is not None:
        # Graph requires an unencoded plain-text response within 10 seconds.
        return PlainTextResponse(validationToken)

    payload: dict[str, Any] = await request.json()
    expected_state = get_settings().webhook_client_state
    if not expected_state:
        raise HTTPException(status_code=503, detail="WEBHOOK_CLIENT_STATE is not configured.")

    accepted = 0
    for notification in payload.get("value", []):
        if notification.get("clientState") != expected_state:
            logger.warning("Rejected a notification with an invalid clientState.")
            continue
        if notification.get("changeType") != "created":
            continue
        message_id = (notification.get("resourceData") or {}).get("id")
        if not message_id:
            continue
        background_tasks.add_task(save_graph_message, message_id)
        accepted += 1

    logger.info("Accepted %s Graph notification(s).", accepted)
    return Response(status_code=202)


@app.post("/demo/new-email", status_code=202)
def demo_new_email(demo_email: DemoEmail) -> dict[str, str]:
    """Local-only way to prove immediate upsert + dashboard refresh without Outlook access."""
    settings = get_settings()
    initialize_database(settings.database_path)
    message_id = f"simulated-trigger-{uuid4()}"
    message = {
        "id": message_id,
        "from": {"emailAddress": {"name": demo_email.sender_name, "address": demo_email.sender_email}},
        "subject": demo_email.subject,
        "receivedDateTime": datetime.now(timezone.utc).isoformat(),
        "bodyPreview": demo_email.body_preview if settings.store_body_preview else "",
        "hasAttachments": demo_email.has_attachments,
        "importance": "normal",
        "conversationId": f"simulated-conversation-{uuid4()}",
        "webLink": None,
    }
    upsert_emails(settings.database_path, [message])
    logger.info("Saved a simulated trigger email: %s", message_id)
    return {"status": "saved", "message_id": message_id}


@app.post("/webhooks/graph/lifecycle", status_code=202)
async def lifecycle_notifications(
    request: Request, validationToken: str | None = Query(default=None)
) -> Response:
    if validationToken is not None:
        return PlainTextResponse(validationToken)
    logger.warning("Received a Graph lifecycle notification: %s", await request.json())
    return Response(status_code=202)


def main() -> None:
    settings = get_settings()
    initialize_database(settings.database_path)
    uvicorn.run(app, host="127.0.0.1", port=settings.webhook_port, log_level="info")


if __name__ == "__main__":
    main()
