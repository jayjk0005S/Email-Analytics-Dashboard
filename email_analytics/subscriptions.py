from __future__ import annotations

import argparse

from .config import get_settings
from .database import get_state, initialize_database, set_state
from .graph_client import GraphConfigurationError, GraphMailClient


def _validated_settings() -> tuple[str, str]:
    settings = get_settings()
    if not settings.public_webhook_url or not settings.public_webhook_url.startswith("https://"):
        raise GraphConfigurationError("PUBLIC_WEBHOOK_URL must be a publicly reachable HTTPS URL.")
    if not settings.webhook_client_state or len(settings.webhook_client_state) < 24:
        raise GraphConfigurationError(
            "WEBHOOK_CLIENT_STATE must be a private random value of at least 24 characters."
        )
    return settings.public_webhook_url.rstrip("/"), settings.webhook_client_state


def create_subscription() -> dict:
    settings = get_settings()
    initialize_database(settings.database_path)
    notification_url, client_state = _validated_settings()
    client = GraphMailClient(
        settings.client_id,
        settings.tenant_id,
        settings.database_path.parent / "token_cache.bin",
        settings.store_body_preview,
    )
    subscription = client.create_inbox_subscription(
        notification_url=notification_url,
        lifecycle_url=f"{notification_url}/lifecycle",
        client_state=client_state,
    )
    set_state(settings.database_path, "graph_subscription_id", subscription["id"])
    set_state(settings.database_path, "graph_subscription_expires_at", subscription["expirationDateTime"])
    return subscription


def renew_subscription() -> dict:
    settings = get_settings()
    initialize_database(settings.database_path)
    subscription_id = get_state(settings.database_path, "graph_subscription_id")
    if not subscription_id:
        raise GraphConfigurationError("No saved subscription. Run `python -m email_analytics.subscriptions create` first.")
    client = GraphMailClient(
        settings.client_id,
        settings.tenant_id,
        settings.database_path.parent / "token_cache.bin",
        settings.store_body_preview,
    )
    subscription = client.renew_subscription(subscription_id)
    set_state(settings.database_path, "graph_subscription_expires_at", subscription["expirationDateTime"])
    return subscription


def main() -> None:
    parser = argparse.ArgumentParser(description="Create or renew a Microsoft Graph Inbox webhook subscription.")
    parser.add_argument("action", choices=["create", "renew", "status"])
    args = parser.parse_args()
    settings = get_settings()
    initialize_database(settings.database_path)
    if args.action == "create":
        subscription = create_subscription()
        print(f"Subscription created: {subscription['id']}\nExpires: {subscription['expirationDateTime']}")
    elif args.action == "renew":
        subscription = renew_subscription()
        print(f"Subscription renewed. Expires: {subscription['expirationDateTime']}")
    else:
        print(f"Subscription ID: {get_state(settings.database_path, 'graph_subscription_id') or 'not created'}")
        print(f"Expires: {get_state(settings.database_path, 'graph_subscription_expires_at') or 'not available'}")


if __name__ == "__main__":
    main()
