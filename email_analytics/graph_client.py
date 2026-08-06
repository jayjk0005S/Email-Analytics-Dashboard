from __future__ import annotations

from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

import msal
import requests
from msal_extensions import FilePersistenceWithDataProtection, PersistedTokenCache


GRAPH_SCOPE = ["Mail.Read"]
GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
IMMUTABLE_ID_HEADER = 'IdType="ImmutableId"'


class GraphConfigurationError(RuntimeError):
    pass


class GraphMailClient:
    """Microsoft Graph client using an interactive device-code sign-in.

    This is intended for a local, single-user dashboard. It accesses only the
    mailbox belonging to the person who signs in, using delegated Mail.Read.
    """

    def __init__(
        self, client_id: str | None, tenant_id: str, cache_path: Path, store_body_preview: bool = False
    ):
        if not client_id:
            raise GraphConfigurationError(
                "GRAPH_CLIENT_ID is missing. Create a Microsoft Entra app registration, "
                "then copy .env.example to .env and add its Application (client) ID."
            )
        self.cache_path = cache_path
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        # Windows DPAPI ties this encrypted cache to the current Windows user.
        # A copied cache file cannot be decrypted by a different Windows user.
        self.cache = PersistedTokenCache(FilePersistenceWithDataProtection(str(cache_path)))
        self.store_body_preview = store_body_preview
        self.app = msal.PublicClientApplication(
            client_id=client_id,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
            token_cache=self.cache,
        )

    def access_token(self) -> str:
        accounts = self.app.get_accounts()
        result = self.app.acquire_token_silent(GRAPH_SCOPE, account=accounts[0]) if accounts else None
        if not result:
            flow = self.app.initiate_device_flow(scopes=GRAPH_SCOPE)
            if "user_code" not in flow:
                raise RuntimeError(f"Could not start Microsoft sign-in: {flow}")
            print(flow["message"])
            result = self.app.acquire_token_by_device_flow(flow)
        if "access_token" not in result:
            detail = result.get("error_description", result)
            raise RuntimeError(f"Microsoft sign-in failed: {detail}")
        return result["access_token"]

    def _selected_message_fields(self) -> str:
        fields = "id,from,subject,receivedDateTime,hasAttachments,importance,conversationId,webLink"
        return f"{fields},bodyPreview" if self.store_body_preview else fields

    def _headers(self) -> dict[str, str]:
        """Use immutable IDs so a folder move doesn't create a second local row."""
        return {
            "Authorization": f"Bearer {self.access_token()}",
            "Prefer": IMMUTABLE_ID_HEADER,
        }

    def get_messages_since(self, since: datetime) -> list[dict[str, Any]]:
        headers = self._headers()
        selected_fields = self._selected_message_fields()
        params: dict[str, str] | None = {
            "$select": selected_fields,
            "$filter": f"receivedDateTime ge {since.strftime('%Y-%m-%dT%H:%M:%SZ')}",
            "$top": "100",
        }
        url = f"{GRAPH_ROOT}/me/mailFolders/inbox/messages"
        messages: list[dict[str, Any]] = []

        while url:
            response = requests.get(url, headers=headers, params=params, timeout=30)
            response.raise_for_status()
            payload = response.json()
            messages.extend(payload.get("value", []))
            url = payload.get("@odata.nextLink")
            params = None
        return messages

    def get_message(self, message_id: str) -> dict[str, Any]:
        """Read one message after a Graph change notification."""
        selected_fields = self._selected_message_fields()
        response = requests.get(
            f"{GRAPH_ROOT}/me/messages/{message_id}",
            headers=self._headers(),
            params={"$select": selected_fields},
            timeout=30,
        )
        response.raise_for_status()
        return response.json()

    def create_inbox_subscription(
        self, notification_url: str, lifecycle_url: str, client_state: str
    ) -> dict[str, Any]:
        """Subscribe to creation of Inbox messages for the signed-in user."""
        me = requests.get(f"{GRAPH_ROOT}/me", headers=self._headers(), timeout=30)
        me.raise_for_status()
        user_id = me.json()["id"]
        payload = {
            "changeType": "created",
            "notificationUrl": notification_url,
            "lifecycleNotificationUrl": lifecycle_url,
            "resource": f"users/{user_id}/mailFolders('inbox')/messages",
            # Outlook message subscriptions without resource data support up to seven days.
            "expirationDateTime": (datetime.now(timezone.utc) + timedelta(days=6)).isoformat(),
            "clientState": client_state,
        }
        response = requests.post(
            f"{GRAPH_ROOT}/subscriptions",
            headers={**self._headers(), "Content-Type": "application/json"},
            json=payload,
            timeout=30,
        )
        response.raise_for_status()
        return response.json()

    def renew_subscription(self, subscription_id: str) -> dict[str, Any]:
        payload = {"expirationDateTime": (datetime.now(timezone.utc) + timedelta(days=6)).isoformat()}
        response = requests.patch(
            f"{GRAPH_ROOT}/subscriptions/{subscription_id}",
            headers={**self._headers(), "Content-Type": "application/json"},
            json=payload,
            timeout=30,
        )
        response.raise_for_status()
        return response.json()
