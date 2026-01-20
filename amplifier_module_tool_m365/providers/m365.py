"""Microsoft 365 / Microsoft Graph collaboration provider."""

import os
from dataclasses import dataclass

import httpx
from azure.identity import ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.groups.groups_request_builder import GroupsRequestBuilder
from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.email_address import EmailAddress
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.message import Message as OutlookMessage
from msgraph.generated.models.recipient import Recipient
from msgraph.generated.users.item.send_mail.send_mail_post_request_body import (
    SendMailPostRequestBody,
)
from msgraph.generated.users.users_request_builder import UsersRequestBuilder

from amplifier_module_tool_collab_core import (
    Channel,
    CollaborationProvider,
    Document,
    Message,
    Task,
    User,
)


@dataclass
class M365Config:
    """Microsoft 365 configuration."""

    tenant_id: str
    client_id: str
    client_secret: str
    webhooks: dict[str, str]  # channel_name -> webhook_url

    @classmethod
    def from_env(cls) -> "M365Config":
        """Load configuration from environment variables."""
        tenant_id = os.environ.get("M365_TENANT_ID")
        client_id = os.environ.get("M365_CLIENT_ID")
        client_secret = os.environ.get("M365_CLIENT_SECRET")

        if not tenant_id or not client_id or not client_secret:
            missing = [
                k
                for k, v in {
                    "M365_TENANT_ID": tenant_id,
                    "M365_CLIENT_ID": client_id,
                    "M365_CLIENT_SECRET": client_secret,
                }.items()
                if not v
            ]
            raise ValueError(
                f"Missing required environment variables: {', '.join(missing)}"
            )

        # Parse webhooks: "general=url1,alerts=url2,handoffs=url3"
        webhooks: dict[str, str] = {}
        webhook_str = os.environ.get("M365_TEAMS_WEBHOOKS", "")
        if webhook_str:
            for pair in webhook_str.split(","):
                if "=" in pair:
                    name, url = pair.split("=", 1)
                    webhooks[name.strip()] = url.strip()

        return cls(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret,
            webhooks=webhooks,
        )


class M365Provider(CollaborationProvider):
    """Microsoft 365 collaboration provider using Microsoft Graph API."""

    def __init__(self, config: dict | None = None):
        """Initialize the M365 provider.

        Args:
            config: Optional config dict. If not provided, loads from environment.
        """
        self._config = M365Config.from_env()
        self._credential = ClientSecretCredential(
            tenant_id=self._config.tenant_id,
            client_id=self._config.client_id,
            client_secret=self._config.client_secret,
        )
        self._client = GraphServiceClient(credentials=self._credential)

    @property
    def name(self) -> str:
        return "m365"

    # =========================================================================
    # Users & Directory
    # =========================================================================

    async def list_users(self, limit: int = 25) -> list[User]:
        """List users in the tenant."""
        query_params = UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
            top=limit,
            select=["id", "displayName", "userPrincipalName", "mail", "department"],
        )
        request_config = UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
            query_parameters=query_params,
        )
        result = await self._client.users.get(request_configuration=request_config)

        return [
            User(
                id=user.id or "",
                display_name=user.display_name or "",
                email=user.mail or user.user_principal_name,
                department=user.department,
            )
            for user in (result.value or [])
        ]

    async def get_user(self, user_id: str) -> User:
        """Get a specific user by ID or UPN."""
        user = await self._client.users.by_user_id(user_id).get()
        if not user:
            raise ValueError(f"User not found: {user_id}")
        return User(
            id=user.id or "",
            display_name=user.display_name or "",
            email=user.mail or user.user_principal_name,
            department=user.department,
        )

    # =========================================================================
    # Channels & Messaging
    # =========================================================================

    async def list_channels(self, team_id: str | None = None) -> list[Channel]:
        """List Teams channels."""
        if team_id:
            # List channels in a specific team
            result = await self._client.teams.by_team_id(team_id).channels.get()
            return [
                Channel(
                    id=ch.id or "",
                    name=ch.display_name or "",
                    description=ch.description,
                    team_id=team_id,
                )
                for ch in (result.value or [])
            ]
        else:
            # List all teams first, then aggregate channels
            query_params = GroupsRequestBuilder.GroupsRequestBuilderGetQueryParameters(
                filter="resourceProvisioningOptions/Any(x:x eq 'Team')",
                select=["id", "displayName"],
            )
            request_config = (
                GroupsRequestBuilder.GroupsRequestBuilderGetRequestConfiguration(
                    query_parameters=query_params,
                )
            )
            teams = await self._client.groups.get(request_configuration=request_config)

            all_channels: list[Channel] = []
            for team in (teams.value or [])[:5]:  # Limit to first 5 teams
                try:
                    if not team.id:
                        continue
                    channels = await self._client.teams.by_team_id(
                        team.id
                    ).channels.get()
                    for ch in channels.value or []:
                        all_channels.append(
                            Channel(
                                id=ch.id or "",
                                name=ch.display_name or "",
                                description=ch.description,
                                team_id=team.id,
                                team_name=team.display_name,
                            )
                        )
                except Exception:
                    continue  # Skip teams we can't access

            return all_channels

    async def get_messages(
        self,
        channel_id: str,
        limit: int = 20,
        team_id: str | None = None,
    ) -> list[Message]:
        """Get recent messages from a Teams channel."""
        if not team_id:
            raise ValueError("team_id is required for M365 channel messages")

        result = (
            await self._client.teams.by_team_id(team_id)
            .channels.by_channel_id(channel_id)
            .messages.get()
        )

        messages: list[Message] = []
        for msg in (result.value or [])[:limit]:
            sender = "Unknown"
            if msg.from_ and msg.from_.user:
                sender = msg.from_.user.display_name or "Unknown"

            messages.append(
                Message(
                    id=msg.id or "",
                    content=msg.body.content if msg.body else "",
                    sender=sender,
                    timestamp=str(msg.created_date_time)
                    if msg.created_date_time
                    else "",
                    channel_id=channel_id,
                )
            )

        return messages

    async def post_message(
        self,
        channel_name: str,
        message: str,
        title: str | None = None,
    ) -> bool:
        """Post a message to a Teams channel via webhook."""
        webhook_url = self._config.webhooks.get(channel_name)
        if not webhook_url:
            available = list(self._config.webhooks.keys())
            raise ValueError(
                f"No webhook configured for channel '{channel_name}'. "
                f"Available channels: {available}"
            )

        # Build payload
        if title:
            payload = {
                "@type": "MessageCard",
                "summary": title,
                "sections": [{"activityTitle": title, "text": message}],
            }
        else:
            payload = {"text": message}

        async with httpx.AsyncClient() as client:
            response = await client.post(webhook_url, json=payload)
            return response.status_code == 200

    # =========================================================================
    # Documents & Files (SharePoint)
    # =========================================================================

    async def list_documents(
        self,
        folder_path: str | None = None,
        site_id: str | None = None,
    ) -> list[Document]:
        """List documents in SharePoint."""
        if not site_id:
            # Get default site
            sites = await self._client.sites.get()
            if not sites or not sites.value:
                return []
            first_site = sites.value[0]
            if not first_site.id:
                return []
            site_id = first_site.id

        if folder_path and folder_path != "root":
            result = (
                await self._client.sites.by_site_id(site_id)
                .drive.root.item_with_path(folder_path)
                .children.get()
            )
        else:
            result = await self._client.sites.by_site_id(
                site_id
            ).drive.root.children.get()

        return [
            Document(
                id=item.id or "",
                name=item.name or "",
                path=folder_path or "/",
                web_url=item.web_url,
                size=item.size,
                is_folder=item.folder is not None,
            )
            for item in (result.value or [])
        ]

    async def upload_document(
        self,
        name: str,
        content: bytes | str,
        folder_path: str | None = None,
        site_id: str | None = None,
    ) -> Document:
        """Upload a document to SharePoint."""
        if not site_id:
            sites = await self._client.sites.get()
            if not sites or not sites.value:
                raise ValueError("No SharePoint sites available")
            first_site = sites.value[0]
            if not first_site.id:
                raise ValueError("SharePoint site has no ID")
            site_id = first_site.id

        if isinstance(content, str):
            content = content.encode("utf-8")

        path = (
            f"{folder_path}/{name}" if folder_path and folder_path != "root" else name
        )

        result = (
            await self._client.sites.by_site_id(site_id)
            .drive.root.item_with_path(path)
            .content.put(content)
        )

        return Document(
            id=result.id if result else "",
            name=name,
            path=path,
            web_url=result.web_url if result else None,
        )

    async def download_document(
        self,
        document_id: str,
        site_id: str | None = None,
    ) -> bytes:
        """Download a document from SharePoint."""
        if not site_id:
            sites = await self._client.sites.get()
            if not sites or not sites.value:
                raise ValueError("No SharePoint sites available")
            first_site = sites.value[0]
            if not first_site.id:
                raise ValueError("SharePoint site has no ID")
            site_id = first_site.id

        content = (
            await self._client.sites.by_site_id(site_id)
            .drive.items.by_drive_item_id(document_id)
            .content.get()
        )
        return content or b""

    # =========================================================================
    # Tasks & Planning (Planner)
    # =========================================================================

    async def list_tasks(self, plan_id: str | None = None) -> list[Task]:
        """List Planner tasks."""
        if not plan_id:
            # Would need to get plans first
            return []

        result = await self._client.planner.plans.by_planner_plan_id(
            plan_id
        ).tasks.get()

        return [
            Task(
                id=task.id or "",
                title=task.title or "",
                status="complete" if task.percent_complete == 100 else "in_progress",
                due_date=str(task.due_date_time) if task.due_date_time else None,
            )
            for task in (result.value or [])
        ]

    # =========================================================================
    # Email (Outlook)
    # =========================================================================

    async def send_email(
        self,
        to: list[str],
        subject: str,
        body: str,
        from_user: str | None = None,
    ) -> bool:
        """Send an email via Outlook."""
        if not from_user:
            # Use first available user (admin typically)
            users = await self.list_users(limit=1)
            if not users:
                raise ValueError("No users available to send email from")
            from_user = users[0].id

        message = OutlookMessage(
            subject=subject,
            body=ItemBody(
                content_type=BodyType.Text,
                content=body,
            ),
            to_recipients=[
                Recipient(email_address=EmailAddress(address=addr)) for addr in to
            ],
        )

        request_body = SendMailPostRequestBody(message=message, save_to_sent_items=True)
        await self._client.users.by_user_id(from_user).send_mail.post(request_body)
        return True
