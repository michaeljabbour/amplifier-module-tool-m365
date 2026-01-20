"""Amplifier Microsoft 365 collaboration module.

Provides M365 integration via Microsoft Graph API:
- Microsoft Teams (channels, messaging)
- SharePoint (documents, files)
- Outlook (email)
- Planner (tasks)
"""

from amplifier_module_tool_collab_core import register_provider
from .providers.m365 import M365Provider

__all__ = ["M365Provider", "mount"]

# Register provider with core module
register_provider("m365", M365Provider)


def mount(session):
    """Mount M365 collaboration to an Amplifier session.

    The provider is auto-registered on import. Tools from collab-core
    can use provider_name='m365' to access M365 functionality.
    """
    pass  # Provider registration happens on import
