"""Local authenticated web application used by the desktop shell."""

from .api import AppServices, create_app, create_services

__all__ = ["AppServices", "create_app", "create_services"]
