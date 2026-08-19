"""Compiling ``vbaProject.bin`` from VBA source, with no Excel involved."""

from __future__ import annotations

from .compression import compress
from .project import VbaBuildError
from .project import build as build_project

__all__ = ["VbaBuildError", "build_project", "compress"]
