"""Backup storage locations. Use only these names."""
from .local import LocalFolderStore
from . import nas

__all__ = ["LocalFolderStore", "nas"]
