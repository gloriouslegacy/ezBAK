"""Windows I/O boundary: external commands, elevation, file attributes, well-known paths, system info."""
from . import elevation, fileattrs, paths, proc, shell, sysinfo
from .sysinfo import free_space

__all__ = ["elevation", "fileattrs", "paths", "proc", "shell", "sysinfo", "free_space"]
