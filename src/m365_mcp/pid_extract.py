from __future__ import annotations

from m365_mcp.pid_policy import BlockedError


class LocalPidExtractor:
    """Local-only PID extraction is not implemented yet.

    When PID-safe mode is on, investor documents stay blocked until a
    local extractor can write approved fields to Excel without returning
    PID to the model.
    """

    def is_available(self) -> bool:
        return False

    def extract(self, *_args: object, **_kwargs: object) -> None:
        raise BlockedError("pid_extractor_unavailable")
