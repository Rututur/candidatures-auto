"""Send direct messages with a Discord bot through the REST API (stdlib only)."""

from __future__ import annotations

import json
import time
import urllib.error
import urllib.request

API = "https://discord.com/api/v10"
USER_AGENT = "DiscordBot (https://github.com/candidatures-auto, 1.0)"

COLOR_RED = 0xE5484D
COLOR_GREEN = 0x30A46C
COLOR_GREY = 0x8B8D98
COLOR_BLUE = 0x3E63DD
COLOR_ORANGE = 0xF76B15


class DiscordError(Exception):
    pass


class DiscordDM:
    def __init__(self, bot_token: str, user_id: str):
        self._token = bot_token.strip()
        self._user_id = str(user_id).strip()
        self._channel_id: str | None = None

    def _request(self, method: str, path: str, payload: dict | None = None, retries: int = 3) -> dict:
        body = json.dumps(payload).encode() if payload is not None else None
        request = urllib.request.Request(
            API + path,
            data=body,
            method=method,
            headers={
                "Authorization": f"Bot {self._token}",
                "User-Agent": USER_AGENT,
                "Content-Type": "application/json",
            },
        )
        for attempt in range(retries):
            try:
                with urllib.request.urlopen(request, timeout=30) as response:
                    raw = response.read()
                    return json.loads(raw) if raw else {}
            except urllib.error.HTTPError as exc:
                detail = exc.read().decode(errors="replace")
                if exc.code == 429 and attempt < retries - 1:
                    try:
                        wait = float(json.loads(detail).get("retry_after", 2))
                    except ValueError:
                        wait = 2.0
                    time.sleep(min(wait, 30))
                    continue
                raise DiscordError(f"Discord API {method} {path} -> HTTP {exc.code}: {detail}") from exc
            except urllib.error.URLError as exc:
                if attempt < retries - 1:
                    time.sleep(2)
                    continue
                raise DiscordError(f"Discord API unreachable: {exc}") from exc
        raise DiscordError("Discord API: too many retries")

    def _channel(self) -> str:
        if self._channel_id is None:
            channel = self._request("POST", "/users/@me/channels", {"recipient_id": self._user_id})
            self._channel_id = channel["id"]
        return self._channel_id

    def send(self, content: str = "", embed: dict | None = None) -> None:
        payload: dict = {"content": content, "allowed_mentions": {"parse": []}}
        if embed:
            payload["embeds"] = [embed]
        self._request("POST", f"/channels/{self._channel()}/messages", payload)
