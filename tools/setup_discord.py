"""One-time Discord bot check: connect the bot once, then send you a test DM.

Discord requires a bot to have connected to its gateway at least once before
it can send messages. Run this script once on your PC after creating the bot:

    .venv\\Scripts\\python tools\\setup_discord.py

It asks for the bot token and your Discord user id (nothing is saved).
Requires the "websocket-client" package (installed by requirements.txt).
"""

from __future__ import annotations

import getpass
import json
import os
import sys

import websocket  # websocket-client

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from common.discord_dm import COLOR_GREEN, DiscordDM, DiscordError  # noqa: E402

GATEWAY = "wss://gateway.discord.gg/?v=10&encoding=json"


def identify_once(token: str) -> str:
    ws = websocket.create_connection(GATEWAY, timeout=20)
    try:
        hello = json.loads(ws.recv())
        if hello.get("op") != 10:
            raise RuntimeError(f"Unexpected gateway greeting: {hello}")
        ws.send(json.dumps({
            "op": 2,
            "d": {
                "token": token,
                "intents": 0,
                "properties": {"os": sys.platform, "browser": "candidatures-auto", "device": "candidatures-auto"},
            },
        }))
        while True:
            raw = ws.recv()
            # A bad token closes the connection (code 4004): recv() then returns "".
            event = json.loads(raw) if raw else None
            if event is None or event.get("op") == 9:
                raise RuntimeError("Discord a refusé le token du bot. Refais « Reset Token » et copie le nouveau.")
            if event.get("t") == "READY":
                user = event["d"]["user"]
                return f"{user['username']}"
    finally:
        ws.close()


def main() -> int:
    token = os.environ.get("DISCORD_BOT_TOKEN") or getpass.getpass("Token du bot Discord (masqué) : ")
    user_id = os.environ.get("DISCORD_USER_ID") or input("Ton identifiant d'utilisateur Discord : ")
    try:
        name = identify_once(token.strip())
        print(f"Bot connecté : {name}")
        DiscordDM(token, user_id).send(embed={
            "title": "✅ Le bot de candidatures est prêt",
            "description": "Tu recevras ici un message à chaque réponse d'entreprise et les rappels de relance.",
            "color": COLOR_GREEN,
        })
    except (OSError, websocket.WebSocketException) as exc:
        print(f"Connexion à Discord impossible : {exc}")
        print("Vérifie ta connexion internet (VPN, pare-feu), puis relance le script.")
        return 1
    except (DiscordError, RuntimeError) as exc:
        print(f"Échec : {exc}")
        print("Vérifie le token, ton identifiant, que le bot est sur un serveur où tu es, "
              "et que tu acceptes les messages privés des membres de ce serveur.")
        return 1
    print("Message de test envoyé : regarde tes messages privés Discord.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
