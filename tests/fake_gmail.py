"""In-memory stand-in for GmailClient, used by the tests."""

from __future__ import annotations

from common.gmail import MessageInfo


class FakeGmail:
    ALL = "[Gmail]/Tous les messages"
    SENT = "[Gmail]/Messages envoy&AOk-s"

    def __init__(self, address: str, messages: list[tuple[MessageInfo, str]]):
        self.address = address
        self.messages = {m.uid: (m, body) for m, body in messages}
        self.selected = None
        self.raw_queries: list[str] = []

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        pass

    def ensure_labels(self, labels=()):
        pass

    def special_folder(self, flag):
        return {"\\All": self.ALL, "\\Sent": self.SENT}[flag]

    def select(self, folder, readonly=True):
        self.selected = folder

    def _in_folder(self):
        for info, _ in self.messages.values():
            if self.selected == self.ALL:
                yield info
            elif self.selected == self.SENT and info.from_addr == self.address:
                yield info
            elif self.selected in info.labels:
                yield info

    def search_all(self):
        return [m.uid for m in self._in_folder()]

    def search_raw(self, query):
        self.raw_queries.append(query)
        terms = query.split("{", 1)[1].rstrip("}").split() if "{" in query else []
        targets = [t.split(":", 1)[1] for t in terms]
        return [
            m.uid for m in self._in_folder()
            if m.from_addr != self.address and any(m.from_addr == t or m.from_addr.endswith("@" + t) for t in targets)
        ]

    def search_thread(self, thread_id):
        return [m.uid for m in self._in_folder() if m.thread_id == thread_id]

    def fetch_headers(self, uids):
        return [self.messages[u][0] for u in sorted(set(uids))]

    def fetch_text(self, uid):
        return self.messages[uid][1]

    def add_label(self, uids, label):
        for uid in uids:
            self.messages[uid][0].labels.add(label)

    def remove_label(self, uids, label):
        for uid in uids:
            self.messages[uid][0].labels.discard(label)
