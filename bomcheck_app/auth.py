from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Set


PERMISSION_LABELS: Dict[str, str] = {
    "invalid_part": "编辑失效料号库",
    "binding": "编辑绑定料号",
    "important": "编辑重要物料",
    "blocked": "编辑屏蔽申请人",
    "asset": "维护料号资源",
}


def _hash_password(raw: str) -> str:
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


@dataclass
class UserAccount:
    username: str
    password_hash: str
    permissions: Set[str] = field(default_factory=set)
    is_admin: bool = False

    @classmethod
    def create(cls, username: str, password: str, *, is_admin: bool = False, permissions: Set[str] | None = None) -> "UserAccount":
        return cls(
            username=username,
            password_hash=_hash_password(password),
            permissions=set(permissions or set(PERMISSION_LABELS)) if is_admin else set(permissions or ()),
            is_admin=is_admin,
        )

    @classmethod
    def from_dict(cls, data: Dict) -> "UserAccount":
        return cls(
            username=data.get("username", ""),
            password_hash=data.get("password_hash", ""),
            permissions=set(data.get("permissions", [])),
            is_admin=bool(data.get("is_admin", False)),
        )

    def to_dict(self) -> Dict:
        return {
            "username": self.username,
            "password_hash": self.password_hash,
            "permissions": sorted(self.permissions),
            "is_admin": self.is_admin,
        }

    def set_password(self, password: str) -> None:
        self.password_hash = _hash_password(password)

    def verify(self, password: str) -> bool:
        return self.password_hash == _hash_password(password)

    def can(self, permission: str) -> bool:
        return self.is_admin or permission in self.permissions


class AccountStore:
    def __init__(self, path: Path) -> None:
        self.path = path
        self.accounts: Dict[str, UserAccount] = {}
        self._load()

    def _load(self) -> None:
        if not self.path.exists():
            self.accounts = {}
            self._ensure_default_admin()
            return
        try:
            raw = json.loads(self.path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            self.accounts = {}
            self._ensure_default_admin()
            return
        self.accounts = {}
        for item in raw:
            user = UserAccount.from_dict(item)
            if user.username:
                self.accounts[user.username] = user
        self._ensure_default_admin()

    def _ensure_default_admin(self) -> None:
        if not self.accounts:
            default_admin = UserAccount.create("admin", "admin", is_admin=True)
            self.accounts[default_admin.username] = default_admin
            self.save()

    def save(self) -> None:
        payload: List[Dict] = [user.to_dict() for user in self.accounts.values()]
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def authenticate(self, username: str, password: str) -> UserAccount | None:
        account = self.accounts.get(username)
        if account and account.verify(password):
            return account
        return None

    def upsert(self, account: UserAccount) -> None:
        if not account.username:
            raise ValueError("用户名不能为空")
        self.accounts[account.username] = account
        self.save()

    def delete(self, username: str) -> None:
        if username in self.accounts:
            del self.accounts[username]
            self.save()

    def list_users(self) -> List[UserAccount]:
        return sorted(self.accounts.values(), key=lambda item: item.username)
