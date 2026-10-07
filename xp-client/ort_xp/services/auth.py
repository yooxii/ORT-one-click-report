# -*- coding: utf-8 -*-
"""登录与角色（与主程序 ``AuthService`` 规则一致）。

散列算法必须与主程序完全一致，否则同一份 ``users`` 表两边不能互相认：

    PasswordHash = Base64( SHA256( UTF8( Salt + Password ) ) )
    Salt         = Guid.NewGuid().ToString("N")   → 32 位小写十六进制

校验规则（对应 ``AuthService.Login``）：
1. ``Username`` 精确匹配（大小写敏感）；
2. ``IsActive`` 必须为真；
3. 有密码 → 比对散列；没密码（Salt 与 PasswordHash 都为空）→ 用户名对即放行；
4. 角色取自 ``user_roles.Role``。
"""

import base64
import hashlib
import uuid


def new_salt():
    """生成与主程序 Guid("N") 等价的盐（32 位小写十六进制）。"""
    return uuid.uuid4().hex


def hash_password(salt, password):
    """Base64(SHA256(UTF8(salt + password)))。"""
    raw = ((salt or "") + (password or "")).encode("utf-8")
    return base64.b64encode(hashlib.sha256(raw).digest()).decode("ascii")


def has_password(salt, password_hash):
    return bool(salt) and bool(password_hash)


class User(object):
    """当前登录用户（只读视图）。"""

    def __init__(self, row):
        self.id = row["Id"]
        self.username = row["Username"]
        self.display_name = row["DisplayName"] or row["Username"]
        self.email = row["Email"] or ""
        self.is_active = bool(row["IsActive"])
        self.created_at = row["CreatedAt"]

    def __repr__(self):
        return "<User %s (%s)>" % (self.username, self.display_name)


class LoginResult(object):
    def __init__(self, ok, message="", passwordless=False):
        self.ok = ok
        self.message = message
        self.passwordless = passwordless


class AuthService(object):
    def __init__(self, database, credential_store=None, logger=None):
        self._db = database
        self._credentials = credential_store
        self._logger = logger
        self.current_user = None
        self.roles = []
        self.passwordless_login = False
        # 仅在内存中保留本次登录用的明文口令，供「记住登录」使用，不落盘
        self._last_password = ""

    # ---------------------------------------------------------------- 登录 / 登出

    def login(self, username, password):
        if not username:
            return LoginResult(False, "请输入用户名")
        row = self._db.query_one("SELECT * FROM users WHERE Username = ?", (username,))
        if row is None:
            return LoginResult(False, "用户名或密码不正确")
        if not bool(row["IsActive"]):
            return LoginResult(False, "该账号已停用")

        user = User(row)
        passwordless = False
        if has_password(row["Salt"], row["PasswordHash"]):
            expected = hash_password(row["Salt"], password or "")
            if expected != row["PasswordHash"]:
                return LoginResult(False, "用户名或密码不正确")
        else:
            # 主程序规则：没设过密码的账号，用户名对就放行（密码框留空）
            passwordless = True

        self.current_user = user
        self.passwordless_login = passwordless
        self._last_password = password or ""
        self.roles = self._load_roles(user.id)
        if self._logger is not None:
            self._logger.info(
                "用户登录：%s（角色：%s）%s"
                % (user.username, ",".join(self.roles) or "无", "，无密码账号直接放行" if passwordless else "")
            )
        return LoginResult(True, "登录成功", passwordless)

    def logout(self):
        if self.current_user is not None and self._logger is not None:
            self._logger.info("用户登出：%s" % self.current_user.username)
        self.current_user = None
        self.roles = []
        self.passwordless_login = False
        if self._credentials is not None:
            self._credentials.clear_login()

    def _load_roles(self, user_id):
        if not self._db.table_exists("user_roles"):
            return []
        rows = self._db.query("SELECT Role FROM user_roles WHERE UserId = ?", (user_id,))
        result = []
        for row in rows:
            role = row["Role"]
            if role:
                result.append(role)
        return result

    # ---------------------------------------------------------------- 便捷判断

    @property
    def is_authenticated(self):
        return self.current_user is not None

    def has_role(self, role):
        wanted = (role or "").lower()
        for item in self.roles:
            if item.lower() == wanted:
                return True
        return False

    def display_name(self):
        if self.current_user is None:
            return "游客"
        return self.current_user.display_name

    # ---------------------------------------------------------------- 记住登录

    def remember(self, days=7):
        """保存当前登录凭据到本机（DPAPI 加密）；需要先登录成功。"""
        if self._credentials is None or self.current_user is None:
            return False
        return self._credentials.save_login(self.current_user.username, self._last_password, days)

    def load_remembered(self):
        """返回 ``(username, password)`` 或 None（供启动时自动填充/自动登录）。"""
        if self._credentials is None:
            return None
        return self._credentials.load_login()

    def clear_remembered(self):
        if self._credentials is not None:
            self._credentials.clear_login()
