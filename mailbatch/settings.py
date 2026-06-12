# -*- coding: utf-8 -*-
"""설정 저장 (%APPDATA%\\MailBatch\\settings.json).

비밀번호는 '비밀번호 저장'을 켠 경우에만 Windows DPAPI로 암호화해 저장한다.
DPAPI는 같은 Windows 사용자 계정에서만 복호화되므로 파일이 유출돼도 안전하다.
"""
import base64
import ctypes
import ctypes.wintypes
import json
import os


def settings_path() -> str:
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    d = os.path.join(base, "MailBatch")
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, "settings.json")


class _DATA_BLOB(ctypes.Structure):
    _fields_ = [("cbData", ctypes.wintypes.DWORD),
                ("pbData", ctypes.POINTER(ctypes.c_char))]


def _blob_to_bytes(blob: _DATA_BLOB) -> bytes:
    data = ctypes.string_at(blob.pbData, blob.cbData)
    ctypes.windll.kernel32.LocalFree(blob.pbData)
    return data


def _crypt(data: bytes, protect: bool) -> bytes:
    blob_in = _DATA_BLOB(len(data), ctypes.cast(ctypes.create_string_buffer(data, len(data)),
                                                ctypes.POINTER(ctypes.c_char)))
    blob_out = _DATA_BLOB()
    fn = (ctypes.windll.crypt32.CryptProtectData if protect
          else ctypes.windll.crypt32.CryptUnprotectData)
    if not fn(ctypes.byref(blob_in), None, None, None, None, 0, ctypes.byref(blob_out)):
        raise OSError("DPAPI 호출 실패")
    return _blob_to_bytes(blob_out)


def dpapi_protect(data: bytes) -> str:
    return base64.b64encode(_crypt(data, True)).decode("ascii")


def dpapi_unprotect(b64: str) -> bytes:
    return _crypt(base64.b64decode(b64), False)


def load_settings() -> dict:
    try:
        with open(settings_path(), "r", encoding="utf-8") as fp:
            d = json.load(fp)
    except (OSError, ValueError):
        return {}
    enc = d.pop("password_enc", None)
    if enc:
        try:
            d["password"] = dpapi_unprotect(enc).decode("utf-8")
        except Exception:
            d["password"] = ""
    return d


def save_settings(d: dict):
    d = dict(d)
    password = d.pop("password", "")
    if d.pop("save_password", False) and password:
        d["password_enc"] = dpapi_protect(password.encode("utf-8"))
        d["save_password"] = True
    with open(settings_path(), "w", encoding="utf-8") as fp:
        json.dump(d, fp, ensure_ascii=False, indent=2)
