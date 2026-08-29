#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
协作中心 - 数据处理脚本
========================
处理两类由网页「协作中心」生成、通过 GitHub 上传的数据：

1. data/registrations/*.json    注册申请  -> 自动写入 users.json（注册即成为会员）
2. data/decisions/*.json        审核决定  -> 通过：应用文件到 source/ 并重建网页
                                            拒绝：仅写通知
                                            任命/撤销管理员：更新 users.json 角色

所有处理结果写入 notifications.json（申请人可见）与 history.json（审核留痕），
处理完毕的申请归档到 data/archive/ 避免重复处理。

用法（在仓库根目录执行）：
  python3 tools/collab_process.py --all        # 处理全部待办
  python3 tools/collab_process.py --registrations
  python3 tools/collab_process.py --decisions
"""
import argparse
import json
import os
import re
import shutil
import sys
import time
from datetime import datetime, timezone, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
USERS = DATA / "users.json"
NOTIF = DATA / "notifications.json"
HISTORY = DATA / "history.json"
REG_DIR = DATA / "registrations"
INBOX_DIR = DATA / "inbox"
DEC_DIR = DATA / "decisions"
ARC_DIR = DATA / "archive"
ARC_REG = ARC_DIR / "registrations"
ARC_INBOX = ARC_DIR / "inbox"
ARC_DEC = ARC_DIR / "decisions"

TZ = timezone(timedelta(hours=8))


def now():
    return datetime.now(TZ).strftime("%Y-%m-%dT%H:%M:%S+08:00")


def load_json(path, default=None):
    if not path.exists():
        return default if default is not None else {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return default if default is not None else {}


def save_json(path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")


def notify(nick, ntype, title, body):
    data = load_json(NOTIF, {"list": []})
    data.setdefault("list", [])
    data["list"].insert(0, {
        "id": "n_" + str(int(time.time() * 1000)),
        "to": nick,
        "type": ntype,
        "title": title,
        "body": body,
        "time": now(),
        "read": False,
    })
    save_json(NOTIF, data)


def add_history(rec):
    data = load_json(HISTORY, {"list": []})
    data.setdefault("list", [])
    data["list"].insert(0, dict(rec, time=now()))
    save_json(HISTORY, data)


def move_to_archive(src: Path, dest_dir: Path):
    dest_dir.mkdir(parents=True, exist_ok=True)
    if src.is_dir():
        target = dest_dir / src.name
        if target.exists():
            shutil.rmtree(target)
        shutil.move(str(src), str(target))
    else:
        target = dest_dir / src.name
        if target.exists():
            target.unlink()
        shutil.move(str(src), str(target))


# ---------------------------------------------------------------- 注册
def process_registrations():
    if not REG_DIR.exists():
        return 0
    n = 0
    users = load_json(USERS, {})
    for f in sorted(REG_DIR.glob("reg_*.json")):
        try:
            reg = json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            continue
        nick = str(reg.get("nick", "")).strip()
        pass_hash = str(reg.get("pass_hash", "")).strip()
        if not nick or not pass_hash:
            continue
        existed = nick in users
        users[nick] = {
            "nick": nick,
            "pass_hash": pass_hash,
            "role": users.get(nick, {}).get("role", "member"),
            "created": users.get(nick, {}).get("created", reg.get("time", now())),
        }
        if existed:
            notify(nick, "system", "注册信息已更新", "你的账号信息已更新，登录信息保持不变。")
        else:
            notify(nick, "register_ok", "注册成功", "欢迎加入！现在可以登录并上传文件，提交后将由管理员审核。")
        add_history({
            "kind": "register",
            "applicant": nick,
            "action": "update" if existed else "create",
            "note": "网页注册申请（自动生效）",
        })
        n += 1
        move_to_archive(f, ARC_REG)
    if n:
        save_json(USERS, users)
    return n


# ---------------------------------------------------------------- 审核决定
def process_decisions():
    if not DEC_DIR.exists():
        return 0
    n = 0
    for f in sorted(DEC_DIR.glob("*.json")):
        try:
            dec = json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            continue
        action = str(dec.get("action", "")).strip()
        if action not in ("approve", "reject"):
            continue
        kind = str(dec.get("type", "upload")).strip()
        reviewer = str(dec.get("reviewer", "")).strip() or "管理员"
        note = str(dec.get("note", "")).strip()
        n += 1
        if kind == "upload":
            _handle_upload_decision(dec, action, reviewer, note)
        elif kind == "promote":
            _handle_promote_decision(dec, action, reviewer, note)
        elif kind == "demote":
            _handle_promote_decision(dec, action, reviewer, note)
        move_to_archive(f, ARC_DEC)
    return n


def _handle_upload_decision(dec, action, reviewer, note):
    aid = str(dec.get("id", "")).strip()
    applicant = str(dec.get("applicant", "")).strip()
    meta = None
    meta_path = INBOX_DIR / f"{aid}.json"
    if meta_path.exists():
        try:
            meta = json.loads(meta_path.read_text(encoding="utf-8"))
        except Exception:
            meta = None

    if action == "approve":
        applied = 0
        applied_files = []
        if meta:
            files = meta.get("files", [])
            position = str(meta.get("position", "other")).strip()
            for fn in files:
                fn = str(fn).strip()
                src = INBOX_DIR / fn
                if not src.exists():
                    continue
                dst = _position_target(src, position)
                dst.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(str(src), str(dst))
                applied_files.append(dst.name)
                applied += 1
        if applied:
            # 触发网站重建
            _rebuild_site()
            notify(applicant, "approved", "审核通过", "你提交的更新已通过审核并应用到网页。")
        else:
            notify(applicant, "approved", "审核通过", "审核已通过，但未在申请中找到可应用的文件，请联系管理员。")
        add_history({
            "kind": "upload", "applicant": applicant, "reviewer": reviewer,
            "action": "approve", "id": aid, "note": note,
            "detail": "应用文件: " + (", ".join(applied_files) if applied_files else "无"),
        })
    else:
        reason = note or "未通过审核"
        notify(applicant, "rejected", "审核被拒绝", reason)
        add_history({
            "kind": "upload", "applicant": applicant, "reviewer": reviewer,
            "action": "reject", "id": aid, "note": note,
        })
    # 归档申请本身
    if meta_path.exists():
        move_to_archive(meta_path, ARC_INBOX)
    if meta:
        for fn in meta.get("files", []):
            p = INBOX_DIR / str(fn).strip()
            if p.exists():
                move_to_archive(p, ARC_INBOX)


def _position_target(src: Path, position):
    """申请中选择的改动位置 -> source/ 下规范目标文件（update_site.py 只识别根目录文件）"""
    ext = src.suffix.lower()
    if position == "standard" and ext == ".pdf":
        return ROOT / "source" / "standard.pdf"
    if position == "sellpoint" and ext in (".xlsx", ".xls"):
        return ROOT / "source" / "sellpoints.xlsx"
    if position == "gallery" and ext == ".zip":
        return ROOT / "source" / "gallery.zip"
    return ROOT / "source" / "misc" / src.name


def _position_guess(src: Path):
    """直传文件自动归类：按扩展名推断更新位置"""
    ext = src.suffix.lower()
    if ext == ".pdf":
        return "standard"
    if ext in (".xlsx", ".xls"):
        return "sellpoint"
    if ext == ".zip":
        return "gallery"
    return "other"


def process_direct_uploads():
    """处理直传文件：会员直接在 GitHub 上传页把文件拖入 data/inbox（无 meta 申请单），
    按扩展名自动归类到 source/ 并重建网页，写通知给超级管理员与所有管理员。"""
    if not INBOX_DIR.exists():
        return 0
    # 已被 meta 申请引用的文件名，跳过（避免与审核流程重复处理）
    claimed = set()
    for f in INBOX_DIR.glob("*.json"):
        try:
            meta = json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            continue
        for fn in meta.get("files", []):
            claimed.add(str(fn).strip())

    n = 0
    for f in sorted(INBOX_DIR.iterdir()):
        if not f.is_file() or f.suffix.lower() == ".json":
            continue
        if f.name in claimed:
            continue
        position = _position_guess(f)
        dst = _position_target(f, position)
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(str(f), str(dst))
        # 写历史与通知（无申请人信息则通知管理员）
        add_history({
            "kind": "direct_upload", "applicant": "(直传)", "reviewer": "系统",
            "action": "apply", "file": f.name, "position": position,
            "detail": f"直传文件 {f.name} 自动归类为 {position}，已应用并重建网页",
        })
        print(f"[direct] {f.name} -> {dst.name} ({position})")
        n += 1
        move_to_archive(f, ARC_INBOX)
    if n:
        _rebuild_site()
        for adm in _admins():
            notify(adm, "system", "网页已自动更新",
                   f"收到 {n} 个直传文件，已自动归类并更新网页。可前往查看最新内容。")
    return n


def _admins():
    users = load_json(USERS, {})
    out = []
    for nick, info in users.items():
        if isinstance(info, dict) and info.get("role") == "admin":
            out.append(nick)
    if not out:
        out = ["yingchenDong"]
    return out


def _rebuild_site():
    """运行 update_site.py 重建 index.html（source/ 有新文件时）"""
    script = ROOT / "tools" / "update_site.py"
    try:
        import subprocess
        r = subprocess.run(
            [sys.executable, str(script), "--all", "--source", str(ROOT / "source")],
            cwd=str(ROOT), capture_output=True, text=True, timeout=300,
        )
        print("[rebuild]", r.returncode, (r.stdout or "")[-800:], (r.stderr or "")[-400:])
    except Exception as e:
        print("[rebuild] failed:", e)


def _handle_promote_decision(dec, action, reviewer, note):
    nick = str(dec.get("applicant", "")).strip()
    users = load_json(USERS, {})
    if nick not in users:
        notify(nick, "system", "操作未生效", f"昵称 {nick} 尚未注册，无法调整角色。")
        add_history({"kind": "role", "applicant": nick, "reviewer": reviewer,
                     "action": action, "note": note, "detail": "用户不存在"})
        save_json(USERS, users)
        return
    if action == "approve":
        role = "admin" if dec.get("type") == "promote" else "member"
        users[nick]["role"] = role
        save_json(USERS, users)
        if role == "admin":
            notify(nick, "role", "已任命为管理员", f"你已被 {reviewer} 指定为管理员，现在可以审核其他会员的上传申请。")
        else:
            notify(nick, "role", "管理员权限已撤销", f"你已被 {reviewer} 撤销管理员权限，但仍可上传文件。")
        add_history({"kind": "role", "applicant": nick, "reviewer": reviewer,
                     "action": "promote" if role == "admin" else "demote", "note": note})
    else:
        notify(nick, "role", "任命未通过", f"管理员任命未通过：{note or '未说明原因'}")
        add_history({"kind": "role", "applicant": nick, "reviewer": reviewer,
                     "action": "reject", "note": note})


# ---------------------------------------------------------------- 入口
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--registrations", action="store_true")
    ap.add_argument("--decisions", action="store_true")
    args = ap.parse_args()

    total = 0
    if args.all or args.registrations:
        total += process_registrations()
    if args.all or args.decisions:
        total += process_decisions()
    if args.all:
        total += process_direct_uploads()
    print(f"collab: processed {total} item(s)")


if __name__ == "__main__":
    main()
