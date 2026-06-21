#!/usr/bin/env python3
"""Manage official-docx-formatter profiles."""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any, Dict

SKILL_DIR = Path(__file__).resolve().parents[1]
PROFILES_DIR = SKILL_DIR / "profiles"
BASE_PROFILE = "standard-party-government"


def slugify(value: str) -> str:
    value = value.strip().lower()
    value = re.sub(r"[^a-z0-9\-]+", "-", value)
    value = re.sub(r"-+", "-", value).strip("-")
    if not value:
        raise ValueError("profile id must contain at least one letter or digit")
    return value


def load_json(path: Path) -> Dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def profile_path(profile_id: str) -> Path:
    return PROFILES_DIR / f"{profile_id}.json"


def list_profiles() -> None:
    for path in sorted(PROFILES_DIR.glob("*.json")):
        data = load_json(path)
        print(f"{path.stem}\t{data.get('display_name', path.stem)}")


def show_profile(profile_id: str) -> None:
    path = profile_path(profile_id)
    if not path.exists():
        raise SystemExit(f"profile not found: {profile_id}")
    print(json.dumps(load_json(path), ensure_ascii=False, indent=2))


def set_font(data: Dict[str, Any], key: str, font: str | None) -> None:
    if not font:
        return
    data.setdefault("fonts", {}).setdefault(key, {})["fallbacks"] = [font]


def create_profile(args: argparse.Namespace) -> None:
    profile_id = slugify(args.profile_id)
    path = profile_path(profile_id)
    if path.exists() and not args.force:
        raise SystemExit(f"profile already exists: {profile_id}; pass --force to overwrite")

    data: Dict[str, Any] = {
        "profile_id": profile_id,
        "display_name": args.display_name or args.profile_id,
        "description": args.description or "自定义单位公文格式配置。未指定字段继承党政机关公文标准配置。",
        "inherits": BASE_PROFILE,
    }

    set_font(data, "title", args.title_font)
    set_font(data, "body", args.body_font)
    set_font(data, "level1", args.level1_font)
    set_font(data, "level2", args.level2_font)

    layout: Dict[str, Any] = {}
    if args.line_spacing_pt is not None:
        layout["line_spacing_pt"] = args.line_spacing_pt
    if layout:
        data["layout"] = layout

    page: Dict[str, Any] = {}
    for field in ["top_margin_cm", "bottom_margin_cm", "left_margin_cm", "right_margin_cm"]:
        value = getattr(args, field)
        if value is not None:
            page[field] = value
    if page:
        data["page"] = page

    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"created profile: {path}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Manage official DOCX formatting profiles")
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("list", help="List profiles")

    show = sub.add_parser("show", help="Show a profile JSON")
    show.add_argument("profile_id")

    create = sub.add_parser("create", help="Create a custom organization profile")
    create.add_argument("profile_id")
    create.add_argument("--display-name")
    create.add_argument("--description")
    create.add_argument("--title-font")
    create.add_argument("--body-font")
    create.add_argument("--level1-font")
    create.add_argument("--level2-font")
    create.add_argument("--line-spacing-pt", type=float)
    create.add_argument("--top-margin-cm", type=float)
    create.add_argument("--bottom-margin-cm", type=float)
    create.add_argument("--left-margin-cm", type=float)
    create.add_argument("--right-margin-cm", type=float)
    create.add_argument("--force", action="store_true")

    args = parser.parse_args()
    if args.command == "list":
        list_profiles()
    elif args.command == "show":
        show_profile(args.profile_id)
    elif args.command == "create":
        create_profile(args)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
