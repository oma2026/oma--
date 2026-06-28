# -*- coding: utf-8 -*-
"""
二手車進口價格報價系統 - Streamlit 多人版 v1.11
功能：多人登入、業務／老闆權限分流、CIF/FOB 報價計算、CC數自動關稅、奢侈稅、服務費、特殊車測、不限項數其他加裝、報價紀錄、CSV/Excel 匯出。
"""
from __future__ import annotations

import json
import sqlite3
import uuid
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

import pandas as pd
import streamlit as st

APP_DIR = Path(__file__).resolve().parent
CONFIG_DIR = APP_DIR / "config"
DATA_DIR = APP_DIR / "data"
EXPORT_DIR = APP_DIR / "exports"
CONFIG_PATH = CONFIG_DIR / "settings.json"
DB_PATH = DATA_DIR / "quotes.db"
APP_VERSION = "v1.11 業務隱藏公式輔助版"

DEFAULT_SETTINGS: Dict[str, Any] = {
    "app_settings": {
        "version": APP_VERSION,
        "multi_user_enabled": True,
    },
    "users": [
        {"username": "darren", "display_name": "老闆", "role": "admin", "password": "oma1688", "enabled": True},
        {"username": "peter", "display_name": "Peter", "role": "sales", "password": "1234", "enabled": True},
        {"username": "cbc", "display_name": "CBC", "role": "sales", "password": "1234", "enabled": True},
        {"username": "lai", "display_name": "Lai", "role": "sales", "password": "1234", "enabled": True},
        {"username": "gary", "display_name": "Gary", "role": "sales", "password": "1234", "enabled": True},
    ],
    "quote_settings": {
        "basic_vehicle_test_fee": 80000,
        "cleaning_fee": 30000,
        "customs_trucking_fee": 30000,
        "sales_admin_fixed_profit_fee": 80000,
        "other_fee_default": 0,
        "normal_multiplier": 1.07,
        "cash_multiplier": 1.05,
        # v1.10：依 CC 數自動計算關稅與奢侈稅。2000cc 以下預設 0.545。
        "customs_rate_under_2000": 0.545,
        "customs_rate_over_2000": 0.61,
        "customs_add_amount_foreign": 1500,
        "luxury_tax_threshold": 3000000,
        "luxury_tax_rate": 0.10,
    },
    "currencies": ["USD", "EUR", "JPY", "KRW", "CAD", "TWD"],
    "source_countries": ["美國", "德國", "日本", "韓國", "加拿大"],
    "shipping_settings": [
        {"source_country": "美國", "default_currency": "USD", "default_fee": 0, "enabled": True},
        {"source_country": "德國", "default_currency": "EUR", "default_fee": 0, "enabled": True},
        {"source_country": "日本", "default_currency": "JPY", "default_fee": 0, "enabled": True},
        {"source_country": "韓國", "default_currency": "USD", "default_fee": 0, "enabled": True},
        {"source_country": "加拿大", "default_currency": "CAD", "default_fee": 0, "enabled": True},
    ],
}

DB_COLUMNS: Dict[str, str] = {
    "id": "TEXT PRIMARY KEY",
    "created_at": "TEXT",
    "quote_no": "TEXT",
    "created_by_username": "TEXT",
    "created_by_role": "TEXT",
    "vehicle_name": "TEXT",
    "model_year": "TEXT",
    "engine_cc_category": "TEXT",
    "vin": "TEXT",
    "customer_name": "TEXT",
    "salesperson": "TEXT",
    "source_country": "TEXT",
    "price_term": "TEXT",
    "car_currency": "TEXT",
    "foreign_amount": "REAL",
    "car_exchange_rate": "REAL",
    "shipping_currency": "TEXT",
    "shipping_fee": "REAL",
    "shipping_exchange_rate": "REAL",
    "foreign_car_cost": "REAL",
    "customs_amount": "REAL",
    "customs_rate": "REAL",
    "customs_add_amount_foreign": "REAL",
    "customs_formula": "TEXT",
    "dutiable_price": "REAL",
    "luxury_tax_threshold": "REAL",
    "luxury_tax_rate": "REAL",
    "luxury_tax_amount": "REAL",
    "luxury_tax_formula": "TEXT",
    "customs_note": "TEXT",
    "basic_vehicle_test_fee": "REAL",
    "is_special_vehicle": "INTEGER",
    "special_vehicle_test_surcharge": "REAL",
    "vehicle_test_fee": "REAL",
    "cleaning_fee": "REAL",
    "customs_trucking_fee": "REAL",
    "sales_admin_fixed_profit_fee": "REAL",
    "service_fee_currency": "TEXT",
    "service_fee_amount": "REAL",
    "service_fee_exchange_rate": "REAL",
    "service_fee_twd": "REAL",
    "service_fee_note": "TEXT",
    "other_fee": "REAL",
    "other_fee_note": "TEXT",
    "addon_items_json": "TEXT",
    "addon_total": "REAL",
    "addon_1_name": "TEXT",
    "addon_1_amount": "REAL",
    "addon_2_name": "TEXT",
    "addon_2_amount": "REAL",
    "addon_3_name": "TEXT",
    "addon_3_amount": "REAL",
    "addon_4_name": "TEXT",
    "addon_4_amount": "REAL",
    "addon_5_name": "TEXT",
    "addon_5_amount": "REAL",
    "misc_total": "REAL",
    "total_cost": "REAL",
    "transaction_type": "TEXT",
    "multiplier": "REAL",
    "normal_multiplier": "REAL",
    "cash_multiplier": "REAL",
    "company_base_price": "REAL",
    "cash_price": "REAL",
    "cash_price_difference": "REAL",
    "selected_price": "REAL",
    "base_price": "REAL",
    "suggested_price": "REAL",
    "gross_profit": "REAL",
    "cash_transaction": "INTEGER",
    "notes": "TEXT",
    "payload_json": "TEXT",
}
RECORD_COLUMNS = list(DB_COLUMNS.keys())


def deep_merge(default: Dict[str, Any], loaded: Dict[str, Any]) -> Dict[str, Any]:
    """保留舊設定，同時補上新版本預設欄位。"""
    result = dict(default)
    for key, value in loaded.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def normalize_settings(settings: Dict[str, Any]) -> Dict[str, Any]:
    previous_version = str((settings or {}).get("app_settings", {}).get("version", ""))
    settings = deep_merge(DEFAULT_SETTINGS, settings or {})
    # v1.10 修正：若沿用 v1.9 設定檔且 2000cc 以下仍是舊預設 0.454，自動改成 0.545。
    try:
        if previous_version.startswith("v1.9") and abs(float(settings.get("quote_settings", {}).get("customs_rate_under_2000", 0)) - 0.454) < 1e-9:
            settings.setdefault("quote_settings", {})["customs_rate_under_2000"] = 0.545
    except Exception:
        pass
    settings.setdefault("app_settings", {})["version"] = APP_VERSION

    users = settings.get("users") or DEFAULT_SETTINGS["users"]
    normalized_users: List[Dict[str, Any]] = []
    seen: set[str] = set()
    for user in users:
        username = str(user.get("username", "")).strip().lower()
        if not username or username in seen:
            continue
        seen.add(username)
        role = str(user.get("role", "sales")).strip().lower()
        if role not in {"admin", "sales"}:
            role = "sales"
        normalized_users.append(
            {
                "username": username,
                "display_name": str(user.get("display_name") or username).strip(),
                "role": role,
                "password": str(user.get("password") or "1234"),
                "enabled": bool(user.get("enabled", True)),
            }
        )
    if not any(u["role"] == "admin" and u["enabled"] for u in normalized_users):
        normalized_users.insert(0, DEFAULT_SETTINGS["users"][0])
    settings["users"] = normalized_users

    # 舊版只有 admin_password 時，保留給第一個 admin 參考。
    legacy_admin_password = settings.get("app_settings", {}).get("admin_password")
    if legacy_admin_password:
        for user in settings["users"]:
            if user["role"] == "admin" and not user.get("password"):
                user["password"] = str(legacy_admin_password)
                break
    first_admin = next((u for u in settings["users"] if u["role"] == "admin"), DEFAULT_SETTINGS["users"][0])
    settings["app_settings"]["admin_password"] = first_admin.get("password", "oma1688")
    return settings


def save_settings(settings: Dict[str, Any]) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    normalized = normalize_settings(settings)
    CONFIG_PATH.write_text(json.dumps(normalized, ensure_ascii=False, indent=2), encoding="utf-8")


def load_settings() -> Dict[str, Any]:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    if not CONFIG_PATH.exists():
        save_settings(DEFAULT_SETTINGS)
        return normalize_settings(DEFAULT_SETTINGS)
    try:
        loaded = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        settings = normalize_settings(loaded)
        return settings
    except Exception:
        st.warning("settings.json 讀取失敗，已改用系統預設值。請檢查 config/settings.json。")
        return normalize_settings(DEFAULT_SETTINGS)


def db_connect() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def init_db() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    columns_sql = ",\n                ".join([f"{col} {typ}" for col, typ in DB_COLUMNS.items()])
    with db_connect() as conn:
        conn.execute(
            f"""
            CREATE TABLE IF NOT EXISTS quote_cases (
                {columns_sql}
            )
            """
        )
        existing_cols = {row[1] for row in conn.execute("PRAGMA table_info(quote_cases)").fetchall()}
        for col, typ in DB_COLUMNS.items():
            if col not in existing_cols:
                # SQLite 新增欄位不可重複指定 PRIMARY KEY。
                add_type = typ.replace(" PRIMARY KEY", "")
                conn.execute(f"ALTER TABLE quote_cases ADD COLUMN {col} {add_type}")
        conn.commit()


def ensure_app_ready() -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    if not CONFIG_PATH.exists():
        save_settings(DEFAULT_SETTINGS)
    init_db()


def to_float(value: Any, default: float = 0.0) -> float:
    if value is None or value == "":
        return float(default)
    try:
        return float(value)
    except Exception:
        return float(default)


def fmt_money(value: Any, prefix: str = "NT$") -> str:
    try:
        number = float(value or 0)
        if pd.isna(number):
            number = 0
    except Exception:
        number = 0
    return f"{prefix}{number:,.0f}"


def fmt_number(value: Any, decimals: int = 0) -> str:
    try:
        number = float(value or 0)
        if pd.isna(number):
            number = 0
    except Exception:
        number = 0
    return f"{number:,.{decimals}f}"


def next_quote_no() -> str:
    return f"Q{datetime.now().strftime('%Y%m%d%H%M%S')}"


def get_enabled_shipping(settings: Dict[str, Any]) -> List[Dict[str, Any]]:
    rows = settings.get("shipping_settings", [])
    enabled_rows = [r for r in rows if bool(r.get("enabled", True))]
    return enabled_rows or rows or DEFAULT_SETTINGS["shipping_settings"]


def get_shipping_config(settings: Dict[str, Any], source_country: str) -> Dict[str, Any]:
    for row in settings.get("shipping_settings", []):
        if row.get("source_country") == source_country:
            return row
    return {"source_country": source_country, "default_currency": "USD", "default_fee": 0, "enabled": True}


def get_enabled_users(settings: Dict[str, Any], roles: Optional[Iterable[str]] = None) -> List[Dict[str, Any]]:
    role_set = set(roles or [])
    rows = [u for u in settings.get("users", []) if bool(u.get("enabled", True))]
    if role_set:
        rows = [u for u in rows if u.get("role") in role_set]
    return rows


def display_user(user: Dict[str, Any]) -> str:
    role_label = "老闆" if user.get("role") == "admin" else "業務"
    return f"{user.get('display_name', user.get('username', ''))}（{role_label}）"


def parse_bool(value: Any, default: bool = True) -> bool:
    """將 data_editor 或 settings.json 的啟用欄位整理成布林值。"""
    if isinstance(value, bool):
        return value
    if value is None or value == "":
        return default
    text = str(value).strip().lower()
    if text in {"false", "0", "no", "n", "停用", "關閉", "disabled"}:
        return False
    if text in {"true", "1", "yes", "y", "啟用", "開啟", "enabled"}:
        return True
    return default


def sanitize_user_records(records: Iterable[Dict[str, Any]]) -> tuple[List[Dict[str, Any]], List[str]]:
    """整理帳號資料，防止重複帳號、空白帳號與沒有老闆帳號。"""
    users: List[Dict[str, Any]] = []
    errors: List[str] = []
    seen: set[str] = set()
    for idx, row in enumerate(records, start=1):
        username = str(row.get("username", "") or "").strip().lower()
        display_name = str(row.get("display_name", "") or "").strip()
        role = str(row.get("role", "sales") or "sales").strip().lower()
        password = str(row.get("password", "") or "").strip()
        enabled = parse_bool(row.get("enabled", True), default=True)

        # 完全空白列忽略，方便管理者在表格新增列時還沒填完。
        if not username and not display_name and not password:
            continue
        if not username:
            errors.append(f"第 {idx} 列缺少登入帳號。")
            continue
        if username in seen:
            errors.append(f"帳號重複：{username}")
            continue
        if role not in {"admin", "sales"}:
            errors.append(f"帳號 {username} 的角色錯誤，已略過。")
            continue
        if not password:
            errors.append(f"帳號 {username} 缺少密碼。")
            continue

        seen.add(username)
        users.append(
            {
                "username": username,
                "display_name": display_name or username,
                "role": role,
                "password": password,
                "enabled": enabled,
            }
        )

    if not users:
        errors.append("至少需要保留一個帳號。")
    if not any(u["role"] == "admin" and u["enabled"] for u in users):
        errors.append("至少需要保留一個啟用中的老闆／管理員帳號，避免所有人都無法進入後台。")
    return users, errors


def update_users_in_settings(settings: Dict[str, Any], users: List[Dict[str, Any]]) -> None:
    new_settings = json.loads(json.dumps(settings, ensure_ascii=False))
    new_settings["users"] = users
    save_settings(new_settings)


def normalize_addon_items(value: Any, fallback_amount: Any = 0, fallback_note: str = "") -> List[Dict[str, Any]]:
    """整理其他加裝項目，不限制項數；每筆保留品項與金額。"""
    items: List[Dict[str, Any]] = []
    if isinstance(value, str) and value.strip():
        try:
            value = json.loads(value)
        except Exception:
            value = []
    if isinstance(value, list):
        for item in value:
            if not isinstance(item, dict):
                continue
            name = str(item.get("name", "") or item.get("item", "") or item.get("品項名稱", "") or "").strip()
            amount = to_float(item.get("amount", item.get("金額", 0)))
            if name or amount:
                items.append({"name": name, "amount": round(amount)})

    # 舊版單一其他費用欄位相容。
    if not items and (to_float(fallback_amount) > 0 or str(fallback_note or "").strip()):
        items.append({"name": str(fallback_note or "其他加裝配備／其他費用").strip(), "amount": round(to_float(fallback_amount))})

    return items


def addon_items_total(items: List[Dict[str, Any]]) -> float:
    return sum(to_float(item.get("amount")) for item in items)


def addon_items_summary(items: List[Dict[str, Any]]) -> str:
    parts: List[str] = []
    for idx, item in enumerate(items, start=1):
        raw_name = str(item.get("name", "") or "").strip()
        amount = to_float(item.get("amount"))
        if raw_name or amount:
            name = raw_name or f"加裝項目{idx}"
            if amount:
                parts.append(f"{name} {fmt_money(amount)}")
            else:
                parts.append(name)
    return "；".join(parts)


def add_addon_fields(row: Dict[str, Any], items: List[Dict[str, Any]]) -> Dict[str, Any]:
    normalized = normalize_addon_items(items)
    total = addon_items_total(normalized)
    summary = addon_items_summary(normalized)
    row["addon_items_json"] = json.dumps(normalized, ensure_ascii=False)
    row["addon_total"] = round(total)
    row["other_fee"] = round(total)
    row["other_fee_note"] = summary

    # 舊版資料庫欄位仍保留前 5 項，完整加裝明細以 addon_items_json 儲存，不限制項數。
    for idx in range(1, 6):
        item = normalized[idx - 1] if idx <= len(normalized) else {"name": "", "amount": 0}
        row[f"addon_{idx}_name"] = str(item.get("name", "") or "").strip()
        row[f"addon_{idx}_amount"] = round(to_float(item.get("amount")))
    row["addon_items"] = normalized
    row["addon_summary"] = summary
    return row


def parse_addon_items_from_quote(quote: Dict[str, Any]) -> List[Dict[str, Any]]:
    if quote.get("addon_items"):
        return normalize_addon_items(quote.get("addon_items"))
    if quote.get("addon_items_json"):
        return normalize_addon_items(quote.get("addon_items_json"))
    items = []
    for idx in range(1, 6):
        name = str(quote.get(f"addon_{idx}_name", "") or "").strip()
        amount = to_float(quote.get(f"addon_{idx}_amount"))
        if name or amount:
            items.append({"name": name, "amount": round(amount)})
    return normalize_addon_items(items, fallback_amount=quote.get("other_fee", 0), fallback_note=quote.get("other_fee_note", ""))


def get_customs_rate_for_cc(engine_cc_category: str, settings: Dict[str, Any]) -> float:
    """依 CC 數區間取得關稅基數。"""
    q = settings["quote_settings"]
    category = str(engine_cc_category or "2000cc以下")
    if "以上" in category or "over" in category.lower():
        return to_float(q.get("customs_rate_over_2000"), 0.61)
    return to_float(q.get("customs_rate_under_2000"), 0.545)


def calculate_quote(input_data: Dict[str, Any], settings: Dict[str, Any]) -> Dict[str, Any]:
    q = settings["quote_settings"]
    price_term = input_data.get("price_term", "CIF")
    foreign_amount = to_float(input_data.get("foreign_amount"))
    car_exchange_rate = to_float(input_data.get("car_exchange_rate"))
    shipping_fee = to_float(input_data.get("shipping_fee")) if price_term == "FOB" else 0
    shipping_exchange_rate = to_float(input_data.get("shipping_exchange_rate")) if price_term == "FOB" else 0

    if price_term == "CIF":
        foreign_car_cost = foreign_amount * car_exchange_rate
        foreign_cost_formula = f"{fmt_number(foreign_amount)} × {fmt_number(car_exchange_rate, 4)}"
    else:
        foreign_car_cost = foreign_amount * car_exchange_rate + shipping_fee * shipping_exchange_rate
        foreign_cost_formula = (
            f"{fmt_number(foreign_amount)} × {fmt_number(car_exchange_rate, 4)} + "
            f"{fmt_number(shipping_fee)} × {fmt_number(shipping_exchange_rate, 4)}"
        )

    # v1.10：關稅改由系統依 CC 數自動計算。
    # 使用者指定公式：關稅 =（國外報價金額 + 1,500）× 關稅基數 × 匯率。
    engine_cc_category = str(input_data.get("engine_cc_category") or "2000cc以下")
    customs_rate = get_customs_rate_for_cc(engine_cc_category, settings)
    customs_add_amount_foreign = to_float(
        input_data.get("customs_add_amount_foreign"),
        q.get("customs_add_amount_foreign", 1500),
    )
    customs_amount = (foreign_amount + customs_add_amount_foreign) * customs_rate * car_exchange_rate
    customs_formula = (
        f"({fmt_number(foreign_amount)} + {fmt_number(customs_add_amount_foreign)}) × "
        f"{fmt_number(customs_rate, 4)} × {fmt_number(car_exchange_rate, 4)}"
    )

    # 奢侈稅：完稅價格 =（國外報價金額 + 關稅加計金額）× 匯率 + 關稅；超過門檻才課 10%。
    luxury_tax_threshold = to_float(q.get("luxury_tax_threshold"), 3000000)
    luxury_tax_rate = to_float(q.get("luxury_tax_rate"), 0.10)
    dutiable_price = (foreign_amount + customs_add_amount_foreign) * car_exchange_rate + customs_amount
    luxury_tax_amount = dutiable_price * luxury_tax_rate if dutiable_price > luxury_tax_threshold else 0
    luxury_tax_formula = (
        f"({fmt_money((foreign_amount + customs_add_amount_foreign) * car_exchange_rate)} + {fmt_money(customs_amount)}) × {fmt_number(luxury_tax_rate, 4)}"
        if luxury_tax_amount
        else f"完稅價格 {fmt_money(dutiable_price)} 未超過 {fmt_money(luxury_tax_threshold)}，不計奢侈稅"
    )

    # 服務費：可收台幣或美金。台幣直接計入；美金需乘匯率。
    service_fee_currency = str(input_data.get("service_fee_currency") or "TWD").upper()
    service_fee_amount = to_float(input_data.get("service_fee_amount"))
    if service_fee_currency == "TWD":
        service_fee_exchange_rate = 1.0
        service_fee_twd = service_fee_amount
        service_fee_formula = f"{fmt_number(service_fee_amount)}"
    else:
        service_fee_exchange_rate = to_float(input_data.get("service_fee_exchange_rate"), car_exchange_rate)
        service_fee_twd = service_fee_amount * service_fee_exchange_rate
        service_fee_formula = f"{fmt_number(service_fee_amount)} × {fmt_number(service_fee_exchange_rate, 4)}"
    service_fee_note = str(input_data.get("service_fee_note", "") or "")

    basic_vehicle_test_fee = to_float(input_data.get("basic_vehicle_test_fee"), q["basic_vehicle_test_fee"])
    is_special_vehicle = bool(input_data.get("is_special_vehicle", False))
    special_vehicle_test_surcharge = to_float(input_data.get("special_vehicle_test_surcharge")) if is_special_vehicle else 0
    vehicle_test_fee = basic_vehicle_test_fee + special_vehicle_test_surcharge

    cleaning_fee = to_float(input_data.get("cleaning_fee"))
    customs_trucking_fee = to_float(input_data.get("customs_trucking_fee"))
    sales_admin_fixed_profit_fee = to_float(input_data.get("sales_admin_fixed_profit_fee"))
    addon_items = normalize_addon_items(
        input_data.get("addon_items"),
        fallback_amount=input_data.get("other_fee", 0),
        fallback_note=str(input_data.get("other_fee_note", "") or ""),
    )
    addon_total = addon_items_total(addon_items)
    addon_summary = addon_items_summary(addon_items)
    other_fee = addon_total
    other_fee_note = addon_summary
    misc_total = cleaning_fee + customs_trucking_fee + sales_admin_fixed_profit_fee + addon_total

    total_cost = foreign_car_cost + customs_amount + luxury_tax_amount + service_fee_twd + vehicle_test_fee + misc_total
    normal_multiplier = to_float(q.get("normal_multiplier"), 1.07)
    cash_multiplier = to_float(q.get("cash_multiplier"), 1.05)
    company_base_price = total_cost * normal_multiplier
    cash_price = total_cost * cash_multiplier
    cash_price_difference = company_base_price - cash_price

    cash_transaction = bool(input_data.get("cash_transaction", False))
    transaction_type = "現金交易" if cash_transaction else "一般交易"
    multiplier = cash_multiplier if cash_transaction else normal_multiplier
    selected_price = cash_price if cash_transaction else company_base_price
    gross_profit = selected_price - total_cost

    result = dict(input_data)
    result.update(
        {
            "engine_cc_category": engine_cc_category,
            "foreign_car_cost": round(foreign_car_cost),
            "foreign_cost_formula": foreign_cost_formula,
            "customs_amount": round(customs_amount),
            "customs_rate": customs_rate,
            "customs_add_amount_foreign": round(customs_add_amount_foreign, 4),
            "customs_formula": customs_formula,
            "dutiable_price": round(dutiable_price),
            "luxury_tax_threshold": round(luxury_tax_threshold),
            "luxury_tax_rate": luxury_tax_rate,
            "luxury_tax_amount": round(luxury_tax_amount),
            "luxury_tax_formula": luxury_tax_formula,
            "service_fee_currency": service_fee_currency,
            "service_fee_amount": round(service_fee_amount),
            "service_fee_exchange_rate": service_fee_exchange_rate,
            "service_fee_twd": round(service_fee_twd),
            "service_fee_note": service_fee_note,
            "basic_vehicle_test_fee": round(basic_vehicle_test_fee),
            "is_special_vehicle": 1 if is_special_vehicle else 0,
            "special_vehicle_test_surcharge": round(special_vehicle_test_surcharge),
            "vehicle_test_fee": round(vehicle_test_fee),
            "cleaning_fee": round(cleaning_fee),
            "customs_trucking_fee": round(customs_trucking_fee),
            "sales_admin_fixed_profit_fee": round(sales_admin_fixed_profit_fee),
            "other_fee": round(other_fee),
            "other_fee_note": other_fee_note,
            "addon_items_json": json.dumps(addon_items, ensure_ascii=False),
            "addon_total": round(addon_total),
            "misc_total": round(misc_total),
            "total_cost": round(total_cost),
            "transaction_type": transaction_type,
            "multiplier": multiplier,
            "normal_multiplier": normal_multiplier,
            "cash_multiplier": cash_multiplier,
            "company_base_price": round(company_base_price),
            "cash_price": round(cash_price),
            "cash_price_difference": round(cash_price_difference),
            "selected_price": round(selected_price),
            # 舊欄位相容：base_price 固定代表公司底價；suggested_price 改存本次適用底價，不再顯示建議售價。
            "base_price": round(company_base_price),
            "suggested_price": round(selected_price),
            "gross_profit": round(gross_profit),
            "cash_transaction": 1 if cash_transaction else 0,
            "created_at": input_data.get("created_at") or datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
    )
    add_addon_fields(result, addon_items)
    return result

def save_quote_case(quote: Dict[str, Any]) -> str:
    init_db()
    quote_id = quote.get("id") or str(uuid.uuid4())
    quote_no = quote.get("quote_no") or next_quote_no()
    quote["id"] = quote_id
    quote["quote_no"] = quote_no
    quote["payload_json"] = json.dumps(quote, ensure_ascii=False)

    row = {col: quote.get(col) for col in RECORD_COLUMNS}
    with db_connect() as conn:
        placeholders = ", ".join(["?"] * len(RECORD_COLUMNS))
        columns_sql = ", ".join(RECORD_COLUMNS)
        update_sql = ", ".join([f"{col}=excluded.{col}" for col in RECORD_COLUMNS if col != "id"])
        conn.execute(
            f"""
            INSERT INTO quote_cases ({columns_sql}) VALUES ({placeholders})
            ON CONFLICT(id) DO UPDATE SET {update_sql}
            """,
            [row[col] for col in RECORD_COLUMNS],
        )
        conn.commit()
    return quote_id


def load_quote_cases() -> pd.DataFrame:
    init_db()
    with db_connect() as conn:
        df = pd.read_sql_query("SELECT * FROM quote_cases ORDER BY created_at DESC", conn)
    return df


def delete_quote_case(quote_id: str) -> None:
    with db_connect() as conn:
        conn.execute("DELETE FROM quote_cases WHERE id = ?", (quote_id,))
        conn.commit()


def get_selected_quote_from_df(df: pd.DataFrame, quote_id: str) -> Optional[Dict[str, Any]]:
    if df.empty or not quote_id:
        return None
    row = df.loc[df["id"] == quote_id]
    if row.empty:
        return None
    payload = row.iloc[0].get("payload_json")
    try:
        quote = json.loads(payload) if payload else row.iloc[0].dropna().to_dict()
    except Exception:
        quote = row.iloc[0].dropna().to_dict()
    # 舊資料補欄位。
    if "company_base_price" not in quote or quote.get("company_base_price") in (None, ""):
        quote["company_base_price"] = quote.get("base_price", 0)
    if "selected_price" not in quote or quote.get("selected_price") in (None, ""):
        quote["selected_price"] = quote.get("suggested_price", quote.get("base_price", 0))
    if "cash_price" not in quote or quote.get("cash_price") in (None, ""):
        total_cost = to_float(quote.get("total_cost"))
        cash_multiplier = to_float(quote.get("cash_multiplier"), 1.05)
        quote["cash_price"] = round(total_cost * cash_multiplier)
    if "cash_price_difference" not in quote or quote.get("cash_price_difference") in (None, ""):
        quote["cash_price_difference"] = round(to_float(quote.get("company_base_price")) - to_float(quote.get("cash_price")))
    add_addon_fields(quote, parse_addon_items_from_quote(quote))
    return quote


def login_page(settings: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    users = get_enabled_users(settings)
    st.title("二手車進口報價系統")
    st.caption(APP_VERSION)
    st.info("請先選擇使用者並輸入密碼。登入後，畫面右上方會出現『登出／切換帳號』按鈕。")
    st.write("登入後會依帳號顯示可使用的功能。")

    if not users:
        st.error("目前沒有啟用的使用者。請檢查 config/settings.json。")
        return None

    with st.form("login_form"):
        user_options = {f"{display_user(user)}｜帳號：{user.get('username', '')}": user for user in users}
        selected_label = st.selectbox("使用者", list(user_options.keys()))
        password = st.text_input("密碼", type="password")
        submitted = st.form_submit_button("登入", type="primary")
        if submitted:
            selected_user = user_options[selected_label]
            if password == str(selected_user.get("password", "")):
                st.session_state["current_user"] = selected_user
                st.rerun()
            else:
                st.error("密碼不正確")

    with st.expander("預設帳號說明", expanded=False):
        st.write("預設業務帳號：peter、cbc、lai、gary，預設密碼：1234。正式使用前請管理者修改密碼。")
    return None


def page_quote(settings: Dict[str, Any], current_user: Dict[str, Any], is_admin: bool = False) -> None:
    if is_admin:
        st.header("前台報價（老闆模式）")
        st.caption("老闆模式會顯示完整成本明細與固定雜費，可做單筆案件調整。")
    else:
        st.header("業務報價")
        st.caption("請輸入車輛基本資料、國外買車價格、CC 數、服務費與本次加選項目，系統會自動計算關稅、奢侈稅並產生公司底價。")

    q = settings["quote_settings"]
    currencies = settings.get("currencies", DEFAULT_SETTINGS["currencies"])
    enabled_shipping = get_enabled_shipping(settings)
    country_options = [r.get("source_country", "") for r in enabled_shipping if r.get("source_country")]
    if not country_options:
        country_options = DEFAULT_SETTINGS["source_countries"]

    users = get_enabled_users(settings)
    salesperson_options = [u.get("display_name", u.get("username", "")) for u in users]
    current_display_name = current_user.get("display_name", current_user.get("username", ""))

    with st.expander("車輛基本資料", expanded=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            quote_no = st.text_input("報價紀錄編號", value=st.session_state.get("quote_no", next_quote_no()))
            vehicle_name = st.text_input("車輛名稱／型號", placeholder="例如 Mercedes-Benz S500 / BMW X5")
            model_year = st.text_input("年份", placeholder="例如 2024")
        with col2:
            customer_name = st.text_input("客戶名稱", placeholder="選填")
            if is_admin:
                default_index = salesperson_options.index(current_display_name) if current_display_name in salesperson_options else 0
                salesperson = st.selectbox("業務姓名", salesperson_options, index=default_index)
            else:
                salesperson = current_display_name
                st.text_input("業務姓名", value=salesperson, disabled=True)
        with col3:
            source_country = st.selectbox("車輛來源國", country_options)
            price_term = st.selectbox("報價條件", ["CIF", "FOB"])
            engine_cc_category = st.selectbox("CC 數級距", ["2000cc以下", "2000cc以上"])
            quote_date = st.date_input("報價日期", value=date.today())

    shipping_config = get_shipping_config(settings, source_country)

    st.subheader("一、國外買車價格")
    col1, col2, col3 = st.columns(3)
    with col1:
        car_currency = st.selectbox("車價幣別", currencies, index=currencies.index("USD") if "USD" in currencies else 0)
        foreign_amount = st.number_input("國外報價金額", min_value=0.0, value=0.0, step=1000.0, format="%.0f")
    with col2:
        car_exchange_rate = st.number_input("車價匯率", min_value=0.0, value=32.0, step=0.01, format="%.4f")
    with col3:
        st.info("CIF 已含海運，不會再加海運費。FOB 才會顯示並計算海運欄位。")

    shipping_currency = ""
    shipping_fee = 0.0
    shipping_exchange_rate = 0.0
    if price_term == "FOB":
        st.markdown("**FOB 海運費**")
        col1, col2, col3 = st.columns(3)
        default_ship_currency = shipping_config.get("default_currency", "USD")
        default_ship_fee = float(shipping_config.get("default_fee", 0) or 0)
        with col1:
            shipping_currency = st.selectbox(
                "海運幣別",
                currencies,
                index=currencies.index(default_ship_currency) if default_ship_currency in currencies else 0,
                key=f"shipping_currency_{source_country}",
            )
        with col2:
            shipping_fee = st.number_input(
                "海運費",
                min_value=0.0,
                value=default_ship_fee,
                step=100.0,
                format="%.0f",
                key=f"shipping_fee_{source_country}",
            )
        with col3:
            shipping_exchange_rate = st.number_input("海運匯率", min_value=0.0, value=float(car_exchange_rate), step=0.01, format="%.4f")

    st.subheader("二、系統計算關稅")
    customs_rate_preview = get_customs_rate_for_cc(engine_cc_category, settings)
    customs_add_amount_foreign = to_float(q.get("customs_add_amount_foreign"), 1500)
    customs_amount = (foreign_amount + customs_add_amount_foreign) * customs_rate_preview * car_exchange_rate
    customs_formula_preview = (
        f"({fmt_number(foreign_amount)} + {fmt_number(customs_add_amount_foreign)}) × "
        f"{fmt_number(customs_rate_preview, 4)} × {fmt_number(car_exchange_rate, 4)}"
    )
    if is_admin:
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("CC 數級距", engine_cc_category)
        with col2:
            st.metric("關稅基數", fmt_number(customs_rate_preview, 4))
        with col3:
            st.metric("系統計算關稅", fmt_money(customs_amount))
        st.caption(f"關稅公式：{customs_formula_preview}")
    else:
        col1, col2 = st.columns(2)
        with col1:
            st.metric("CC 數級距", engine_cc_category)
        with col2:
            st.metric("系統計算關稅", fmt_money(customs_amount))
    customs_note = st.text_input("關稅備註", placeholder="例如報關行、估算日期、稅費依據，可選填")

    st.subheader("三、奢侈稅")
    dutiable_price_preview = (foreign_amount + customs_add_amount_foreign) * car_exchange_rate + customs_amount
    luxury_tax_threshold_preview = to_float(q.get("luxury_tax_threshold"), 3000000)
    luxury_tax_rate_preview = to_float(q.get("luxury_tax_rate"), 0.10)
    luxury_tax_amount_preview = dutiable_price_preview * luxury_tax_rate_preview if dutiable_price_preview > luxury_tax_threshold_preview else 0
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("完稅價格", fmt_money(dutiable_price_preview))
    with col2:
        st.metric("奢侈稅門檻", fmt_money(luxury_tax_threshold_preview))
    with col3:
        st.metric("奢侈稅", fmt_money(luxury_tax_amount_preview))
    if is_admin:
        st.caption("完稅價格 =（國外報價金額 + 1,500）× 匯率 + 系統計算關稅；超過門檻時，奢侈稅 = 完稅價格 × 奢侈稅率。")

    st.subheader("四、服務費")
    col1, col2, col3 = st.columns(3)
    with col1:
        service_fee_currency = st.selectbox("服務費幣別", ["TWD", "USD"])
    with col2:
        service_fee_amount = st.number_input("服務費金額", min_value=0.0, value=0.0, step=1000.0, format="%.0f")
    with col3:
        if service_fee_currency == "USD":
            service_fee_exchange_rate = st.number_input("服務費匯率", min_value=0.0, value=float(car_exchange_rate), step=0.01, format="%.4f")
        else:
            service_fee_exchange_rate = 1.0
            st.number_input("服務費匯率", min_value=0.0, value=1.0, step=0.01, format="%.4f", disabled=True)
    service_fee_note = st.text_input("服務費備註", placeholder="例如國外代辦服務費、台灣服務費等，可選填")

    st.subheader("五、特殊車測")
    if is_admin:
        st.caption("老闆可調整基本車測費；業務只輸入特殊車種與特殊車測加價。")
        col1, col2, col3 = st.columns(3)
        with col1:
            basic_vehicle_test_fee = st.number_input(
                "基本車測費",
                min_value=0.0,
                value=float(q["basic_vehicle_test_fee"]),
                step=10000.0,
                format="%.0f",
            )
        with col2:
            is_special_vehicle = st.checkbox("是否特殊車種", value=False)
        with col3:
            if is_special_vehicle:
                special_vehicle_test_surcharge = st.number_input(
                    "特殊車測加價",
                    min_value=0.0,
                    value=0.0,
                    step=10000.0,
                    format="%.0f",
                    key="admin_special_vehicle_test_surcharge",
                )
            else:
                special_vehicle_test_surcharge = 0.0
                st.number_input(
                    "特殊車測加價",
                    min_value=0.0,
                    value=0.0,
                    step=10000.0,
                    format="%.0f",
                    disabled=True,
                    key="admin_special_vehicle_test_surcharge_disabled",
                )

        with st.expander("老闆內部固定雜費調整", expanded=False):
            col1, col2, col3 = st.columns(3)
            with col1:
                cleaning_fee = st.number_input("整理費", min_value=0.0, value=float(q["cleaning_fee"]), step=10000.0, format="%.0f")
            with col2:
                customs_trucking_fee = st.number_input("報關／拖車費", min_value=0.0, value=float(q["customs_trucking_fee"]), step=10000.0, format="%.0f")
            with col3:
                sales_admin_fixed_profit_fee = st.number_input(
                    "業務獎金／行政費用／固定利潤",
                    min_value=0.0,
                    value=float(q["sales_admin_fixed_profit_fee"]),
                    step=10000.0,
                    format="%.0f",
                )
    else:
        basic_vehicle_test_fee = float(q["basic_vehicle_test_fee"])
        cleaning_fee = float(q["cleaning_fee"])
        customs_trucking_fee = float(q["customs_trucking_fee"])
        sales_admin_fixed_profit_fee = float(q["sales_admin_fixed_profit_fee"])
        col1, col2 = st.columns(2)
        with col1:
            is_special_vehicle = st.checkbox("是否特殊車種", value=False)
        with col2:
            if is_special_vehicle:
                special_vehicle_test_surcharge = st.number_input(
                    "特殊車測加價",
                    min_value=0.0,
                    value=0.0,
                    step=10000.0,
                    format="%.0f",
                    key="sales_special_vehicle_test_surcharge",
                )
            else:
                special_vehicle_test_surcharge = 0.0
                st.number_input(
                    "特殊車測加價",
                    min_value=0.0,
                    value=0.0,
                    step=10000.0,
                    format="%.0f",
                    disabled=True,
                    key="sales_special_vehicle_test_surcharge_disabled",
                )

    st.subheader("六、交易型態")
    col1, col2 = st.columns([1, 2])
    with col1:
        cash_transaction = st.checkbox("是否現金交易", value=False)
    with col2:
        notes = st.text_area("備註", height=80, placeholder="可填贈品、付款條件、特殊條件等")

    st.subheader("七、其他加裝配備")
    st.caption("加裝配備不限制項數；需要幾項就新增幾列，每一列都會獨立儲存品項與金額。沒有加裝可以留空。")
    addon_editor_default = pd.DataFrame([{"品項名稱": "", "金額": 0}])
    st.caption("範例：椅子 12,000、隔熱紙 32,000。按表格左下方的新增列，可以繼續加第 3 項、第 4 項、第 10 項以上。")
    edited_addons = st.data_editor(
        addon_editor_default,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="addon_items_editor",
        column_config={
            "品項名稱": st.column_config.TextColumn("加裝配備／其他項目", help="例如：椅子、隔熱紙、影音系統、外觀套件"),
            "金額": st.column_config.NumberColumn("金額", min_value=0, step=1000, format="%d"),
        },
    )
    addon_items: List[Dict[str, Any]] = []
    for row in edited_addons.fillna({"品項名稱": "", "金額": 0}).to_dict("records"):
        addon_name = str(row.get("品項名稱", "") or "").strip()
        addon_amount = round(to_float(row.get("金額", 0)))
        if addon_name or addon_amount:
            addon_items.append({"name": addon_name, "amount": addon_amount})
    addon_total = addon_items_total(addon_items)
    if addon_total > 0:
        st.metric("其他加裝配備合計", fmt_money(addon_total))
    other_fee = addon_total
    other_fee_note = addon_items_summary(addon_items)

    input_data = {
        "quote_no": quote_no,
        "created_at": datetime.combine(quote_date, datetime.now().time()).strftime("%Y-%m-%d %H:%M:%S"),
        "created_by_username": current_user.get("username", ""),
        "created_by_role": current_user.get("role", "sales"),
        "vehicle_name": vehicle_name,
        "model_year": model_year,
        "engine_cc_category": engine_cc_category,
        "vin": "",
        "customer_name": customer_name,
        "salesperson": salesperson,
        "source_country": source_country,
        "price_term": price_term,
        "car_currency": car_currency,
        "foreign_amount": foreign_amount,
        "car_exchange_rate": car_exchange_rate,
        "shipping_currency": shipping_currency,
        "shipping_fee": shipping_fee,
        "shipping_exchange_rate": shipping_exchange_rate,
        "customs_amount": customs_amount,
        "customs_add_amount_foreign": customs_add_amount_foreign,
        "customs_note": customs_note,
        "service_fee_currency": service_fee_currency,
        "service_fee_amount": service_fee_amount,
        "service_fee_exchange_rate": service_fee_exchange_rate,
        "service_fee_note": service_fee_note,
        "basic_vehicle_test_fee": basic_vehicle_test_fee,
        "is_special_vehicle": is_special_vehicle,
        "special_vehicle_test_surcharge": special_vehicle_test_surcharge,
        "cleaning_fee": cleaning_fee,
        "customs_trucking_fee": customs_trucking_fee,
        "sales_admin_fixed_profit_fee": sales_admin_fixed_profit_fee,
        "other_fee": other_fee,
        "other_fee_note": other_fee_note,
        "addon_items": addon_items,
        "cash_transaction": cash_transaction,
        "notes": notes,
    }
    quote = calculate_quote(input_data, settings)

    st.divider()
    st.subheader("試算結果")
    if is_admin:
        metric_cols = st.columns(7)
        metric_cols[0].metric("國外買車成本", fmt_money(quote["foreign_car_cost"]))
        metric_cols[1].metric("關稅", fmt_money(quote["customs_amount"]))
        metric_cols[2].metric("奢侈稅", fmt_money(quote["luxury_tax_amount"]))
        metric_cols[3].metric("服務費", fmt_money(quote["service_fee_twd"]))
        metric_cols[4].metric("車測費", fmt_money(quote["vehicle_test_fee"]))
        metric_cols[5].metric("雜費小計", fmt_money(quote["misc_total"]))
        metric_cols[6].metric("總成本", fmt_money(quote["total_cost"]))

        metric_cols = st.columns(4)
        metric_cols[0].metric("公司底價", fmt_money(quote["company_base_price"]))
        metric_cols[1].metric("公司底價（現金交易）", fmt_money(quote["cash_price"]))
        metric_cols[2].metric("底價價差", fmt_money(quote["cash_price_difference"]))
        metric_cols[3].metric("本次採用底價", fmt_money(quote["selected_price"]))

        with st.expander("查看計算明細", expanded=False):
            detail_df = pd.DataFrame(
                [
                    ["報價條件", quote["price_term"]],
                    ["CC 數級距", quote.get("engine_cc_category", "")],
                    ["國外成本公式", quote["foreign_cost_formula"]],
                    ["國外買車成本", fmt_money(quote["foreign_car_cost"])],
                    ["關稅基數", fmt_number(quote.get("customs_rate", 0), 4)],
                    ["關稅公式", quote.get("customs_formula", "")],
                    ["關稅金額", fmt_money(quote["customs_amount"])],
                    ["完稅價格", fmt_money(quote.get("dutiable_price", 0))],
                    ["奢侈稅公式", quote.get("luxury_tax_formula", "")],
                    ["奢侈稅", fmt_money(quote.get("luxury_tax_amount", 0))],
                    ["服務費幣別", quote.get("service_fee_currency", "")],
                    ["服務費原幣金額", fmt_number(quote.get("service_fee_amount", 0))],
                    ["服務費匯率", fmt_number(quote.get("service_fee_exchange_rate", 0), 4)],
                    ["服務費台幣", fmt_money(quote.get("service_fee_twd", 0))],
                    ["服務費備註", quote.get("service_fee_note", "")],
                    ["基本車測費", fmt_money(quote["basic_vehicle_test_fee"])],
                    ["特殊車測加價", fmt_money(quote["special_vehicle_test_surcharge"])],
                    ["車測費用", fmt_money(quote["vehicle_test_fee"])],
                    ["其他加裝配備合計", fmt_money(quote.get("addon_total", quote.get("other_fee", 0)))],
                    ["其他加裝明細", quote.get("addon_summary", quote.get("other_fee_note", ""))],
                    ["雜費小計", fmt_money(quote["misc_total"])],
                    ["總成本", fmt_money(quote["total_cost"])],
                    ["公司底價公式", f"{fmt_money(quote['total_cost'])} × {fmt_number(quote['normal_multiplier'], 4)}"],
                    ["公司底價", fmt_money(quote["company_base_price"])],
                    ["公司底價（現金交易）公式", f"{fmt_money(quote['total_cost'])} × {fmt_number(quote['cash_multiplier'], 4)}"],
                    ["公司底價（現金交易）", fmt_money(quote["cash_price"])],
                    ["價差", fmt_money(quote["cash_price_difference"])],
                ],
                columns=["項目", "內容"],
            )
            st.dataframe(detail_df, use_container_width=True, hide_index=True)
    else:
        if cash_transaction:
            metric_cols = st.columns(3)
            metric_cols[0].metric("公司底價", fmt_money(quote["company_base_price"]))
            metric_cols[1].metric("公司底價（現金交易）", fmt_money(quote["cash_price"]))
            metric_cols[2].metric("價差", fmt_money(quote["cash_price_difference"]))
            st.caption("價差 = 公司底價 − 公司底價（現金交易）。")
        else:
            metric_cols = st.columns(2)
            metric_cols[0].metric("交易型態", quote["transaction_type"])
            metric_cols[1].metric("公司底價", fmt_money(quote["company_base_price"]))

    if st.button("儲存報價紀錄", type="primary", use_container_width=True):
        quote_id = save_quote_case(quote)
        st.session_state["last_saved_quote_id"] = quote_id
        st.session_state["quote_no"] = next_quote_no()
        st.success(f"已儲存：{quote['quote_no']}")


def filter_records_for_user(df: pd.DataFrame, current_user: Dict[str, Any], is_admin: bool) -> pd.DataFrame:
    if is_admin or df.empty:
        return df
    username = str(current_user.get("username", ""))
    display_name = str(current_user.get("display_name", ""))
    mask = pd.Series(False, index=df.index)
    if "created_by_username" in df.columns:
        mask = mask | df["created_by_username"].fillna("").astype(str).eq(username)
    if "salesperson" in df.columns:
        mask = mask | df["salesperson"].fillna("").astype(str).eq(display_name)
    return df[mask]


def enrich_addon_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    enriched = df.copy()
    summaries: List[str] = []
    totals: List[float] = []
    for _, row in enriched.iterrows():
        quote = row.dropna().to_dict()
        items = parse_addon_items_from_quote(quote)
        summaries.append(addon_items_summary(items))
        total = addon_items_total(items)
        if total == 0:
            total = to_float(quote.get("addon_total", quote.get("other_fee", 0)))
        totals.append(round(total))
    enriched["addon_summary"] = summaries
    enriched["addon_total"] = totals
    return enriched


def format_display_table(df: pd.DataFrame, money_cols: List[str]) -> pd.DataFrame:
    display_df = df.copy()
    for col in money_cols:
        if col in display_df.columns:
            display_df[col] = display_df[col].apply(fmt_money)
    rename_map = {
        "created_at": "建立時間",
        "quote_no": "報價紀錄編號",
        "vehicle_name": "車輛名稱",
        "model_year": "年份",
        "engine_cc_category": "CC 數級距",
        "customer_name": "客戶",
        "salesperson": "業務",
        "source_country": "來源國",
        "price_term": "條件",
        "transaction_type": "交易型態",
        "customs_amount": "關稅",
        "customs_rate": "關稅基數",
        "dutiable_price": "完稅價格",
        "luxury_tax_amount": "奢侈稅",
        "service_fee_twd": "服務費",
        "service_fee_currency": "服務費幣別",
        "service_fee_amount": "服務費原幣",
        "special_vehicle_test_surcharge": "特殊車測加價",
        "other_fee": "其他加裝金額",
        "other_fee_note": "其他加裝說明",
        "addon_total": "其他加裝合計",
        "addon_summary": "其他加裝明細",
        "total_cost": "總成本",
        "company_base_price": "公司底價",
        "cash_price": "公司底價（現金交易）",
        "cash_price_difference": "價差",
        "selected_price": "本次採用底價",
        "created_by_username": "建立帳號",
    }
    return display_df.rename(columns=rename_map)


def page_records(current_user: Dict[str, Any], is_admin: bool = False) -> None:
    st.header("全部報價紀錄" if is_admin else "我的報價紀錄")
    df = load_quote_cases()
    df = filter_records_for_user(df, current_user, is_admin)
    if df.empty:
        st.info("目前尚無報價紀錄。")
        return

    search_text = st.text_input("搜尋", placeholder="輸入車名、客戶、業務、報價紀錄編號")
    filtered = enrich_addon_columns(df.copy())
    if search_text:
        search_text_lower = search_text.lower()
        mask = pd.Series(False, index=filtered.index)
        for col in ["quote_no", "vehicle_name", "customer_name", "salesperson", "source_country"]:
            if col in filtered.columns:
                mask = mask | filtered[col].fillna("").astype(str).str.lower().str.contains(search_text_lower, na=False)
        filtered = filtered[mask]

    if filtered.empty:
        st.warning("沒有符合搜尋條件的報價紀錄。")
        return

    if is_admin:
        show_cols = [
            "created_at",
            "quote_no",
            "vehicle_name",
            "model_year",
            "engine_cc_category",
            "customer_name",
            "salesperson",
            "source_country",
            "price_term",
            "transaction_type",
            "customs_amount",
            "luxury_tax_amount",
            "service_fee_twd",
            "special_vehicle_test_surcharge",
            "addon_total",
            "addon_summary",
            "total_cost",
            "company_base_price",
            "cash_price",
            "cash_price_difference",
            "selected_price",
            "created_by_username",
        ]
        money_cols = ["customs_amount", "luxury_tax_amount", "service_fee_twd", "special_vehicle_test_surcharge", "addon_total", "total_cost", "company_base_price", "cash_price", "cash_price_difference", "selected_price"]
    else:
        show_cols = [
            "created_at",
            "quote_no",
            "vehicle_name",
            "model_year",
            "engine_cc_category",
            "customer_name",
            "salesperson",
            "source_country",
            "price_term",
            "transaction_type",
            "customs_amount",
            "luxury_tax_amount",
            "service_fee_twd",
            "special_vehicle_test_surcharge",
            "addon_total",
            "addon_summary",
            "company_base_price",
            "cash_price",
            "cash_price_difference",
            "selected_price",
        ]
        money_cols = ["customs_amount", "luxury_tax_amount", "service_fee_twd", "special_vehicle_test_surcharge", "addon_total", "company_base_price", "cash_price", "cash_price_difference", "selected_price"]

    for col in ["company_base_price", "cash_price", "cash_price_difference", "selected_price"]:
        if col not in filtered.columns:
            filtered[col] = 0

    display_cols = [c for c in show_cols if c in filtered.columns]
    display_df = format_display_table(filtered[display_cols], money_cols)
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    if is_admin:
        export_df = filtered.drop(columns=["payload_json", "suggested_price", "base_price"], errors="ignore")
        csv_bytes = export_df.to_csv(index=False).encode("utf-8-sig")
        excel_buffer = BytesIO()
        export_df.to_excel(excel_buffer, index=False)

        col1, col2 = st.columns(2)
        with col1:
            st.download_button("下載 CSV", csv_bytes, file_name="quote_records.csv", mime="text/csv", use_container_width=True)
        with col2:
            st.download_button(
                "下載 Excel",
                excel_buffer.getvalue(),
                file_name="quote_records.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    if is_admin:
        st.subheader("單筆報價操作")
        options = [f"{row.quote_no}｜{row.vehicle_name or ''}｜{row.customer_name or ''}" for row in filtered.itertuples()]
        id_map = {options[i]: filtered.iloc[i]["id"] for i in range(len(options))}
        selected_label = st.selectbox("選擇報價", options)
        selected_quote = get_selected_quote_from_df(filtered, id_map[selected_label])
        if selected_quote:
            with st.expander("查看完整資料", expanded=False):
                st.json({k: v for k, v in selected_quote.items() if k not in {"suggested_price", "base_price"}}, expanded=False)
            if st.button("刪除此筆報價", use_container_width=True):
                delete_quote_case(selected_quote["id"])
                st.warning("已刪除，請重新整理頁面。")


def page_account_management(settings: Dict[str, Any], current_user: Dict[str, Any]) -> None:
    st.header("帳號管理")
    st.caption("只有老闆／管理員可以進入。這裡可以新增業務帳號、修改密碼、停用帳號或調整權限。")

    users = normalize_settings(settings).get("users", [])
    enabled_users = [u for u in users if u.get("enabled")]
    sales_users = [u for u in users if u.get("role") == "sales"]
    admin_users = [u for u in users if u.get("role") == "admin"]
    disabled_users = [u for u in users if not u.get("enabled")]

    metric_cols = st.columns(4)
    metric_cols[0].metric("全部帳號", len(users))
    metric_cols[1].metric("啟用帳號", len(enabled_users))
    metric_cols[2].metric("業務帳號", len(sales_users))
    metric_cols[3].metric("停用帳號", len(disabled_users))

    st.subheader("新增帳號")
    with st.form("add_user_form", clear_on_submit=True):
        col1, col2, col3, col4 = st.columns([1.2, 1.2, 1, 1.2])
        with col1:
            new_username = st.text_input("登入帳號", placeholder="例如 kevin")
        with col2:
            new_display_name = st.text_input("顯示名稱", placeholder="例如 Kevin")
        with col3:
            new_role = st.selectbox("角色", ["sales", "admin"], format_func=lambda x: "業務" if x == "sales" else "老闆／管理員")
        with col4:
            new_password = st.text_input("初始密碼", type="password", placeholder="請輸入密碼")
        new_enabled = st.checkbox("立即啟用", value=True)
        submitted_add = st.form_submit_button("新增帳號", type="primary")
        if submitted_add:
            username = str(new_username or "").strip().lower()
            if not username:
                st.error("請輸入登入帳號。")
            elif any(u.get("username") == username for u in users):
                st.error(f"帳號 {username} 已存在，請改用其他帳號。")
            elif not str(new_password or "").strip():
                st.error("請輸入初始密碼。")
            else:
                new_users = users + [
                    {
                        "username": username,
                        "display_name": str(new_display_name or username).strip(),
                        "role": new_role,
                        "password": str(new_password).strip(),
                        "enabled": bool(new_enabled),
                    }
                ]
                sanitized, errors = sanitize_user_records(new_users)
                if errors:
                    for error in errors:
                        st.error(error)
                else:
                    update_users_in_settings(settings, sanitized)
                    st.success(f"已新增帳號：{username}")
                    st.rerun()

    st.divider()
    st.subheader("現有帳號編輯")
    st.caption("可以直接在表格中修改顯示名稱、角色、密碼與啟用狀態；登入帳號建議建立後不要再改，避免舊報價紀錄對不到帳號。")

    users_df = pd.DataFrame(users)
    for col, default in [("username", ""), ("display_name", ""), ("role", "sales"), ("password", ""), ("enabled", True)]:
        if col not in users_df.columns:
            users_df[col] = default
    users_df = users_df[["username", "display_name", "role", "password", "enabled"]]

    edited_users = st.data_editor(
        users_df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "username": st.column_config.TextColumn("登入帳號", help="建議使用英文小寫，例如 peter、gary"),
            "display_name": st.column_config.TextColumn("顯示名稱", help="報價紀錄上看到的名稱"),
            "role": st.column_config.SelectboxColumn("角色", options=["admin", "sales"], help="admin=老闆/管理員；sales=業務"),
            "password": st.column_config.TextColumn("密碼", help="管理者可在此重設密碼"),
            "enabled": st.column_config.CheckboxColumn("啟用", help="取消勾選即可關閉該帳號登入權限"),
        },
        key="account_management_editor",
    )

    col1, col2 = st.columns([1, 2])
    with col1:
        save_accounts = st.button("儲存帳號變更", type="primary", use_container_width=True)
    with col2:
        st.info("停用帳號後，該帳號無法登入；舊報價紀錄仍會保留。")

    if save_accounts:
        records = edited_users.fillna({"username": "", "display_name": "", "role": "sales", "password": "", "enabled": True}).to_dict("records")
        sanitized, errors = sanitize_user_records(records)
        if errors:
            for error in errors:
                st.error(error)
        else:
            update_users_in_settings(settings, sanitized)
            st.success("帳號設定已儲存。")
            st.rerun()

    with st.expander("快速停用／啟用／刪除帳號", expanded=False):
        account_labels = {f"{u.get('display_name')}｜{u.get('username')}｜{'老闆' if u.get('role') == 'admin' else '業務'}｜{'啟用' if u.get('enabled') else '停用'}": u for u in users}
        if account_labels:
            selected_label = st.selectbox("選擇帳號", list(account_labels.keys()))
            selected_user = account_labels[selected_label]
            selected_username = selected_user.get("username")
            c1, c2, c3 = st.columns(3)
            with c1:
                if st.button("啟用此帳號", use_container_width=True):
                    new_users = [{**u, "enabled": True} if u.get("username") == selected_username else u for u in users]
                    sanitized, errors = sanitize_user_records(new_users)
                    if errors:
                        for error in errors:
                            st.error(error)
                    else:
                        update_users_in_settings(settings, sanitized)
                        st.success("已啟用帳號。")
                        st.rerun()
            with c2:
                if st.button("停用此帳號", use_container_width=True):
                    new_users = [{**u, "enabled": False} if u.get("username") == selected_username else u for u in users]
                    sanitized, errors = sanitize_user_records(new_users)
                    if errors:
                        for error in errors:
                            st.error(error)
                    else:
                        update_users_in_settings(settings, sanitized)
                        st.warning("已停用帳號。")
                        st.rerun()
            with c3:
                if st.button("刪除此帳號", use_container_width=True):
                    if selected_username == current_user.get("username"):
                        st.error("不能刪除目前登入中的自己。")
                    else:
                        new_users = [u for u in users if u.get("username") != selected_username]
                        sanitized, errors = sanitize_user_records(new_users)
                        if errors:
                            for error in errors:
                                st.error(error)
                        else:
                            update_users_in_settings(settings, sanitized)
                            st.warning("已刪除帳號；舊報價紀錄仍會保留建立帳號名稱。")
                            st.rerun()

    with st.expander("權限說明", expanded=False):
        st.markdown(
            """
- **老闆／管理員 admin**：可以看全部報價紀錄、後台參數、帳號管理、匯出與備份。
- **業務 sales**：只能前台報價、儲存報價、查看自己的報價紀錄。
- **停用帳號**：該帳號不能登入，但以前建立的報價紀錄不會刪除。
- **刪除帳號**：不會刪除報價資料，但之後不能再用該帳號登入。
            """
        )



def page_settings(settings: Dict[str, Any]) -> None:
    st.header("後台參數設定")
    st.caption("修改後會儲存在 config/settings.json，前台報價會自動套用。")

    quote_settings = settings["quote_settings"]
    with st.form("settings_form"):
        st.subheader("報價參數")
        col1, col2, col3 = st.columns(3)
        with col1:
            basic_vehicle_test_fee = st.number_input("基本車測費", value=float(quote_settings["basic_vehicle_test_fee"]), step=10000.0, format="%.0f")
            cleaning_fee = st.number_input("整理費", value=float(quote_settings["cleaning_fee"]), step=10000.0, format="%.0f")
            customs_trucking_fee = st.number_input("報關／拖車費", value=float(quote_settings["customs_trucking_fee"]), step=10000.0, format="%.0f")
        with col2:
            sales_admin_fixed_profit_fee = st.number_input(
                "業務獎金／行政費用／固定利潤",
                value=float(quote_settings["sales_admin_fixed_profit_fee"]),
                step=10000.0,
                format="%.0f",
            )
            other_fee_default = float(quote_settings.get("other_fee_default", 0))
        with col3:
            normal_multiplier = st.number_input("公司底價倍率", value=float(quote_settings["normal_multiplier"]), step=0.01, format="%.4f")
            cash_multiplier = st.number_input("現金交易倍率", value=float(quote_settings["cash_multiplier"]), step=0.01, format="%.4f")

        st.subheader("關稅／奢侈稅參數")
        tax_col1, tax_col2, tax_col3, tax_col4, tax_col5 = st.columns(5)
        with tax_col1:
            customs_rate_under_2000 = st.number_input(
                "2000cc 以下關稅基數",
                value=float(quote_settings.get("customs_rate_under_2000", 0.545)),
                step=0.001,
                format="%.4f",
            )
        with tax_col2:
            customs_rate_over_2000 = st.number_input(
                "2000cc 以上關稅基數",
                value=float(quote_settings.get("customs_rate_over_2000", 0.61)),
                step=0.001,
                format="%.4f",
            )
        with tax_col3:
            customs_add_amount_foreign = st.number_input(
                "關稅加計金額（外幣）",
                value=float(quote_settings.get("customs_add_amount_foreign", 1500)),
                step=100.0,
                format="%.0f",
            )
        with tax_col4:
            luxury_tax_threshold = st.number_input(
                "奢侈稅門檻",
                value=float(quote_settings.get("luxury_tax_threshold", 3000000)),
                step=100000.0,
                format="%.0f",
            )
        with tax_col5:
            luxury_tax_rate = st.number_input(
                "奢侈稅率",
                value=float(quote_settings.get("luxury_tax_rate", 0.10)),
                step=0.01,
                format="%.4f",
            )
        st.caption("目前預設 2000cc 以下關稅基數為 0.545；2000cc 以上預設 0.61。")

        st.subheader("來源國與海運預設")
        shipping_df = pd.DataFrame(settings.get("shipping_settings", DEFAULT_SETTINGS["shipping_settings"]))
        edited_shipping = st.data_editor(
            shipping_df,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "source_country": st.column_config.TextColumn("來源國"),
                "default_currency": st.column_config.SelectboxColumn("預設海運幣別", options=settings.get("currencies", DEFAULT_SETTINGS["currencies"])),
                "default_fee": st.column_config.NumberColumn("預設海運費", min_value=0, step=1000, format="%d"),
                "enabled": st.column_config.CheckboxColumn("是否啟用"),
            },
        )

        st.subheader("幣別設定")
        currencies_text = st.text_area(
            "幣別清單（一行一個）",
            value="\n".join(settings.get("currencies", DEFAULT_SETTINGS["currencies"])),
            height=140,
        )

        submitted = st.form_submit_button("儲存後台設定", type="primary")
        if submitted:
            new_currencies = [x.strip().upper() for x in currencies_text.splitlines() if x.strip()]
            if not new_currencies:
                new_currencies = DEFAULT_SETTINGS["currencies"]

            new_settings = {
                "app_settings": {"version": APP_VERSION, "multi_user_enabled": True},
                "users": settings.get("users", DEFAULT_SETTINGS["users"]),
                "quote_settings": {
                    "basic_vehicle_test_fee": round(basic_vehicle_test_fee),
                    "cleaning_fee": round(cleaning_fee),
                    "customs_trucking_fee": round(customs_trucking_fee),
                    "sales_admin_fixed_profit_fee": round(sales_admin_fixed_profit_fee),
                    "other_fee_default": round(other_fee_default),
                    "normal_multiplier": float(normal_multiplier),
                    "cash_multiplier": float(cash_multiplier),
                    "customs_rate_under_2000": float(customs_rate_under_2000),
                    "customs_rate_over_2000": float(customs_rate_over_2000),
                    "customs_add_amount_foreign": float(customs_add_amount_foreign),
                    "luxury_tax_threshold": round(luxury_tax_threshold),
                    "luxury_tax_rate": float(luxury_tax_rate),
                },
                "currencies": new_currencies,
                "source_countries": [str(x).strip() for x in edited_shipping["source_country"].tolist() if str(x).strip()],
                "shipping_settings": edited_shipping.fillna({"default_currency": "USD", "default_fee": 0, "enabled": True}).to_dict("records"),
            }
            for row in new_settings["shipping_settings"]:
                row["source_country"] = str(row.get("source_country", "")).strip()
                row["default_currency"] = str(row.get("default_currency", "USD")).strip().upper() or "USD"
                row["default_fee"] = float(row.get("default_fee", 0) or 0)
                row["enabled"] = bool(row.get("enabled", True))
            save_settings(new_settings)
            st.success("後台設定已儲存，請重新登入或切換頁面使用新設定。")
            st.rerun()

    with st.expander("目前 settings.json", expanded=False):
        safe_settings = json.loads(json.dumps(settings, ensure_ascii=False))
        for user in safe_settings.get("users", []):
            user["password"] = "******"
        st.json(safe_settings)


def page_backup() -> None:
    st.header("匯出／備份")
    st.write("建議定期備份資料庫與設定檔。多人使用時，請至少每週備份一次。")
    col1, col2 = st.columns(2)
    with col1:
        if CONFIG_PATH.exists():
            st.download_button(
                "下載設定檔 settings.json",
                data=CONFIG_PATH.read_bytes(),
                file_name="settings.json",
                mime="application/json",
                use_container_width=True,
            )
    with col2:
        if DB_PATH.exists():
            st.download_button(
                "下載報價資料庫 quotes.db",
                data=DB_PATH.read_bytes(),
                file_name="quotes.db",
                mime="application/octet-stream",
                use_container_width=True,
            )
    df = load_quote_cases()
    if not df.empty:
        excel_buffer = BytesIO()
        df.drop(columns=["payload_json", "suggested_price", "base_price"], errors="ignore").to_excel(excel_buffer, index=False)
        st.download_button(
            "下載全部報價 Excel",
            data=excel_buffer.getvalue(),
            file_name="all_quote_records.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


def page_help(current_user: Optional[Dict[str, Any]] = None, is_admin: bool = False) -> None:
    st.header("使用說明")
    if is_admin:
        st.markdown(
            """
### 一鍵式操作
1. 第一次使用：在資料夾中雙擊 `start.command`。
2. 系統會自動建立 `.venv`、安裝套件，並啟動瀏覽器頁面。
3. 多人使用時，請把這台電腦當主機，其他業務用同一個區網連到 `http://主機IP:8501`。

### v1.11 業務隱藏公式輔助版
- 預設 5 個帳號：老闆 `darren`，業務 `peter`、`cbc`、`lai`、`gary`。
- 老闆預設密碼：`oma1688`。
- 業務預設密碼：`1234`。
- 老闆可在「帳號管理」新增帳號、修改姓名、角色與密碼，也可以停用或刪除帳號。
- 登入後，主畫面右上方有「登出／切換帳號」按鈕；左側功能選單也有同樣按鈕。

### 業務操作流程
- 輸入車輛基本資料、國外買車價格、CC 數、服務費、特殊車測、是否現金交易與其他加裝配備。
- 車輛基本資料已移除 VIN／車身號碼欄位。
- 可選「是否特殊車種」，並輸入特殊車測加價。
- 其他加裝配備不限制項數，可依實際需要新增多項，每項都有品項名稱與金額。
- 不再輸入「建議售價」。
- 一般交易顯示公司底價。
- 現金交易會顯示公司底價、公司底價（現金交易），以及兩者價差。
- 業務只能看到自己建立的報價紀錄。

### 老闆模式
- 老闆登入後可看到完整成本明細。
- 老闆可調整基本車測費、特殊車測加價、固定雜費、關稅基數、奢侈稅門檻與奢侈稅率。
- 關稅公式與完稅價格／奢侈稅公式的輔助說明只在老闆模式顯示，業務模式不顯示。
- 老闆可看到全部業務的報價紀錄，並可匯出 Excel／CSV。

### 多人連線方式
- 同一個辦公室 Wi-Fi／有線網路：業務開瀏覽器輸入 `http://主機IP:8501`。
- 不同網路或外出使用：不要用主機的內部 IP，需要部署到雲端主機，或使用 VPN／內網穿透服務。
- 正式多人長期使用，建議改成雲端主機版本，避免主機電腦關機後全部人不能使用。

### 報價邏輯
- CIF：國外買車成本 = 國外報價金額 × 車價匯率。
- FOB：國外買車成本 = 國外報價金額 × 車價匯率 + 海運費 × 海運匯率。
- 關稅：系統依 CC 數自動計算。2000cc 以下預設基數 0.545；2000cc 以上預設基數 0.61。公式 =（國外報價金額 + 1,500）× 關稅基數 × 匯率。
- 奢侈稅：完稅價格 =（國外報價金額 + 1,500）× 匯率 + 關稅；若完稅價格超過 3,000,000，奢侈稅 = 完稅價格 × 10%。
- 服務費：可選 TWD 或 USD；TWD 直接計入，USD 需輸入服務費匯率後換算台幣。
- 車測費 = 基本車測費 + 特殊車測加價。
- 雜費 = 整理費 + 報關／拖車費 + 業務獎金／行政費用／固定利潤 + 其他加裝配備合計。
- 總成本 = 國外買車成本 + 關稅 + 奢侈稅 + 服務費 + 車測費 + 雜費。
- 公司底價 = 總成本 × 後台公司底價倍率，預設 1.07。
- 公司底價（現金交易） = 總成本 × 後台現金交易倍率，預設 1.05。
- 價差 = 公司底價 − 公司底價（現金交易）。

### 資料位置
- 報價紀錄：`data/quotes.db`
- 後台設定：`config/settings.json`
- 一鍵啟動：`start.command`
            """
        )
    else:
        st.markdown(
            """
### 業務操作流程
1. 登入自己的帳號。
2. 點選「前台報價」。
3. 輸入車輛基本資料、CC 數、國外買車價格與服務費；系統會自動計算關稅與奢侈稅。
4. 若有特殊車種，請輸入特殊車測加價。
5. 選擇是否現金交易。
6. 在「七、其他加裝配備」可依實際需要新增多項，每項輸入品項與金額。
7. 系統會產生公司底價；現金交易會另外顯示公司底價（現金交易）與價差。
8. 按「儲存報價紀錄」即可。

### 帳號與紀錄
- 右上方可按「登出／切換帳號」。
- 左側「報價紀錄」可以查詢自己建立的報價。
- 車輛基本資料不需要輸入 VIN／車身號碼。

### 連線方式
- 在公司同一個 Wi-Fi／有線網路時，請向管理者索取系統網址，例如 `http://主機IP:8501`。
- 在外面或不同網路時，需要使用公司提供的雲端網址或 VPN 連線網址。
            """
        )

def main() -> None:
    st.set_page_config(page_title="二手車進口報價系統", page_icon="🚗", layout="wide", initial_sidebar_state="expanded")
    ensure_app_ready()
    settings = load_settings()

    current_user = st.session_state.get("current_user")
    if not current_user:
        login_page(settings)
        return

    # 重新從設定檔同步使用者資料，避免後台改名後 session 仍用舊資料。
    matched_user = next((u for u in get_enabled_users(settings) if u.get("username") == current_user.get("username")), None)
    if matched_user:
        current_user = matched_user
        st.session_state["current_user"] = matched_user
    else:
        st.session_state.pop("current_user", None)
        st.warning("目前帳號已被停用或不存在，請重新登入。")
        login_page(settings)
        return
    is_admin = current_user.get("role") == "admin"

    # 主畫面上方固定顯示登入身分與登出／切換帳號，避免使用者找不到側邊欄。
    user_col, logout_col = st.columns([5, 1])
    with user_col:
        st.markdown(f"**目前登入：{display_user(current_user)}**")
    with logout_col:
        if st.button("登出／切換帳號", type="secondary", use_container_width=True, key="top_logout_button"):
            st.session_state.pop("current_user", None)
            st.rerun()
    st.divider()

    st.sidebar.title("二手車進口報價系統")
    st.sidebar.caption(APP_VERSION)
    st.sidebar.success(f"已登入：{display_user(current_user)}")
    if st.sidebar.button("登出／切換帳號", key="sidebar_logout_button"):
        st.session_state.pop("current_user", None)
        st.rerun()

    if is_admin:
        page_options = ["前台報價", "報價紀錄", "帳號管理", "後台參數設定", "匯出／備份", "使用說明"]
    else:
        page_options = ["前台報價", "報價紀錄", "使用說明"]
        pass

    page = st.sidebar.radio("功能選單", page_options)

    if page == "前台報價":
        page_quote(settings, current_user=current_user, is_admin=is_admin)
    elif page == "報價紀錄":
        page_records(current_user=current_user, is_admin=is_admin)
    elif page == "帳號管理" and is_admin:
        page_account_management(settings, current_user=current_user)
    elif page == "後台參數設定" and is_admin:
        page_settings(settings)
    elif page == "匯出／備份" and is_admin:
        page_backup()
    else:
        page_help(current_user=current_user, is_admin=is_admin)


if __name__ == "__main__":
    main()
