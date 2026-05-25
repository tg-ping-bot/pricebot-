"""
АлкоПартнёр — Telegram-бот для поиска по прайс-листу МАИ + пиво
Версия: финальная

Установка:
  pip install python-telegram-bot pandas xlrd==1.2.0 openpyxl requests

Запуск:
  python bot.py
"""

import os, re, json, logging, sys, requests, pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, BotCommand
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler,
    CallbackQueryHandler, filters, ContextTypes
)

# ══════════════════════════════════════════════════════════════════
TOKEN          = os.environ.get("BOT_TOKEN", "8974641448:AAFhIYLya0lVhRsldJj4a2UXNKvXv0rQPD8")
GOOGLE_API_KEY = os.environ.get("GOOGLE_API_KEY", "AIzaSyCJzZvFuEAz534XWnqHj3lysKWjDox9dKU")
GOOGLE_CX      = os.environ.get("GOOGLE_CX",      "641ca1d3242234c40")
PRICE_XLS      = "price.xls"
PRICE_BEER     = "pricepivo.xls"
ORDERS_FILE    = "orders.txt"
ORDERS_JSON    = "orders.json"
PHOTO_CACHE    = "photo_cache.json"
GOOGLE_SHEET_ID   = os.environ.get("GOOGLE_SHEET_ID", "")
GOOGLE_CREDS_JSON = os.environ.get("GOOGLE_CREDS_JSON", "")
# ══════════════════════════════════════════════════════════════════

logging.basicConfig(format="%(asctime)s | %(levelname)s | %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

df_price     = None
_name_index: list[str] = []
df_beer      = None
_beer_index: list[str] = []
_photo_cache: dict = {}

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124.0.0.0"}


# ─────────────────────────────────────────────────────────────────
# УТИЛИТЫ
# ─────────────────────────────────────────────────────────────────

def normalize(t: str) -> str:
    return re.sub(r"\s+", " ", t.lower().strip().replace(".", ","))


def load_photo_cache():
    global _photo_cache
    try:
        if os.path.exists(PHOTO_CACHE):
            with open(PHOTO_CACHE, encoding="utf-8") as f:
                _photo_cache = json.load(f)
    except Exception:
        _photo_cache = {}

load_photo_cache()


# ─────────────────────────────────────────────────────────────────
# GOOGLE SHEETS
# ─────────────────────────────────────────────────────────────────

_gs_client = None
_gs_sheet  = None

def get_sheet():
    """Возвращает объект листа Google Sheets. Переподключается если нужно."""
    global _gs_client, _gs_sheet
    if not GOOGLE_SHEET_ID or not GOOGLE_CREDS_JSON:
        return None
    try:
        if _gs_sheet is None:
            creds_dict = json.loads(GOOGLE_CREDS_JSON)
            creds = Credentials.from_service_account_info(
                creds_dict,
                scopes=[
                    "https://spreadsheets.google.com/feeds",
                    "https://www.googleapis.com/auth/drive",
                ]
            )
            _gs_client = gspread.authorize(creds)
            spreadsheet = _gs_client.open_by_key(GOOGLE_SHEET_ID)
            _gs_sheet   = spreadsheet.sheet1
        return _gs_sheet
    except Exception as e:
        logger.error(f"Google Sheets connect: {e}")
        _gs_sheet = None
        return None


def save_order_to_sheets(cart, buy_tot, sell_tot, margin):
    """Дописывает строку заказа в Google Sheets."""
    sheet = get_sheet()
    if sheet is None:
        return False
    try:
        now      = datetime.now().strftime("%d.%m.%Y %H:%M")
        # Формируем список позиций в одну строку
        items_str = "; ".join(
            f"{i['name']} ×{i['qty']} по {i['price']:.2f}₽"
            + (f" → продажа {i['sell_price']:.2f}₽" if i.get("sell_price") else "")
            for i in cart["items"]
        )
        sheet.append_row([
            now,
            cart["client"],
            round(buy_tot,  2),
            round(sell_tot, 2),
            round(margin,   2),
            items_str,
        ])
        logger.info("Заказ записан в Google Sheets")
        return True
    except Exception as e:
        logger.error(f"save_order_to_sheets: {e}")
        return False


def load_orders_from_sheets() -> list:
    """Загружает все заказы из Google Sheets."""
    sheet = get_sheet()
    if sheet is None:
        return []
    try:
        rows   = sheet.get_all_values()
        orders = []
        for row in rows[1:]:  # пропускаем заголовок
            if len(row) < 5 or not row[0]:
                continue
            try:
                # Парсим дату обратно в нужный формат
                dt_str = row[0]
                try:
                    dt_obj = datetime.strptime(dt_str, "%d.%m.%Y %H:%M")
                    dt_iso = dt_obj.strftime("%Y-%m-%d %H:%M:%S")
                except Exception:
                    dt_iso = dt_str

                orders.append({
                    "dt":         dt_iso,
                    "client":     row[1] if len(row) > 1 else "?",
                    "buy_total":  float(row[2]) if len(row) > 2 and row[2] else 0,
                    "sell_total": float(row[3]) if len(row) > 3 and row[3] else 0,
                    "margin":     float(row[4]) if len(row) > 4 and row[4] else 0,
                    "items_str":  row[5] if len(row) > 5 else "",
                    "items":      [],  # детали позиций не храним отдельно в Sheets
                })
            except Exception:
                continue
        return orders
    except Exception as e:
        logger.error(f"load_orders_from_sheets: {e}")
        return []



def find_photo(name: str, barcode: str = "") -> str | None:
    key = name[:80]
    if key in _photo_cache:
        return _photo_cache[key]
    url = None
    if barcode and barcode not in ("", "nan", "0"):
        try:
            r = requests.get(
                f"https://world.openfoodfacts.org/api/v0/product/{barcode}.json",
                headers=HEADERS, timeout=6)
            p = r.json().get("product", {})
            url = p.get("image_url") or p.get("image_front_url")
        except Exception:
            pass
    if not url and GOOGLE_API_KEY and GOOGLE_CX:
        try:
            short = " ".join(re.sub(r"\(.*?\)", "", name).strip().split()[:4])
            r = requests.get(
                "https://www.googleapis.com/customsearch/v1",
                params={"key": GOOGLE_API_KEY, "cx": GOOGLE_CX,
                        "q": short + " бутылка", "searchType": "image",
                        "num": 1, "imgSize": "medium"},
                timeout=8)
            items = r.json().get("items", [])
            if items:
                url = items[0].get("link")
        except Exception:
            pass
    _photo_cache[key] = url
    try:
        with open(PHOTO_CACHE, "w", encoding="utf-8") as f:
            json.dump(_photo_cache, f, ensure_ascii=False)
    except Exception:
        pass
    return url


# ─────────────────────────────────────────────────────────────────
# ЗАГРУЗКА ПРАЙСОВ
# ─────────────────────────────────────────────────────────────────

def load_price_from_xls(xls_path: str):
    global df_price, _name_index
    try:
        import xlrd
        wb   = xlrd.open_workbook(xls_path)
        rows = []
        for si in range(wb.nsheets):
            sh = wb.sheet_by_index(si)
            for rx in range(sh.nrows):
                row = sh.row_values(rx)
                while len(row) < 11:
                    row.append("")
                rows.append(row[:11])

        raw = pd.DataFrame(rows, columns=[
            "drop", "name", "stock", "barcode", "category",
            "drink_type", "color", "country", "price", "promo_price", "promo_cond"
        ]).drop(columns=["drop"])

        raw["price"] = pd.to_numeric(raw["price"], errors="coerce")
        df = raw[raw["price"].notna() & (raw["price"] > 0)].copy()
        df["name"]        = df["name"].astype(str).str.strip()
        df = df[df["name"] != ""]
        df = df[~df["name"].str.upper().isin(["ИТОГО", "ВСЕГО", "NAN", "НАИМЕНОВАНИЕ"])]
        df["promo_price"] = pd.to_numeric(df["promo_price"], errors="coerce")
        df["stock"]       = pd.to_numeric(df["stock"], errors="coerce").fillna(0).astype(int)
        df["promo_cond"]  = df["promo_cond"].fillna("").astype(str).str.strip()
        df["barcode"]     = df["barcode"].astype(str).str.split(".").str[0].str.strip()

        df_price    = df.reset_index(drop=True)
        _name_index = [normalize(n) for n in df_price["name"]]
        logger.info(f"Алко загружено: {len(df_price)} позиций")
        return True, f"✅ Алкогольный прайс загружен: *{len(df_price)}* позиций"
    except Exception as e:
        logger.error(f"load_xls: {e}")
        return False, f"❌ Ошибка загрузки прайса: {e}"


def load_beer_from_xlsx(xlsx_path: str):
    """Загружает пивной прайс. Поддерживает .xls и .xlsx форматы."""
    global df_beer, _beer_index
    try:
        ext = os.path.splitext(xlsx_path)[1].lower()

        if ext == ".xls":
            # Читаем через xlrd (старый формат)
            import xlrd
            wb   = xlrd.open_workbook(xlsx_path)
            rows = []
            sh   = wb.sheet_by_index(0)
            for rx in range(sh.nrows):
                rows.append(sh.row_values(rx)[:3])
            raw = pd.DataFrame(rows, columns=["name", "price_factory", "price_nw"])
        else:
            # Читаем через openpyxl (новый формат)
            raw = pd.read_excel(xlsx_path, header=None, engine="openpyxl")
            raw = raw.iloc[:, :3]
            raw.columns = ["name", "price_factory", "price_nw"]

        raw["price_factory"] = pd.to_numeric(raw["price_factory"], errors="coerce")
        raw["price_nw"]      = pd.to_numeric(raw["price_nw"],      errors="coerce")

        # Оставляем только строки с реальной ценой
        df = raw[raw["price_factory"].notna() & (raw["price_factory"] > 0)].copy()
        df["name"] = df["name"].astype(str).str.strip()
        df = df[df["name"] != ""]
        df = df[~df["name"].str.upper().isin(["NAN", "НАИМЕНОВАНИЕ", "НАЗВАНИЕ",
                                               "ЗАВОД НАЛ", "ЦЕНА СЕВЕРО ЗАПАД"])]
        # Убираем строки-разделители (только объём без цены — например "0,45 стекло")
        df = df[df["name"].str.len() > 5]

        df_beer     = df.reset_index(drop=True)
        _beer_index = [normalize(n) for n in df_beer["name"]]
        logger.info(f"Пиво загружено: {len(df_beer)} позиций")
        return True, f"✅ Пивной прайс загружен: *{len(df_beer)}* позиций"
    except Exception as e:
        logger.error(f"load_beer: {e}")
        return False, f"❌ Ошибка загрузки пивного прайса: {e}"


# ─────────────────────────────────────────────────────────────────
# ПОИСК
# ─────────────────────────────────────────────────────────────────

def parse_price_filter(query: str):
    q = query.strip()
    min_p = max_p = None
    for pat, handler in [
        (r'от\s+(\d+)\s+до\s+(\d+)', lambda m: (float(m[1]), float(m[2]))),
        (r'(\d+)\s*[-–]\s*(\d+)',     lambda m: (float(m[1]), float(m[2]))),
    ]:
        m = re.search(pat, q, re.IGNORECASE)
        if m:
            min_p, max_p = handler(m)
            q = (q[:m.start()] + q[m.end():]).strip()
            return q, min_p, max_p
    for pat, key in [
        (r'до\s+(\d+)',        'max'),
        (r'дешевле\s+(\d+)',   'max'),
        (r'не дороже\s+(\d+)', 'max'),
        (r'от\s+(\d+)',        'min'),
        (r'дороже\s+(\d+)',    'min'),
        (r'не дешевле\s+(\d+)','min'),
    ]:
        m = re.search(pat, q, re.IGNORECASE)
        if m:
            val = float(m.group(1))
            if key == 'max': max_p = val
            else:            min_p = val
            q = (q[:m.start()] + q[m.end():]).strip()
            return q, min_p, max_p
    return q, None, None


def _search(index, df, query, min_price=None, max_price=None, price_col="price"):
    words = normalize(query).split()
    if not words:
        return []
    results = [df.iloc[i] for i, n in enumerate(index) if all(w in n for w in words)]
    if not results:
        thr = max(1, len(words) // 2)
        scored = sorted(
            [(sum(1 for w in words if w in n), i) for i, n in enumerate(index)],
            key=lambda x: -x[0]
        )
        results = [df.iloc[i] for s, i in scored if s >= thr]
    if min_price is not None:
        results = [r for r in results if pd.notna(r[price_col]) and float(r[price_col]) >= min_price]
    if max_price is not None:
        results = [r for r in results if pd.notna(r[price_col]) and float(r[price_col]) <= max_price]
    return results


def do_search(query, min_price=None, max_price=None):
    if df_price is None:
        return []
    return _search(_name_index, df_price, query, min_price, max_price, "price")


def do_search_beer(query, min_price=None, max_price=None):
    if df_beer is None:
        return []
    return _search(_beer_index, df_beer, query, min_price, max_price, "price_factory")


# ─────────────────────────────────────────────────────────────────
# ФОРМАТИРОВАНИЕ
# ─────────────────────────────────────────────────────────────────

def fmt_card(row) -> str:
    p = float(row["price"])
    p_str = f"{p:.0f}" if p == int(p) else f"{p:.2f}"
    lines = [f"📦 *{row['name']}*", f"💰 Цена МАИ: *{p_str} ₽*"]
    if pd.notna(row.get("promo_price")) and float(row.get("promo_price", 0)) > 0:
        pp = float(row["promo_price"])
        cond = f" — _{row['promo_cond']}_" if row.get("promo_cond") else ""
        lines.append(f"🎯 Маркетинг: *{pp:.2f} ₽*{cond}")
    lines.append(f"📊 Остаток: {int(row['stock'])} шт.")
    return "\n".join(lines)


def fmt_cart(cart, title="🛒 *Корзина*") -> str:
    if not cart["items"]:
        return "🛒 Корзина пуста"
    client = f" — {cart['client']}" if cart["client"] else ""
    lines  = [f"{title}{client}"]

    lines.append("\n📦 *ЗАКУПКА:*")
    buy_total = 0.0
    for n, item in enumerate(cart["items"], 1):
        p   = round(float(item["price"]), 2)
        tot = round(p * item["qty"], 2)
        buy_total += tot
        p_s   = f"{p:.0f}"   if p   == int(p)   else f"{p:.2f}"
        tot_s = f"{tot:.0f}" if tot == int(tot) else f"{tot:.2f}"
        lines.append(f"{n}. {item['name']}")
        lines.append(f"   {item['qty']} шт. × {p_s} ₽ = *{tot_s} ₽*")
    lines.append(f"💰 *Итого закупка: {buy_total:.2f} ₽*")

    has_sell = any(i.get("sell_price") for i in cart["items"])
    if has_sell:
        sell_total = 0.0
        lines.append("\n💸 *ПЕРЕПРОДАЖА:*")
        for n, item in enumerate(cart["items"], 1):
            sp  = round(float(item.get("sell_price") or item["price"]), 2)
            tot = round(sp * item["qty"], 2)
            sell_total += tot
            sp_s  = f"{sp:.0f}"  if sp  == int(sp)  else f"{sp:.2f}"
            tot_s = f"{tot:.0f}" if tot == int(tot) else f"{tot:.2f}"
            note  = "" if item.get("sell_price") else " _(не изменена)_"
            lines.append(f"{n}. {item['name']}")
            lines.append(f"   {item['qty']} шт. × {sp_s} ₽ = *{tot_s} ₽*{note}")
        margin = round(sell_total - buy_total, 2)
        lines.append(f"💰 *Итого продажа: {sell_total:.2f} ₽*")
        lines.append(f"📈 *Маржа: {margin:.2f} ₽*")

    return "\n".join(lines)


# ─────────────────────────────────────────────────────────────────
# КОРЗИНА
# ─────────────────────────────────────────────────────────────────

def get_cart(ctx) -> dict:
    if "cart" not in ctx.user_data:
        ctx.user_data["cart"] = {"client": "", "items": []}
    return ctx.user_data["cart"]


def add_to_cart(cart, name: str, price: float, qty: int):
    for item in cart["items"]:
        if item["name"] == name:
            item["qty"] += qty
            return
    cart["items"].append({"name": name, "price": price, "qty": qty, "sell_price": None})


def cart_total(cart) -> float:
    return sum(i["price"] * i["qty"] for i in cart["items"])


# ─────────────────────────────────────────────────────────────────
# ЗАКАЗЫ И АНАЛИТИКА
# ─────────────────────────────────────────────────────────────────

def load_orders_json() -> list:
    """Загружает заказы: сначала из Google Sheets, потом из локального JSON."""
    # Пробуем Google Sheets
    if GOOGLE_SHEET_ID and GOOGLE_CREDS_JSON:
        try:
            orders = load_orders_from_sheets()
            if orders is not None:
                return orders
        except Exception as e:
            logger.warning(f"Sheets fallback to JSON: {e}")
    # Локальный JSON как запасной вариант
    try:
        if os.path.exists(ORDERS_JSON):
            with open(ORDERS_JSON, encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return []


def save_order(cart):
    now      = datetime.now()
    now_str  = now.strftime("%d.%m.%Y %H:%M")
    buy_tot  = sum(i["price"] * i["qty"] for i in cart["items"])
    has_sell = any(i.get("sell_price") for i in cart["items"])
    sell_tot = sum((i.get("sell_price") or i["price"]) * i["qty"] for i in cart["items"])
    margin   = sell_tot - buy_tot if has_sell else 0.0

    lines = [f"\n{'='*50}", f"📅 {now_str}", f"👤 {cart['client']}", "", "ЗАКУПКА:"]
    for item in cart["items"]:
        tot = item["price"] * item["qty"]
        lines.append(f"  • {item['name']} × {item['qty']} шт. × {item['price']:.2f}₽ = {tot:.2f}₽")
    lines.append(f"  Итого: {buy_tot:.2f}₽")
    if has_sell:
        lines.append("ПЕРЕПРОДАЖА:")
        for item in cart["items"]:
            sp  = item.get("sell_price") or item["price"]
            tot = sp * item["qty"]
            lines.append(f"  • {item['name']} × {item['qty']} шт. × {sp:.2f}₽ = {tot:.2f}₽")
        lines.append(f"  Итого: {sell_tot:.2f}₽  Маржа: {margin:.2f}₽")

    text = "\n".join(lines)
    print("ORDER:", text, flush=True)
    try:
        with open(ORDERS_FILE, "a", encoding="utf-8") as f:
            f.write(text + "\n")
    except Exception:
        pass

    record = {
        "dt": now.strftime("%Y-%m-%d %H:%M:%S"),
        "client": cart["client"],
        "buy_total":  round(buy_tot,  2),
        "sell_total": round(sell_tot, 2),
        "margin":     round(margin,   2),
        "items": [{"name": i["name"], "qty": i["qty"],
                   "price": i["price"], "sell_price": i.get("sell_price")}
                  for i in cart["items"]]
    }
    # Сохраняем в Google Sheets (основное хранилище)
    sheets_ok = save_order_to_sheets(cart, buy_tot, sell_tot, margin)
    if sheets_ok:
        logger.info("Заказ → Google Sheets ✅")
    else:
        logger.warning("Google Sheets недоступен, сохраняем в JSON")

    # JSON как резервное хранилище
    try:
        history = []
        if os.path.exists(ORDERS_JSON):
            with open(ORDERS_JSON, encoding="utf-8") as f:
                history = json.load(f)
        history.append(record)
        with open(ORDERS_JSON, "w", encoding="utf-8") as f:
            json.dump(history, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.warning(f"save JSON: {e}")


def analytics_for_period(orders: list, days: int) -> str:
    from datetime import timedelta
    cutoff  = datetime.now() - timedelta(days=days)
    period  = [o for o in orders
               if datetime.strptime(o["dt"], "%Y-%m-%d %H:%M:%S") >= cutoff]
    if not period:
        return "Заказов за этот период нет."
    buy    = sum(o["buy_total"]  for o in period)
    sell   = sum(o["sell_total"] for o in period)
    margin = sum(o["margin"]     for o in period)
    from collections import defaultdict
    by_client = defaultdict(float)
    for o in period:
        by_client[o["client"]] += o["margin"]
    top = sorted(by_client.items(), key=lambda x: -x[1])[:5]
    lines = [
        f"📊 Заказов: *{len(period)}*",
        f"📦 Закупка: *{buy:.0f} ₽*",
        f"💸 Продажа: *{sell:.0f} ₽*",
        f"📈 Маржа: *{margin:.0f} ₽*",
    ]
    if top:
        lines += ["", "🏆 *Топ клиентов:*"]
        for client, m in top:
            lines.append(f"  • {client}: {m:.0f} ₽")
    return "\n".join(lines)


# ─────────────────────────────────────────────────────────────────
# ПОДСКАЗКА ПО КЛИЕНТУ
# ─────────────────────────────────────────────────────────────────

def client_price_hint(client: str, product_name: str) -> str:
    if not client:
        return ""
    prices = []
    for o in load_orders_json():
        if o.get("client", "").lower() != client.lower():
            continue
        for item in o.get("items", []):
            if product_name.lower() in item.get("name", "").lower():
                sp = item.get("sell_price")
                if sp:
                    prices.append(float(sp))
    if not prices:
        return ""
    return (
        f"💡 *Подсказка по {client}:*\n"
        f"   Последняя цена продажи: *{prices[-1]:.0f} ₽*\n"
        f"   Средняя: *{sum(prices)/len(prices):.0f} ₽*\n\n"
    )


# ─────────────────────────────────────────────────────────────────
# ГЛАВНОЕ МЕНЮ
# ─────────────────────────────────────────────────────────────────

def main_menu_kb() -> InlineKeyboardMarkup:
    beer_s = f" ({len(df_beer)} поз.)" if df_beer is not None else " (не загружен)"
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🔍 Поиск",           callback_data="menu:search"),
         InlineKeyboardButton("🛒 Корзина",         callback_data="menu:cart")],
        [InlineKeyboardButton("📋 Новый заказ",     callback_data="menu:neworder"),
         InlineKeyboardButton("💸 Цены продажи",    callback_data="menu:sell")],
        [InlineKeyboardButton("✅ Подтвердить",     callback_data="menu:confirm"),
         InlineKeyboardButton("✏️ Редактировать",   callback_data="menu:edit")],
        [InlineKeyboardButton("📦 Заказы",          callback_data="menu:orders"),
         InlineKeyboardButton("📊 Аналитика",       callback_data="menu:stats")],
        [InlineKeyboardButton(f"🍺 Пивной прайс{beer_s}", callback_data="menu:uploadbeer")],
    ])


async def send_main_menu(msg_obj, context):
    n      = len(df_price) if df_price is not None else 0
    status = f"\n✅ Алко: *{n}* поз." if n else "\n⚠️ Алко прайс не загружен"
    if df_beer is not None:
        status += f"  |  🍺 *{len(df_beer)}* поз."
    await msg_obj.reply_text(
        "🍾 *АлкоПартнёр — МАИ 2026*" + status + "\n\nВыбери действие или напиши запрос:",
        parse_mode="Markdown",
        reply_markup=main_menu_kb()
    )


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await send_main_menu(update.message, context)

async def cmd_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await send_main_menu(update.message, context)

async def cmd_myid(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        f"Твой ID: `{update.effective_user.id}`", parse_mode="Markdown"
    )


# ─────────────────────────────────────────────────────────────────
# ЗАГРУЗКА ФАЙЛОВ
# ─────────────────────────────────────────────────────────────────

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc   = update.message.document
    fname = (doc.file_name or "").lower()
    if not (fname.endswith(".xls") or fname.endswith(".xlsx")):
        await update.message.reply_text("❌ Нужен файл `.xls` или `.xlsx`")
        return
    await update.message.reply_text("⏳ Загружаю файл...")
    upload_type = context.user_data.pop("upload_type", None)
    is_beer = (upload_type == "beer" or
               any(x in fname for x in ["пив", "beer", "pivo", "pricepivo"]))
    if is_beer:
        await (await doc.get_file()).download_to_drive(PRICE_BEER)
        ok, msg = load_beer_from_xlsx(PRICE_BEER)
    else:
        await (await doc.get_file()).download_to_drive(PRICE_XLS)
        ok, msg = load_price_from_xls(PRICE_XLS)
    await update.message.reply_text(msg, parse_mode="Markdown")


# ─────────────────────────────────────────────────────────────────
# ПОИСК И ВЫВОД РЕЗУЛЬТАТОВ
# ─────────────────────────────────────────────────────────────────

async def _delete_messages_later(bot, chat_id: int, message_ids: list[int], delay: int = 120):
    """Удаляет список сообщений через delay секунд."""
    import asyncio
    await asyncio.sleep(delay)
    for mid in message_ids:
        try:
            await bot.delete_message(chat_id=chat_id, message_id=mid)
        except Exception:
            pass


async def _send_alco_results(msg_obj, results, query, context=None):
    if not results:
        await msg_obj.reply_text(f"🍷 «{query}» — не найдено в алко прайсе.")
        return
    MAX = 20
    hdr = f"🍷 *{len(results)}* позиций по «{query}»"
    if len(results) > MAX:
        hdr += f"\n_(первые {MAX})_"
    hdr += "\n_⏱ Результаты исчезнут через 2 минуты_"

    sent_ids = []
    m = await msg_obj.reply_text(hdr, parse_mode="Markdown")
    sent_ids.append(m.message_id)

    for r in results[:MAX]:
        p     = float(r["price"])
        p_s   = f"{p:.0f}" if p == int(p) else f"{p:.2f}"
        promo = ""
        has_p = pd.notna(r.get("promo_price")) and float(r.get("promo_price", 0)) > 0
        if has_p:
            pp    = float(r["promo_price"])
            promo = f"\n   🎯 Маркетинг: *{pp:.2f} ₽*"
        text = (f"🍷 *{r['name']}*\n"
                f"   💰 *{p_s} ₽*  |  {int(r['stock'])} шт.{promo}")
        if has_p:
            kb = InlineKeyboardMarkup([[
                InlineKeyboardButton(f"➕ 1 МАИ {p_s}₽",  callback_data=f"add:1:{r.name}"),
                InlineKeyboardButton(f"➕ 6",              callback_data=f"add:6:{r.name}"),
                InlineKeyboardButton(f"➕ 12",             callback_data=f"add:12:{r.name}"),
            ], [
                InlineKeyboardButton(f"🎯 Маркетинг {float(r['promo_price']):.2f}₽",
                                     callback_data=f"price:promo:{r.name}:1"),
            ]])
        else:
            kb = InlineKeyboardMarkup([[
                InlineKeyboardButton("➕ 1",   callback_data=f"add:1:{r.name}"),
                InlineKeyboardButton("➕ 6",   callback_data=f"add:6:{r.name}"),
                InlineKeyboardButton("➕ 12",  callback_data=f"add:12:{r.name}"),
                InlineKeyboardButton("➕ ...", callback_data=f"addx:{r.name}"),
            ]])
        m = await msg_obj.reply_text(text, parse_mode="Markdown", reply_markup=kb)
        sent_ids.append(m.message_id)

    # Запускаем удаление через 2 минуты
    if context:
        import asyncio
        asyncio.create_task(
            _delete_messages_later(context.bot, msg_obj.chat_id, sent_ids, delay=120)
        )


async def _send_beer_results(msg_obj, results, query, context=None):
    if not results:
        await msg_obj.reply_text(f"🍺 «{query}» — не найдено в пивном прайсе.")
        return
    MAX = 20
    hdr = f"🍺 *{len(results)}* позиций по «{query}»"
    if len(results) > MAX:
        hdr += f"\n_(первые {MAX})_"
    hdr += "\n_⏱ Результаты исчезнут через 2 минуты_"

    sent_ids = []
    m = await msg_obj.reply_text(hdr, parse_mode="Markdown")
    sent_ids.append(m.message_id)

    for r in results[:MAX]:
        pf  = float(r["price_factory"])
        pnw = float(r["price_nw"]) if pd.notna(r.get("price_nw")) and float(r.get("price_nw", 0)) > 0 else None
        text = f"🍺 *{r['name']}*\n   🏭 Завод: *{pf:.2f} ₽*"
        if pnw:
            text += f"  |  📦 С-З: *{pnw:.2f} ₽*"
        kb = InlineKeyboardMarkup([[
            InlineKeyboardButton(f"➕ 1 Завод {pf:.2f}₽",     callback_data=f"badd:1:{r.name}:factory"),
            InlineKeyboardButton(f"➕ 1 С-З {pnw:.2f}₽" if pnw else "➕ 1",
                                 callback_data=f"badd:1:{r.name}:nw"),
        ], [
            InlineKeyboardButton("➕ 6",   callback_data=f"badd:6:{r.name}:factory"),
            InlineKeyboardButton("➕ 12",  callback_data=f"badd:12:{r.name}:factory"),
            InlineKeyboardButton("➕ ...", callback_data=f"baddx:{r.name}"),
        ]])
        m = await msg_obj.reply_text(text, parse_mode="Markdown", reply_markup=kb)
        sent_ids.append(m.message_id)

    if context:
        import asyncio
        asyncio.create_task(
            _delete_messages_later(context.bot, msg_obj.chat_id, sent_ids, delay=120)
        )


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    aw   = context.user_data.get("awaiting")

    # ── Ожидание имени клиента ────────────────────────────────────
    if aw == "client":
        cart = get_cart(context)
        cart["client"] = text
        context.user_data.pop("awaiting", None)
        pending = context.user_data.pop("pending", None)
        added_msg = ""
        if pending:
            if "beer_idx" in pending:
                row   = df_beer.iloc[pending["beer_idx"]]
                pt    = pending.get("price_type", "factory")
                price = (float(row["price_nw"])
                         if pt == "nw" and pd.notna(row.get("price_nw")) and float(row.get("price_nw", 0)) > 0
                         else float(row["price_factory"]))
                name  = f"🍺 {row['name']}"
                add_to_cart(cart, name, price, pending["qty"])
                added_msg = f"✅ *{name}* × {pending['qty']} шт.\n\n"
            elif "idx" in pending and df_price is not None:
                row = df_price.iloc[pending["idx"]]
                add_to_cart(cart, row["name"], float(row["price"]), pending["qty"])
                added_msg = f"✅ *{row['name']}* × {pending['qty']} шт.\n\n"
        await update.message.reply_text(
            f"👤 Клиент: *{cart['client']}*\n\n{added_msg}"
            + fmt_cart(cart) + "\n\nИщи товары или /confirm",
            parse_mode="Markdown"
        )
        return

    # ── Ожидание количества (алко) ────────────────────────────────
    if aw == "qty":
        m = re.search(r'\d+', text)
        if not m:
            await update.message.reply_text("Введи число, например `6`", parse_mode="Markdown")
            return
        qty = int(m.group())
        context.user_data.pop("awaiting", None)
        idx = context.user_data.pop("pending_idx", None)
        if idx is not None and df_price is not None:
            row  = df_price.iloc[idx]
            cart = get_cart(context)
            if not cart["client"]:
                context.user_data["awaiting"] = "client"
                context.user_data["pending"]  = {"idx": idx, "qty": qty}
                await update.message.reply_text("👤 Введи имя клиента:")
                return
            add_to_cart(cart, row["name"], float(row["price"]), qty)
            hint = client_price_hint(cart["client"], row["name"])
            await update.message.reply_text(
                f"✅ *{row['name']}* × {qty} шт.\n{hint}" + fmt_cart(cart) + "\n\nИщи или /confirm",
                parse_mode="Markdown"
            )
        return

    # ── Ожидание количества (пиво) ────────────────────────────────
    if aw == "beer_qty":
        m = re.search(r'\d+', text)
        if not m:
            await update.message.reply_text("Введи число, например `6`", parse_mode="Markdown")
            return
        qty = int(m.group())
        idx = context.user_data.pop("pending_beer_idx", None)
        context.user_data.pop("awaiting", None)
        if idx is not None and df_beer is not None:
            row  = df_beer.iloc[idx]
            cart = get_cart(context)
            if not cart["client"]:
                context.user_data["awaiting"] = "client"
                context.user_data["pending"]  = {"beer_idx": idx, "qty": qty, "price_type": "factory"}
                await update.message.reply_text("👤 Введи имя клиента:")
                return
            name  = f"🍺 {row['name']}"
            price = float(row["price_factory"])
            add_to_cart(cart, name, price, qty)
            await update.message.reply_text(
                f"✅ *{name}* × {qty} шт.\n" + fmt_cart(cart) + "\n\nИщи или /confirm",
                parse_mode="Markdown"
            )
        return

    # ── Ожидание нового количества при редактировании ────────────
    if aw == "edit_qty":
        m = re.search(r'\d+', text)
        if not m:
            await update.message.reply_text("Введи число, например `6`", parse_mode="Markdown")
            return
        qty  = int(m.group())
        idx  = context.user_data.pop("edit_idx", None)
        context.user_data.pop("awaiting", None)
        cart = get_cart(context)
        if idx is not None and 0 <= idx < len(cart["items"]):
            cart["items"][idx]["qty"] = qty
            await update.message.reply_text(
                f"✅ Количество изменено на *{qty} шт.*", parse_mode="Markdown"
            )
            await _show_edit_menu(update.message, context)
        return

    # ── Ожидание цены перепродажи ─────────────────────────────────
    if aw == "sell_price":
        m = re.search(r'[\d.,]+', text.replace(",", "."))
        if not m:
            await update.message.reply_text("Введи число, например `390`", parse_mode="Markdown")
            return
        try:
            price = float(m.group().replace(",", "."))
        except ValueError:
            await update.message.reply_text("Введи число, например `390`", parse_mode="Markdown")
            return
        idx  = context.user_data.get("sell_idx", 0)
        cart = get_cart(context)
        if idx < len(cart["items"]):
            cart["items"][idx]["sell_price"] = price
        context.user_data["sell_idx"] = idx + 1
        await _ask_sell_price(update.message, context)
        return

    # ── Обычный поиск ─────────────────────────────────────────────
    if len(text) < 2:
        return

    if df_price is None and df_beer is None:
        await update.message.reply_text("📤 Загрузи прайс — отправь `.xls` файл.")
        return

    # Оба прайса загружены — спрашиваем категорию
    if df_price is not None and df_beer is not None:
        context.user_data["pending_search"] = text
        await update.message.reply_text(
            f"Ищем «{text}» — выбери категорию:",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton("🍷 Алкоголь", callback_data="cat:alco"),
                InlineKeyboardButton("🍺 Пиво",     callback_data="cat:beer"),
                InlineKeyboardButton("🔍 Везде",    callback_data="cat:all"),
            ]])
        )
        return

    query, min_p, max_p = parse_price_filter(text)
    if not query.strip():
        query = text

    if df_price is None:
        results = do_search_beer(query, min_p, max_p)
        await _send_beer_results(update.message, results, query, context)
        return

    results = do_search(query, min_p, max_p)
    if not results:
        await update.message.reply_text(f"🔍 «{text}» — ничего не найдено.")
        return

    if len(results) == 1:
        row   = results[0]
        card  = fmt_card(row)
        has_p = pd.notna(row.get("promo_price")) and float(row.get("promo_price", 0)) > 0
        if has_p:
            kb = InlineKeyboardMarkup([[
                InlineKeyboardButton(f"➕ 1 МАИ",  callback_data=f"add:1:{row.name}"),
                InlineKeyboardButton(f"➕ 6",       callback_data=f"add:6:{row.name}"),
                InlineKeyboardButton(f"➕ 12",      callback_data=f"add:12:{row.name}"),
            ], [
                InlineKeyboardButton(f"🎯 Маркетинг {float(row['promo_price']):.2f}₽",
                                     callback_data=f"price:promo:{row.name}:1"),
            ]])
        else:
            kb = InlineKeyboardMarkup([[
                InlineKeyboardButton("➕ 1",   callback_data=f"add:1:{row.name}"),
                InlineKeyboardButton("➕ 6",   callback_data=f"add:6:{row.name}"),
                InlineKeyboardButton("➕ 12",  callback_data=f"add:12:{row.name}"),
                InlineKeyboardButton("➕ ...", callback_data=f"addx:{row.name}"),
            ]])
        photo = find_photo(row["name"], str(row.get("barcode", "")))
        sent  = False
        if photo:
            try:
                await update.message.reply_photo(
                    photo=photo, caption=card, parse_mode="Markdown", reply_markup=kb
                )
                sent = True
            except Exception:
                pass
        if not sent:
            await update.message.reply_text(card, parse_mode="Markdown", reply_markup=kb)
        return

    await _send_alco_results(update.message, results, query, context)


# ─────────────────────────────────────────────────────────────────
# CALLBACK — КАТЕГОРИЯ ПОИСКА
# ─────────────────────────────────────────────────────────────────

async def cb_category(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q   = update.callback_query
    await q.answer()
    cat = q.data.split(":")[1]
    raw = context.user_data.pop("pending_search", "")
    if not raw:
        await q.message.reply_text("Введи запрос ещё раз.")
        return
    query, min_p, max_p = parse_price_filter(raw)
    if not query.strip():
        query = raw
    if cat == "alco":
        await _send_alco_results(q.message, do_search(query, min_p, max_p), query, context)
    elif cat == "beer":
        await _send_beer_results(q.message, do_search_beer(query, min_p, max_p), query, context)
    else:
        alco = do_search(query, min_p, max_p)
        beer = do_search_beer(query, min_p, max_p)
        if not alco and not beer:
            await q.message.reply_text(f"🔍 «{query}» — нигде не найдено.")
            return
        if alco:
            await _send_alco_results(q.message, alco, query, context)
        if beer:
            await _send_beer_results(q.message, beer, query, context)


# ─────────────────────────────────────────────────────────────────
# CALLBACK — ДОБАВЛЕНИЕ В КОРЗИНУ (АЛКО)
# ─────────────────────────────────────────────────────────────────

async def cb_add(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q      = update.callback_query
    await q.answer()
    parts  = q.data.split(":", 2)
    action = parts[0]
    idx    = int(parts[-1])

    if action == "addx":
        context.user_data["awaiting"]    = "qty"
        context.user_data["pending_idx"] = idx
        await q.message.reply_text("Введи количество штук:")
        return

    qty  = int(parts[1])
    row  = df_price.iloc[idx]
    cart = get_cart(context)

    if not cart["client"]:
        context.user_data["awaiting"] = "client"
        context.user_data["pending"]  = {"idx": idx, "qty": qty}
        await q.message.reply_text("👤 Введи имя клиента:")
        return

    has_p = pd.notna(row.get("promo_price")) and float(row.get("promo_price", 0)) > 0
    if has_p:
        promo = float(row["promo_price"])
        base  = float(row["price"])
        cond  = str(row.get("promo_cond", "")).strip()
        cond_s = f"\n📋 Условие: _{cond}_" if cond else ""
        await q.message.reply_text(
            f"📦 *{row['name']}* × {qty} шт.\n\nВыбери цену закупки:{cond_s}",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton(f"💰 МАИ {base:.2f}₽",        callback_data=f"price:base:{idx}:{qty}"),
                InlineKeyboardButton(f"🎯 Маркетинг {promo:.2f}₽", callback_data=f"price:promo:{idx}:{qty}"),
            ]])
        )
        return

    add_to_cart(cart, row["name"], float(row["price"]), qty)
    hint = client_price_hint(cart["client"], row["name"])
    await q.message.reply_text(
        f"✅ *{row['name']}* × {qty} шт.\n{hint}" + fmt_cart(cart) + "\n\nИщи или /confirm",
        parse_mode="Markdown"
    )


async def cb_price_choice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q      = update.callback_query
    await q.answer()
    parts  = q.data.split(":")
    choice = parts[1]
    idx    = int(parts[2])
    qty    = int(parts[3])
    row    = df_price.iloc[idx]
    cart   = get_cart(context)

    if choice == "promo" and pd.notna(row.get("promo_price")) and float(row.get("promo_price", 0)) > 0:
        price = float(row["promo_price"])
        label = "🎯 маркетинговая"
    else:
        price = float(row["price"])
        label = "💰 МАИ"

    add_to_cart(cart, row["name"], price, qty)
    hint = client_price_hint(cart["client"], row["name"])
    await q.message.reply_text(
        f"✅ *{row['name']}* × {qty} шт. ({label}: {price:.2f}₽)\n{hint}"
        + fmt_cart(cart) + "\n\nИщи или /confirm",
        parse_mode="Markdown"
    )


# ─────────────────────────────────────────────────────────────────
# CALLBACK — ДОБАВЛЕНИЕ В КОРЗИНУ (ПИВО)
# ─────────────────────────────────────────────────────────────────

async def cb_beer_add(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q      = update.callback_query
    await q.answer()
    parts  = q.data.split(":")
    action = parts[0]

    if action == "baddx":
        idx = int(parts[1])
        context.user_data["awaiting"]         = "beer_qty"
        context.user_data["pending_beer_idx"] = idx
        await q.message.reply_text("Введи количество штук:")
        return

    qty        = int(parts[1])
    idx        = int(parts[2])
    price_type = parts[3] if len(parts) > 3 else "factory"
    row        = df_beer.iloc[idx]
    cart       = get_cart(context)

    pnw   = row.get("price_nw")
    price = (float(pnw) if price_type == "nw" and pd.notna(pnw) and float(pnw) > 0
             else float(row["price_factory"]))
    label = "С-З" if price_type == "nw" else "Завод"
    name  = f"🍺 {row['name']}"

    if not cart["client"]:
        context.user_data["awaiting"] = "client"
        context.user_data["pending"]  = {"beer_idx": idx, "qty": qty, "price_type": price_type}
        await q.message.reply_text("👤 Введи имя клиента:")
        return

    add_to_cart(cart, name, price, qty)
    hint = client_price_hint(cart["client"], row["name"])
    await q.message.reply_text(
        f"✅ *{name}* × {qty} шт. ({label}: {price:.2f}₽)\n{hint}"
        + fmt_cart(cart) + "\n\nИщи или /confirm",
        parse_mode="Markdown"
    )


# ─────────────────────────────────────────────────────────────────
# КОРЗИНА — ПРОСМОТР И РЕДАКТИРОВАНИЕ
# ─────────────────────────────────────────────────────────────────

async def _cmd_cart_msg(msg_obj, context):
    cart = get_cart(context)
    if not cart["items"]:
        await msg_obj.reply_text(
            "🛒 Корзина пуста. Найди товар и нажми ➕",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton("🏠 Меню", callback_data="menu:start")
            ]])
        )
        return
    btns = [[InlineKeyboardButton(f"❌ {item['name'][:35]}", callback_data=f"rm:{i}")]
            for i, item in enumerate(cart["items"])]
    btns += [
        [InlineKeyboardButton("🗑 Очистить всё",     callback_data="rm:all"),
         InlineKeyboardButton("✏️ Редактировать",    callback_data="edit:open:0")],
        [InlineKeyboardButton("🏠 Меню",             callback_data="menu:start")],
    ]
    await msg_obj.reply_text(
        fmt_cart(cart) + "\n\n✅ /confirm — подтвердить",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(btns)
    )

async def cmd_cart(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await _cmd_cart_msg(update.message, context)

async def cb_remove(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q    = update.callback_query
    await q.answer()
    cart = get_cart(context)
    key  = q.data.split(":")[1]
    if key == "all":
        context.user_data["cart"] = {"client": "", "items": []}
        await q.message.edit_text("🛒 Корзина очищена.")
        return
    idx = int(key)
    if 0 <= idx < len(cart["items"]):
        cart["items"].pop(idx)
    if cart["items"]:
        await q.message.edit_text(fmt_cart(cart), parse_mode="Markdown")
    else:
        await q.message.edit_text("🛒 Корзина пуста.")


async def _show_edit_menu(msg_obj, context):
    cart  = get_cart(context)
    lines = ["✏️ *Редактирование корзины*\n"]
    btns  = []
    for i, item in enumerate(cart["items"]):
        p   = float(item["price"])
        tot = round(p * item["qty"], 2)
        lines.append(f"{i+1}. {item['name']}\n   {item['qty']} шт. × {p:.2f}₽ = {tot:.2f}₽")
        btns.append([
            InlineKeyboardButton(f"✏️ {i+1}. кол-во", callback_data=f"edit:qty:{i}"),
            InlineKeyboardButton("❌ Удалить",          callback_data=f"edit:del:{i}"),
        ])
    btns.append([InlineKeyboardButton("✅ Готово", callback_data="menu:confirm")])
    await msg_obj.reply_text(
        "\n".join(lines), parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(btns)
    )

async def cmd_edit(update: Update, context: ContextTypes.DEFAULT_TYPE):
    cart = get_cart(context)
    if not cart["items"]:
        await update.message.reply_text("🛒 Корзина пуста.")
        return
    await _show_edit_menu(update.message, context)

async def cb_edit(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q      = update.callback_query
    await q.answer()
    parts  = q.data.split(":")
    action = parts[1]
    idx    = int(parts[2])
    cart   = get_cart(context)

    if action == "open":
        await _show_edit_menu(q.message, context)
    elif action == "del":
        if 0 <= idx < len(cart["items"]):
            removed = cart["items"].pop(idx)
            await q.message.reply_text(
                f"❌ *{removed['name']}* удалена.", parse_mode="Markdown"
            )
            if cart["items"]:
                await _show_edit_menu(q.message, context)
    elif action == "qty":
        context.user_data["awaiting"] = "edit_qty"
        context.user_data["edit_idx"] = idx
        item = cart["items"][idx]
        await q.message.reply_text(
            f"✏️ *{item['name']}*\nСейчас: {item['qty']} шт.\n\nВведи новое количество:",
            parse_mode="Markdown"
        )


# ─────────────────────────────────────────────────────────────────
# ПОДТВЕРЖДЕНИЕ ЗАКАЗА
# ─────────────────────────────────────────────────────────────────

async def cmd_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    cart = get_cart(context)
    if not cart["items"]:
        await update.message.reply_text("🛒 Корзина пуста.")
        return
    if not cart["client"]:
        context.user_data["awaiting"] = "client"
        await update.message.reply_text("👤 Введи имя клиента:")
        return
    await update.message.reply_text(
        fmt_cart(cart, title="📋 *Итог заказа*") + "\n\nПодтверди или вернись к редактированию:",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup([[
            InlineKeyboardButton("✅ Подтвердить", callback_data="order:ok"),
            InlineKeyboardButton("✏️ Изменить",   callback_data="order:edit"),
        ]])
    )

async def cb_order(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q      = update.callback_query
    await q.answer()
    cart   = get_cart(context)
    action = q.data.split(":")[1]
    if action == "edit":
        await q.message.edit_text(
            fmt_cart(cart) + "\n\nДобавляй товары или /confirm",
            parse_mode="Markdown"
        )
    elif action == "ok":
        save_order(cart)
        await q.message.edit_text(
            fmt_cart(cart, title="✅ *Заказ принят*"), parse_mode="Markdown"
        )
        await q.message.reply_text("/neworder — новый заказ")
        context.user_data["cart"] = {"client": "", "items": []}


# ─────────────────────────────────────────────────────────────────
# ЦЕНЫ ПЕРЕПРОДАЖИ
# ─────────────────────────────────────────────────────────────────

async def _ask_sell_price(msg_obj, context):
    cart = get_cart(context)
    idx  = context.user_data.get("sell_idx", 0)
    if idx >= len(cart["items"]):
        context.user_data.pop("sell_idx", None)
        context.user_data.pop("awaiting", None)
        await msg_obj.reply_text(
            fmt_cart(cart, title="✅ *Цены введены*") + "\n\n/confirm — подтвердить",
            parse_mode="Markdown"
        )
        return
    item  = cart["items"][idx]
    p_s   = f"{item['price']:.2f}".rstrip("0").rstrip(".")
    cur   = (f"текущая: *{item['sell_price']:.2f} ₽*"
             if item.get("sell_price") else f"цена закупки: *{p_s} ₽*")
    await msg_obj.reply_text(
        f"💸 *Цена продажи {idx+1} из {len(cart['items'])}*\n\n"
        f"📦 {item['name']}\n"
        f"   Закупка: *{p_s} ₽* × {item['qty']} шт.\n"
        f"   {cur}\n\nВведи цену продажи или пропусти:",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup([[
            InlineKeyboardButton("➡️ Оставить", callback_data=f"sell:skip:{idx}"),
        ]])
    )
    context.user_data["awaiting"] = "sell_price"

async def cmd_sell(update: Update, context: ContextTypes.DEFAULT_TYPE):
    cart = get_cart(context)
    if not cart["items"]:
        await update.message.reply_text("🛒 Корзина пуста.")
        return
    context.user_data["sell_idx"] = 0
    await _ask_sell_price(update.message, context)

async def cb_sell_skip(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q   = update.callback_query
    await q.answer()
    idx = int(q.data.split(":")[2])
    context.user_data["sell_idx"] = idx + 1
    await _ask_sell_price(q.message, context)


# ─────────────────────────────────────────────────────────────────
# СПИСОК ЗАКАЗОВ
# ─────────────────────────────────────────────────────────────────

async def cmd_orders(update: Update, context: ContextTypes.DEFAULT_TYPE):
    orders = load_orders_json()
    if not orders:
        await update.message.reply_text("📦 Заказов пока нет.")
        return
    await _show_orders_page(update.message, context, page=0)

async def _show_orders_page(msg_obj, context, page: int = 0):
    orders   = load_orders_json()
    total    = len(orders)
    per_page = 5
    start    = page * per_page
    end      = min(start + per_page, total)

    lines = [f"📦 *Все заказы* ({total} шт.) — стр. {page+1}/{(total-1)//per_page+1}\n"]
    btns  = []
    for i, o in enumerate(orders[start:end], start=start):
        dt     = o.get("dt", "")[:10]
        client = o.get("client", "?")
        buy    = o.get("buy_total",  0)
        sell   = o.get("sell_total", 0)
        margin = o.get("margin",     0)
        lines.append(
            f"{i+1}. *{client}* — {dt}\n"
            f"   Закупка: {buy:.0f}₽  Продажа: {sell:.0f}₽  Маржа: {margin:.0f}₽"
        )
        btns.append([
            InlineKeyboardButton(f"👁 #{i+1} {client} {dt}", callback_data=f"vieworder:{i}"),
            InlineKeyboardButton("🗑",                        callback_data=f"delorder:{i}"),
        ])
    nav = []
    if page > 0:
        nav.append(InlineKeyboardButton("◀️ Назад",  callback_data=f"ordpage:{page-1}"))
    if end < total:
        nav.append(InlineKeyboardButton("▶️ Вперёд", callback_data=f"ordpage:{page+1}"))
    if nav:
        btns.append(nav)
    btns.append([InlineKeyboardButton("🏠 Меню", callback_data="menu:start")])
    await msg_obj.reply_text(
        "\n".join(lines), parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(btns)
    )


async def cb_view_order(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Показывает полную информацию по конкретному заказу."""
    q   = update.callback_query
    await q.answer()
    idx = int(q.data.split(":")[1])
    orders = load_orders_json()
    if idx >= len(orders):
        await q.message.reply_text("Заказ не найден.")
        return
    o      = orders[idx]
    dt     = o.get("dt", "")
    client = o.get("client", "?")
    items  = o.get("items", [])
    buy    = o.get("buy_total",  0)
    sell   = o.get("sell_total", 0)
    margin = o.get("margin",     0)

    lines = [
        f"📋 *Заказ #{idx+1}*",
        f"👤 Клиент: *{client}*",
        f"📅 Дата: {dt[:16]}",
        "",
    ]

    # Если есть детальные позиции (из JSON)
    if items:
        lines.append("📦 *ЗАКУПКА:*")
        for item in items:
            p   = float(item.get("price", 0))
            qty = int(item.get("qty", 0))
            tot = p * qty
            p_s   = f"{p:.0f}"   if p   == int(p)   else f"{p:.2f}"
            tot_s = f"{tot:.0f}" if tot == int(tot) else f"{tot:.2f}"
            lines.append(f"• {item['name']}")
            lines.append(f"  {qty} шт. × {p_s} ₽ = *{tot_s} ₽*")
        lines.append(f"💰 *Итого закупка: {buy:.2f} ₽*")
        has_sell = any(i.get("sell_price") for i in items)
        if has_sell:
            lines += ["", "💸 *ПЕРЕПРОДАЖА:*"]
            for item in items:
                sp  = float(item.get("sell_price") or item.get("price", 0))
                qty = int(item.get("qty", 0))
                tot = sp * qty
                sp_s  = f"{sp:.0f}"  if sp  == int(sp)  else f"{sp:.2f}"
                tot_s = f"{tot:.0f}" if tot == int(tot) else f"{tot:.2f}"
                lines.append(f"• {item['name']}")
                lines.append(f"  {qty} шт. × {sp_s} ₽ = *{tot_s} ₽*")
            lines.append(f"💰 *Итого продажа: {sell:.2f} ₽*")
            lines.append(f"📈 *Маржа: {margin:.2f} ₽*")
    else:
        # Из Google Sheets — показываем строку позиций
        items_str = o.get("items_str", "")
        if items_str:
            lines.append("📦 *Позиции:*")
            for part in items_str.split("; "):
                lines.append(f"• {part.strip()}")
        lines.append(f"\n💰 *Закупка: {buy:.2f} ₽*")
        if sell > 0 and sell != buy:
            lines.append(f"💸 *Продажа: {sell:.2f} ₽*")
            lines.append(f"📈 *Маржа: {margin:.2f} ₽*")

    kb = InlineKeyboardMarkup([[
        InlineKeyboardButton("🗑 Удалить заказ", callback_data=f"delorder:{idx}"),
        InlineKeyboardButton("◀️ К списку",      callback_data="menu:orders"),
    ]])
    await q.message.reply_text(
        "\n".join(lines), parse_mode="Markdown", reply_markup=kb
    )

async def cb_orders_page(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q    = update.callback_query
    await q.answer()
    page = int(q.data.split(":")[1])
    await _show_orders_page(q.message, context, page=page)

async def cb_delete_order(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q   = update.callback_query
    await q.answer()
    idx = int(q.data.split(":")[1])
    orders = load_orders_json()
    if 0 <= idx < len(orders):
        removed = orders.pop(idx)
        try:
            with open(ORDERS_JSON, "w", encoding="utf-8") as f:
                json.dump(orders, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.warning(f"delete order: {e}")
        client = removed.get("client", "?")
        dt     = removed.get("dt", "")[:10]
        await q.message.reply_text(
            f"🗑 Заказ *{client}* от {dt} удалён. Осталось: {len(orders)}",
            parse_mode="Markdown"
        )
        if orders:
            await _show_orders_page(q.message, context, page=0)


# ─────────────────────────────────────────────────────────────────
# АНАЛИТИКА
# ─────────────────────────────────────────────────────────────────

async def cmd_stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    orders = load_orders_json()
    if not orders:
        await update.message.reply_text("📊 Нет данных. Подтверди хотя бы один заказ.")
        return
    await update.message.reply_text(
        "📊 *Аналитика — выбери период:*",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup([[
            InlineKeyboardButton("📅 День",   callback_data="stats:1"),
            InlineKeyboardButton("📆 Неделя", callback_data="stats:7"),
            InlineKeyboardButton("🗓 Месяц",  callback_data="stats:30"),
        ]])
    )

async def cb_stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q      = update.callback_query
    await q.answer()
    days   = int(q.data.split(":")[1])
    labels = {1: "сегодня", 7: "за неделю", 30: "за месяц"}
    text   = analytics_for_period(load_orders_json(), days)
    await q.message.edit_text(
        f"📊 *Аналитика — {labels[days]}*\n\n{text}",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup([[
            InlineKeyboardButton("📅 День",   callback_data="stats:1"),
            InlineKeyboardButton("📆 Неделя", callback_data="stats:7"),
            InlineKeyboardButton("🗓 Месяц",  callback_data="stats:30"),
        ]])
    )


# ─────────────────────────────────────────────────────────────────
# НОВЫЙ ЗАКАЗ
# ─────────────────────────────────────────────────────────────────

async def cmd_neworder(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["cart"]    = {"client": "", "items": []}
    context.user_data["awaiting"] = "client"
    context.user_data.pop("pending", None)
    await update.message.reply_text("🛒 Новый заказ.\n\n👤 Введи имя клиента:")


# ─────────────────────────────────────────────────────────────────
# CALLBACK — ГЛАВНОЕ МЕНЮ
# ─────────────────────────────────────────────────────────────────

async def cb_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q      = update.callback_query
    await q.answer()
    action = q.data.split(":")[1]

    if action == "start":
        await send_main_menu(q.message, context)
    elif action == "search":
        await q.message.reply_text(
            "🔍 Напиши что ищешь:\n\n"
            "▪ `водка 0,5`\n▪ `вино до 250`\n▪ `белуга`\n▪ `пиво 0,5`",
            parse_mode="Markdown"
        )
    elif action == "cart":
        await _cmd_cart_msg(q.message, context)
    elif action == "neworder":
        context.user_data["cart"]    = {"client": "", "items": []}
        context.user_data["awaiting"] = "client"
        context.user_data.pop("pending", None)
        await q.message.reply_text("🛒 Новый заказ.\n\n👤 Введи имя клиента:")
    elif action == "sell":
        cart = get_cart(context)
        if not cart["items"]:
            await q.message.reply_text("🛒 Корзина пуста. Добавь товары.")
        else:
            context.user_data["sell_idx"] = 0
            await _ask_sell_price(q.message, context)
    elif action == "confirm":
        cart = get_cart(context)
        if not cart["items"]:
            await q.message.reply_text("🛒 Корзина пуста.")
        elif not cart["client"]:
            context.user_data["awaiting"] = "client"
            await q.message.reply_text("👤 Введи имя клиента:")
        else:
            await q.message.reply_text(
                fmt_cart(cart, title="📋 *Итог заказа*") + "\n\nПодтверди:",
                parse_mode="Markdown",
                reply_markup=InlineKeyboardMarkup([[
                    InlineKeyboardButton("✅ Подтвердить", callback_data="order:ok"),
                    InlineKeyboardButton("✏️ Изменить",   callback_data="order:edit"),
                ]])
            )
    elif action == "edit":
        cart = get_cart(context)
        if not cart["items"]:
            await q.message.reply_text("🛒 Корзина пуста.")
        else:
            await _show_edit_menu(q.message, context)
    elif action == "orders":
        orders = load_orders_json()
        if not orders:
            await q.message.reply_text("📦 Заказов пока нет.")
        else:
            await _show_orders_page(q.message, context, page=0)
    elif action == "stats":
        orders = load_orders_json()
        if not orders:
            await q.message.reply_text("📊 Нет данных по заказам.")
        else:
            await q.message.reply_text(
                "📊 *Аналитика — выбери период:*",
                parse_mode="Markdown",
                reply_markup=InlineKeyboardMarkup([[
                    InlineKeyboardButton("📅 День",   callback_data="stats:1"),
                    InlineKeyboardButton("📆 Неделя", callback_data="stats:7"),
                    InlineKeyboardButton("🗓 Месяц",  callback_data="stats:30"),
                ]])
            )
    elif action == "uploadbeer":
        context.user_data["upload_type"] = "beer"
        await q.message.reply_text(
            "📤 Отправь файл пивного прайса (.xlsx)\n"
            "Колонки: название | Завод нал | Цена северо-запад"
        )


# ─────────────────────────────────────────────────────────────────
# ЗАПУСК
# ─────────────────────────────────────────────────────────────────

async def _post_init(application):
    await application.bot.set_my_commands([
        BotCommand("start",    "🏠 Главное меню"),
        BotCommand("neworder", "📋 Новый заказ"),
        BotCommand("cart",     "🛒 Корзина"),
        BotCommand("edit",     "✏️ Редактировать корзину"),
        BotCommand("sell",     "💸 Цены продажи"),
        BotCommand("confirm",  "✅ Подтвердить заказ"),
        BotCommand("orders",   "📦 Все заказы"),
        BotCommand("stats",    "📊 Аналитика"),
        BotCommand("myid",     "🔑 Мой ID"),
    ])
    print("✅ Команды меню установлены", flush=True)


def main():
    if not TOKEN:
        print("❌ BOT_TOKEN не задан!", flush=True)
        sys.exit(1)
    print(f"✅ Токен: {TOKEN[:10]}...", flush=True)

    if os.path.exists(PRICE_XLS):
        ok, msg = load_price_from_xls(PRICE_XLS)
        print(msg.replace("*", ""))
    else:
        print(f"⚠️ {PRICE_XLS} не найден — загрузи через Telegram")

    if os.path.exists(PRICE_BEER):
        ok, msg = load_beer_from_xlsx(PRICE_BEER)
        print(msg.replace("*", ""))
    else:
        print(f"⚠️ {PRICE_BEER} не найден — загрузи через Telegram")

    app = ApplicationBuilder().token(TOKEN).post_init(_post_init).build()

    app.add_handler(CommandHandler("start",    cmd_start))
    app.add_handler(CommandHandler("menu",     cmd_menu))
    app.add_handler(CommandHandler("myid",     cmd_myid))
    app.add_handler(CommandHandler("cart",     cmd_cart))
    app.add_handler(CommandHandler("edit",     cmd_edit))
    app.add_handler(CommandHandler("confirm",  cmd_confirm))
    app.add_handler(CommandHandler("neworder", cmd_neworder))
    app.add_handler(CommandHandler("sell",     cmd_sell))
    app.add_handler(CommandHandler("stats",    cmd_stats))
    app.add_handler(CommandHandler("orders",   cmd_orders))

    app.add_handler(CallbackQueryHandler(cb_menu,         pattern=r"^menu:"))
    app.add_handler(CallbackQueryHandler(cb_category,     pattern=r"^cat:"))
    app.add_handler(CallbackQueryHandler(cb_add,          pattern=r"^add"))
    app.add_handler(CallbackQueryHandler(cb_price_choice, pattern=r"^price:"))
    app.add_handler(CallbackQueryHandler(cb_beer_add,     pattern=r"^badd"))
    app.add_handler(CallbackQueryHandler(cb_sell_skip,    pattern=r"^sell:skip:"))
    app.add_handler(CallbackQueryHandler(cb_remove,       pattern=r"^rm:"))
    app.add_handler(CallbackQueryHandler(cb_edit,         pattern=r"^edit:"))
    app.add_handler(CallbackQueryHandler(cb_order,        pattern=r"^order:"))
    app.add_handler(CallbackQueryHandler(cb_orders_page,  pattern=r"^ordpage:"))
    app.add_handler(CallbackQueryHandler(cb_delete_order, pattern=r"^delorder:"))
    app.add_handler(CallbackQueryHandler(cb_view_order,   pattern=r"^vieworder:"))
    app.add_handler(CallbackQueryHandler(cb_stats,        pattern=r"^stats:"))

    app.add_handler(MessageHandler(filters.Document.ALL,            handle_file))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    print("✅ Бот запущен.", flush=True)
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
