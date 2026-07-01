"""
АлкоПартнёр — Telegram-бот для поиска по прайс-листу МАИ + пиво
Версия: обновлённая (повтор заказа, аналитика по товарам, свой период, автоочистка чата)

Установка:
  pip install python-telegram-bot pandas xlrd==1.2.0 openpyxl gspread google-auth

Запуск:
  python bot.py
"""

import os, re, json, logging, sys, pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from collections import defaultdict
from datetime import datetime, timedelta
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, BotCommand
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler,
    CallbackQueryHandler, filters, ContextTypes
)

# ══════════════════════════════════════════════════════════════════
TOKEN          = os.environ.get("BOT_TOKEN", "8974641448:AAFhIYLya0lVhRsldJj4a2UXNKvXv0rQPD8")
PRICE_XLS      = "price.xls"
PRICE_BEER     = "pricepivo.xls"
ORDERS_FILE    = "orders.txt"
ORDERS_JSON    = "orders.json"
GOOGLE_SHEET_ID   = os.environ.get("GOOGLE_SHEET_ID", "")
GOOGLE_CREDS_JSON = os.environ.get("GOOGLE_CREDS_JSON", "")
KNOWN_CLIENTS_LIMIT = 8
# ══════════════════════════════════════════════════════════════════

logging.basicConfig(format="%(asctime)s | %(levelname)s | %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

df_price     = None
_name_index: list[str] = []
df_beer      = None
_beer_index: list[str] = []


# ─────────────────────────────────────────────────────────────────
# УТИЛИТЫ
# ─────────────────────────────────────────────────────────────────

def normalize(t: str) -> str:
    return re.sub(r"\s+", " ", t.lower().strip().replace(".", ","))


# ─────────────────────────────────────────────────────────────────
# АВТООЧИСТКА ЧАТА ОТ СЛУЖЕБНЫХ СООБЩЕНИЙ
# ─────────────────────────────────────────────────────────────────
# Идея: результаты поиска, подтверждения "добавлено в корзину" и запросы
# цены продажи — временные (flow) сообщения. Как только появляется новое
# сообщение того же типа, предыдущее удаляется. На "чекпоинтах" (открыли
# корзину, подтвердили заказ, начали новый заказ) очередь чистится совсем.

async def _clear_flow(context, bot, chat_id):
    ids = context.user_data.pop("flow_msgs", [])
    for mid in ids:
        try:
            await bot.delete_message(chat_id=chat_id, message_id=mid)
        except Exception:
            pass

async def _set_flow(context, bot, chat_id, new_ids):
    await _clear_flow(context, bot, chat_id)
    context.user_data["flow_msgs"] = [i for i in new_ids if i]

async def _delete_one(bot, chat_id, message_id):
    try:
        await bot.delete_message(chat_id=chat_id, message_id=message_id)
    except Exception:
        pass


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
    """Дописывает строку заказа в Google Sheets (включая JSON с позициями,
    чтобы аналитика по товарам и повтор заказа работали и через Sheets)."""
    sheet = get_sheet()
    if sheet is None:
        return False
    try:
        now = datetime.now().strftime("%d.%m.%Y %H:%M")
        items_str = "; ".join(
            f"{i['name']} ×{i['qty']} по {i['price']:.2f}₽"
            + (f" → продажа {i['sell_price']:.2f}₽" if i.get("sell_price") else "")
            for i in cart["items"]
        )
        items_json = json.dumps(cart["items"], ensure_ascii=False)
        sheet.append_row([
            now,
            cart["client"],
            round(buy_tot,  2),
            round(sell_tot, 2),
            round(margin,   2),
            items_str,
            items_json,
        ])
        logger.info("Заказ записан в Google Sheets")
        return True
    except Exception as e:
        logger.error(f"save_order_to_sheets: {e}")
        return False


def load_orders_from_sheets() -> list:
    """Загружает все заказы из Google Sheets, включая позиции (колонка G)."""
    sheet = get_sheet()
    if sheet is None:
        return []
    try:
        rows   = sheet.get_all_values()
        orders = []
        for row_num, row in enumerate(rows[1:], start=2):  # пропускаем заголовок
            if len(row) < 5 or not row[0]:
                continue
            try:
                dt_str = row[0]
                try:
                    dt_obj = datetime.strptime(dt_str, "%d.%m.%Y %H:%M")
                    dt_iso = dt_obj.strftime("%Y-%m-%d %H:%M:%S")
                except Exception:
                    dt_iso = dt_str

                items = []
                if len(row) > 6 and row[6]:
                    try:
                        items = json.loads(row[6])
                    except Exception:
                        items = []

                orders.append({
                    "dt":         dt_iso,
                    "client":     row[1] if len(row) > 1 else "?",
                    "buy_total":  float(row[2]) if len(row) > 2 and row[2] else 0,
                    "sell_total": float(row[3]) if len(row) > 3 and row[3] else 0,
                    "margin":     float(row[4]) if len(row) > 4 and row[4] else 0,
                    "items_str":  row[5] if len(row) > 5 else "",
                    "items":      items,
                    "sheet_row":  row_num,
                })
            except Exception:
                continue
        return orders
    except Exception as e:
        logger.error(f"load_orders_from_sheets: {e}")
        return []


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
            import xlrd
            wb   = xlrd.open_workbook(xlsx_path)
            rows = []
            sh   = wb.sheet_by_index(0)
            for rx in range(sh.nrows):
                rows.append(sh.row_values(rx)[:3])
            raw = pd.DataFrame(rows, columns=["name", "price_factory", "price_nw"])
        else:
            raw = pd.read_excel(xlsx_path, header=None, engine="openpyxl")
            raw = raw.iloc[:, :3]
            raw.columns = ["name", "price_factory", "price_nw"]

        raw["price_factory"] = pd.to_numeric(raw["price_factory"], errors="coerce")
        raw["price_nw"]      = pd.to_numeric(raw["price_nw"],      errors="coerce")

        df = raw[raw["price_factory"].notna() & (raw["price_factory"] > 0)].copy()
        df["name"] = df["name"].astype(str).str.strip()
        df = df[df["name"] != ""]
        df = df[~df["name"].str.upper().isin(["NAN", "НАИМЕНОВАНИЕ", "НАЗВАНИЕ",
                                               "ЗАВОД НАЛ", "ЦЕНА СЕВЕРО ЗАПАД"])]
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


def find_current_price_and_stock(name: str):
    """Пытается найти актуальную цену/остаток по названию из прошлого заказа
    (используется при повторе заказа, т.к. прайс мог обновиться)."""
    is_beer = name.startswith("🍺 ")
    clean = name[2:].strip() if is_beer else name
    if is_beer and df_beer is not None:
        matches = df_beer[df_beer["name"] == clean]
        if not matches.empty:
            r = matches.iloc[0]
            return float(r["price_factory"]), None
    elif not is_beer and df_price is not None:
        matches = df_price[df_price["name"] == clean]
        if not matches.empty:
            r = matches.iloc[0]
            return float(r["price"]), int(r["stock"])
    return None, None


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
# ЗАКАЗЫ, ИСТОРИЯ КЛИЕНТОВ, ПОВТОР ЗАКАЗА
# ─────────────────────────────────────────────────────────────────

def load_orders_json() -> list:
    """Загружает заказы: сначала из Google Sheets, потом из локального JSON."""
    if GOOGLE_SHEET_ID and GOOGLE_CREDS_JSON:
        try:
            orders = load_orders_from_sheets()
            if orders is not None:
                return orders
        except Exception as e:
            logger.warning(f"Sheets fallback to JSON: {e}")
    try:
        if os.path.exists(ORDERS_JSON):
            with open(ORDERS_JSON, encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return []


def get_known_clients(limit: int = KNOWN_CLIENTS_LIMIT) -> list[str]:
    """Список клиентов от самых недавних заказов, без повторов."""
    orders = load_orders_json()
    seen = []
    for o in reversed(orders):
        c = (o.get("client") or "").strip()
        if c and c not in seen:
            seen.append(c)
        if len(seen) >= limit:
            break
    return seen


def get_last_order_for_client(client: str) -> dict | None:
    orders = load_orders_json()
    for o in reversed(orders):
        if (o.get("client") or "").strip().lower() == client.strip().lower():
            return o
    return None


def build_cart_from_last_order(cart: dict, last_order: dict) -> list[str]:
    """Заполняет корзину позициями из прошлого заказа, обновляя цену/остаток
    по текущему прайсу там, где это возможно. Возвращает список предупреждений."""
    warnings = []
    cart["items"] = []
    for it in last_order.get("items", []):
        name = it.get("name", "")
        qty  = int(it.get("qty", 1))
        stored_price = float(it.get("price", 0))
        cur_price, stock = find_current_price_and_stock(name)
        if cur_price is not None:
            price = cur_price
            if stock is not None and stock < qty:
                warnings.append(f"⚠️ {name}: в наличии только {stock} шт. (запрошено {qty})")
        else:
            price = stored_price
            warnings.append(f"⚠️ {name}: не найден в текущем прайсе, взята прошлая цена {stored_price:.2f}₽")
        cart["items"].append({"name": name, "price": price, "qty": qty, "sell_price": None})
    return warnings


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
    sheets_ok = save_order_to_sheets(cart, buy_tot, sell_tot, margin)
    if sheets_ok:
        logger.info("Заказ → Google Sheets ✅")
    else:
        logger.warning("Google Sheets недоступен, сохраняем в JSON")

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


# ─────────────────────────────────────────────────────────────────
# АНАЛИТИКА — маржа по клиентам И по товарам, произвольный период
# ─────────────────────────────────────────────────────────────────

def _orders_in_range(orders, start_dt, end_dt):
    result = []
    for o in orders:
        try:
            dt = datetime.strptime(o["dt"], "%Y-%m-%d %H:%M:%S")
        except Exception:
            continue
        if start_dt <= dt <= end_dt:
            result.append(o)
    return result


def analytics_for_range(orders: list, start_dt: datetime, end_dt: datetime) -> str:
    period = _orders_in_range(orders, start_dt, end_dt)
    if not period:
        return "Заказов за этот период нет."

    buy    = sum(o["buy_total"]  for o in period)
    sell   = sum(o["sell_total"] for o in period)
    margin = sum(o["margin"]     for o in period)

    by_client  = defaultdict(lambda: {"margin": 0.0, "orders": 0})
    by_product = defaultdict(lambda: {"margin": 0.0, "qty": 0})
    has_items  = False

    for o in period:
        by_client[o["client"]]["margin"] += o["margin"]
        by_client[o["client"]]["orders"] += 1
        for it in o.get("items", []):
            has_items = True
            name  = it.get("name", "?")
            qty   = int(it.get("qty", 0))
            price = float(it.get("price", 0))
            sp    = float(it.get("sell_price") or price)
            by_product[name]["margin"] += (sp - price) * qty
            by_product[name]["qty"]    += qty

    top_clients  = sorted(by_client.items(),  key=lambda x: -x[1]["margin"])[:5]
    top_products = sorted(by_product.items(), key=lambda x: -x[1]["margin"])[:5]

    lines = [
        f"📊 Заказов: *{len(period)}*",
        f"📦 Закупка: *{buy:.0f} ₽*",
        f"💸 Продажа: *{sell:.0f} ₽*",
        f"📈 Маржа: *{margin:.0f} ₽*",
    ]
    if top_clients:
        lines += ["", "🏆 *Топ клиентов по марже:*"]
        for client, d in top_clients:
            lines.append(f"  • {client}: {d['margin']:.0f} ₽ ({d['orders']} зак.)")
    if has_items and top_products:
        lines += ["", "📦 *Топ товаров по марже:*"]
        for name, d in top_products:
            lines.append(f"  • {name[:40]}: {d['margin']:.0f} ₽ ({d['qty']} шт.)")
    elif not has_items:
        lines += ["", "_ℹ️ Разбивка по товарам недоступна для заказов, оформленных до обновления бота_"]
    return "\n".join(lines)


def analytics_for_period(orders: list, days: int) -> str:
    end   = datetime.now()
    start = end - timedelta(days=days)
    return analytics_for_range(orders, start, end)


# ─────────────────────────────────────────────────────────────────
# ПОДСКАЗКА / ИСТОРИЯ ЦЕН ПО КЛИЕНТУ
# ─────────────────────────────────────────────────────────────────

def client_item_history(client: str, product_name: str) -> list[float]:
    if not client:
        return []
    prices = []
    for o in load_orders_json():
        if (o.get("client") or "").strip().lower() != client.strip().lower():
            continue
        for item in o.get("items", []):
            if product_name.lower() in item.get("name", "").lower():
                sp = item.get("sell_price")
                if sp:
                    prices.append(float(sp))
    return prices


def client_price_hint(client: str, product_name: str) -> str:
    prices = client_item_history(client, product_name)
    if not prices:
        return ""
    trend = " → ".join(f"{p:.0f}" for p in prices[-3:])
    return (
        f"💡 *Подсказка по {client}:*\n"
        f"   История продаж: {trend} ₽\n"
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
    await _clear_flow(context, context.bot, msg_obj.chat_id)
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

async def _send_alco_results(msg_obj, results, query, extra_ids=None) -> list[int]:
    if not results:
        await msg_obj.reply_text(f"🍷 «{query}» — не найдено в алко прайсе.")
        return []
    MAX = 20
    hdr = f"🍷 *{len(results)}* позиций по «{query}»"
    if len(results) > MAX:
        hdr += f"\n_(первые {MAX})_"

    sent_ids = list(extra_ids or [])
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

    return sent_ids


async def _send_beer_results(msg_obj, results, query, extra_ids=None) -> list[int]:
    if not results:
        await msg_obj.reply_text(f"🍺 «{query}» — не найдено в пивном прайсе.")
        return []
    MAX = 20
    hdr = f"🍺 *{len(results)}* позиций по «{query}»"
    if len(results) > MAX:
        hdr += f"\n_(первые {MAX})_"

    sent_ids = list(extra_ids or [])
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

    return sent_ids


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    aw   = context.user_data.get("awaiting")
    chat_id = update.message.chat_id

    # ── Ожидание дат для аналитики за свой период ─────────────────
    if aw == "stats_from":
        try:
            dt = datetime.strptime(text, "%d.%m.%Y")
        except ValueError:
            await update.message.reply_text("Формат даты: `ДД.ММ.ГГГГ`, например `01.06.2026`", parse_mode="Markdown")
            return
        context.user_data["stats_from"] = dt
        context.user_data["awaiting"] = "stats_to"
        await update.message.reply_text("Введи дату окончания периода (`ДД.ММ.ГГГГ`):", parse_mode="Markdown")
        return

    if aw == "stats_to":
        try:
            dt_to = datetime.strptime(text, "%d.%m.%Y")
        except ValueError:
            await update.message.reply_text("Формат даты: `ДД.ММ.ГГГГ`", parse_mode="Markdown")
            return
        dt_from = context.user_data.pop("stats_from", None)
        context.user_data.pop("awaiting", None)
        if dt_from is None:
            await update.message.reply_text("Начни заново: /stats")
            return
        dt_to_full = dt_to.replace(hour=23, minute=59, second=59)
        stats_text = analytics_for_range(load_orders_json(), dt_from, dt_to_full)
        await update.message.reply_text(
            f"📊 *Аналитика {dt_from.strftime('%d.%m.%Y')} — {dt_to.strftime('%d.%m.%Y')}*\n\n{stats_text}",
            parse_mode="Markdown"
        )
        return

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
        m = await update.message.reply_text(
            f"👤 Клиент: *{cart['client']}*\n\n{added_msg}"
            + fmt_cart(cart) + "\n\nИщи товары или /confirm",
            parse_mode="Markdown"
        )
        await _set_flow(context, context.bot, chat_id, [m.message_id, update.message.message_id])
        return

    # ── Ожидание количества (алко) ────────────────────────────────
    if aw == "qty":
        m_ = re.search(r'\d+', text)
        if not m_:
            await update.message.reply_text("Введи число, например `6`", parse_mode="Markdown")
            return
        qty = int(m_.group())
        context.user_data.pop("awaiting", None)
        idx = context.user_data.pop("pending_idx", None)
        if idx is not None and df_price is not None:
            row  = df_price.iloc[idx]
            cart = get_cart(context)
            if not cart["client"]:
                context.user_data["awaiting"] = "client"
                context.user_data["pending"]  = {"idx": idx, "qty": qty}
                m = await update.message.reply_text("👤 Введи имя клиента:")
                await _set_flow(context, context.bot, chat_id, [m.message_id, update.message.message_id])
                return
            add_to_cart(cart, row["name"], float(row["price"]), qty)
            hint = client_price_hint(cart["client"], row["name"])
            m = await update.message.reply_text(
                f"✅ *{row['name']}* × {qty} шт.\n{hint}" + fmt_cart(cart) + "\n\nИщи или /confirm",
                parse_mode="Markdown"
            )
            await _set_flow(context, context.bot, chat_id, [m.message_id, update.message.message_id])
        return

    # ── Ожидание количества (пиво) ────────────────────────────────
    if aw == "beer_qty":
        m_ = re.search(r'\d+', text)
        if not m_:
            await update.message.reply_text("Введи число, например `6`", parse_mode="Markdown")
            return
        qty = int(m_.group())
        idx = context.user_data.pop("pending_beer_idx", None)
        context.user_data.pop("awaiting", None)
        if idx is not None and df_beer is not None:
            row  = df_beer.iloc[idx]
            cart = get_cart(context)
            if not cart["client"]:
                context.user_data["awaiting"] = "client"
                context.user_data["pending"]  = {"beer_idx": idx, "qty": qty, "price_type": "factory"}
                m = await update.message.reply_text("👤 Введи имя клиента:")
                await _set_flow(context, context.bot, chat_id, [m.message_id, update.message.message_id])
                return
            name  = f"🍺 {row['name']}"
            price = float(row["price_factory"])
            add_to_cart(cart, name, price, qty)
            m = await update.message.reply_text(
                f"✅ *{name}* × {qty} шт.\n" + fmt_cart(cart) + "\n\nИщи или /confirm",
                parse_mode="Markdown"
            )
            await _set_flow(context, context.bot, chat_id, [m.message_id, update.message.message_id])
        return

    # ── Ожидание нового количества при редактировании ────────────
    if aw == "edit_qty":
        m_ = re.search(r'\d+', text)
        if not m_:
            await update.message.reply_text("Введи число, например `6`", parse_mode="Markdown")
            return
        qty  = int(m_.group())
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
        m_ = re.search(r'[\d.,]+', text.replace(",", "."))
        if not m_:
            await update.message.reply_text("Введи число, например `390`", parse_mode="Markdown")
            return
        try:
            price = float(m_.group().replace(",", "."))
        except ValueError:
            await update.message.reply_text("Введи число, например `390`", parse_mode="Markdown")
            return
        idx  = context.user_data.get("sell_idx", 0)
        cart = get_cart(context)
        if idx < len(cart["items"]):
            cart["items"][idx]["sell_price"] = price
        context.user_data["sell_idx"] = idx + 1
        await _delete_one(context.bot, chat_id, update.message.message_id)
        await _ask_sell_price(update.message, context)
        return

    # ── Обычный поиск ─────────────────────────────────────────────
    if len(text) < 2:
        return

    if df_price is None and df_beer is None:
        await update.message.reply_text("📤 Загрузи прайс — отправь `.xls` файл.")
        return

    if df_price is not None and df_beer is not None:
        context.user_data["pending_search"] = text
        m = await update.message.reply_text(
            f"Ищем «{text}» — выбери категорию:",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton("🍷 Алкоголь", callback_data="cat:alco"),
                InlineKeyboardButton("🍺 Пиво",     callback_data="cat:beer"),
                InlineKeyboardButton("🔍 Везде",    callback_data="cat:all"),
            ]])
        )
        await _set_flow(context, context.bot, chat_id, [m.message_id, update.message.message_id])
        return

    query, min_p, max_p = parse_price_filter(text)
    if not query.strip():
        query = text

    if df_price is None:
        results = do_search_beer(query, min_p, max_p)
        ids = await _send_beer_results(update.message, results, query, extra_ids=[update.message.message_id])
        if ids:
            await _set_flow(context, context.bot, chat_id, ids)
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
        m = await update.message.reply_text(card, parse_mode="Markdown", reply_markup=kb)
        await _set_flow(context, context.bot, chat_id, [m.message_id, update.message.message_id])
        return

    ids = await _send_alco_results(update.message, results, query, extra_ids=[update.message.message_id])
    if ids:
        await _set_flow(context, context.bot, chat_id, ids)


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
        ids = await _send_alco_results(q.message, do_search(query, min_p, max_p), query,
                                        extra_ids=[q.message.message_id])
        if ids:
            await _set_flow(context, context.bot, q.message.chat_id, ids)
    elif cat == "beer":
        ids = await _send_beer_results(q.message, do_search_beer(query, min_p, max_p), query,
                                        extra_ids=[q.message.message_id])
        if ids:
            await _set_flow(context, context.bot, q.message.chat_id, ids)
    else:
        alco = do_search(query, min_p, max_p)
        beer = do_search_beer(query, min_p, max_p)
        if not alco and not beer:
            await q.message.reply_text(f"🔍 «{query}» — нигде не найдено.")
            return
        all_ids = [q.message.message_id]
        if alco:
            all_ids += await _send_alco_results(q.message, alco, query)
        if beer:
            all_ids += await _send_beer_results(q.message, beer, query)
        await _set_flow(context, context.bot, q.message.chat_id, all_ids)


# ─────────────────────────────────────────────────────────────────
# CALLBACK — ДОБАВЛЕНИЕ В КОРЗИНУ (АЛКО)
# ─────────────────────────────────────────────────────────────────

async def cb_add(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q      = update.callback_query
    await q.answer()
    parts  = q.data.split(":", 2)
    action = parts[0]
    idx    = int(parts[-1])
    chat_id = q.message.chat_id

    if action == "addx":
        context.user_data["awaiting"]    = "qty"
        context.user_data["pending_idx"] = idx
        m = await q.message.reply_text("Введи количество штук:")
        await _set_flow(context, context.bot, chat_id, [m.message_id])
        return

    qty  = int(parts[1])
    row  = df_price.iloc[idx]
    cart = get_cart(context)

    if not cart["client"]:
        context.user_data["awaiting"] = "client"
        context.user_data["pending"]  = {"idx": idx, "qty": qty}
        m = await q.message.reply_text("👤 Введи имя клиента:")
        await _set_flow(context, context.bot, chat_id, [m.message_id])
        return

    has_p = pd.notna(row.get("promo_price")) and float(row.get("promo_price", 0)) > 0
    if has_p:
        promo = float(row["promo_price"])
        base  = float(row["price"])
        cond  = str(row.get("promo_cond", "")).strip()
        cond_s = f"\n📋 Условие: _{cond}_" if cond else ""
        m = await q.message.reply_text(
            f"📦 *{row['name']}* × {qty} шт.\n\nВыбери цену закупки:{cond_s}",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton(f"💰 МАИ {base:.2f}₽",        callback_data=f"price:base:{idx}:{qty}"),
                InlineKeyboardButton(f"🎯 Маркетинг {promo:.2f}₽", callback_data=f"price:promo:{idx}:{qty}"),
            ]])
        )
        await _set_flow(context, context.bot, chat_id, [m.message_id])
        return

    add_to_cart(cart, row["name"], float(row["price"]), qty)
    hint = client_price_hint(cart["client"], row["name"])
    m = await q.message.reply_text(
        f"✅ *{row['name']}* × {qty} шт.\n{hint}" + fmt_cart(cart) + "\n\nИщи или /confirm",
        parse_mode="Markdown"
    )
    await _set_flow(context, context.bot, chat_id, [m.message_id])


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
    m = await q.message.reply_text(
        f"✅ *{row['name']}* × {qty} шт. ({label}: {price:.2f}₽)\n{hint}"
        + fmt_cart(cart) + "\n\nИщи или /confirm",
        parse_mode="Markdown"
    )
    await _set_flow(context, context.bot, q.message.chat_id, [m.message_id])


# ─────────────────────────────────────────────────────────────────
# CALLBACK — ДОБАВЛЕНИЕ В КОРЗИНУ (ПИВО)
# ─────────────────────────────────────────────────────────────────

async def cb_beer_add(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q      = update.callback_query
    await q.answer()
    parts  = q.data.split(":")
    action = parts[0]
    chat_id = q.message.chat_id

    if action == "baddx":
        idx = int(parts[1])
        context.user_data["awaiting"]         = "beer_qty"
        context.user_data["pending_beer_idx"] = idx
        m = await q.message.reply_text("Введи количество штук:")
        await _set_flow(context, context.bot, chat_id, [m.message_id])
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
        m = await q.message.reply_text("👤 Введи имя клиента:")
        await _set_flow(context, context.bot, chat_id, [m.message_id])
        return

    add_to_cart(cart, name, price, qty)
    hint = client_price_hint(cart["client"], row["name"])
    m = await q.message.reply_text(
        f"✅ *{name}* × {qty} шт. ({label}: {price:.2f}₽)\n{hint}"
        + fmt_cart(cart) + "\n\nИщи или /confirm",
        parse_mode="Markdown"
    )
    await _set_flow(context, context.bot, chat_id, [m.message_id])


# ─────────────────────────────────────────────────────────────────
# КОРЗИНА — ПРОСМОТР И РЕДАКТИРОВАНИЕ
# ─────────────────────────────────────────────────────────────────

async def _cmd_cart_msg(msg_obj, context):
    await _clear_flow(context, context.bot, msg_obj.chat_id)
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
    await _clear_flow(context, context.bot, msg_obj.chat_id)
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
    await _clear_flow(context, context.bot, update.message.chat_id)
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
        await _clear_flow(context, context.bot, q.message.chat_id)
        await q.message.edit_text(
            fmt_cart(cart, title="✅ *Заказ принят*"), parse_mode="Markdown"
        )
        await q.message.reply_text("/neworder — новый заказ")
        context.user_data["cart"] = {"client": "", "items": []}


# ─────────────────────────────────────────────────────────────────
# ЦЕНЫ ПЕРЕПРОДАЖИ (всегда с историей цен клиента)
# ─────────────────────────────────────────────────────────────────

async def _ask_sell_price(msg_obj, context):
    cart = get_cart(context)
    idx  = context.user_data.get("sell_idx", 0)
    if idx >= len(cart["items"]):
        context.user_data.pop("sell_idx", None)
        context.user_data.pop("awaiting", None)
        await _clear_flow(context, context.bot, msg_obj.chat_id)
        await msg_obj.reply_text(
            fmt_cart(cart, title="✅ *Цены введены*") + "\n\n/confirm — подтвердить",
            parse_mode="Markdown"
        )
        return
    item  = cart["items"][idx]
    p_s   = f"{item['price']:.2f}".rstrip("0").rstrip(".")
    cur   = (f"текущая: *{item['sell_price']:.2f} ₽*"
             if item.get("sell_price") else f"цена закупки: *{p_s} ₽*")

    hist = client_item_history(cart["client"], item["name"])
    if hist:
        trend = " → ".join(f"{p:.0f}" for p in hist[-3:])
        hist_line = f"📜 Ранее продавал по: {trend} ₽ (сред. {sum(hist)/len(hist):.0f} ₽)\n"
    else:
        hist_line = "📜 Раньше этому клиенту не продавали этот товар\n"

    m = await msg_obj.reply_text(
        f"💸 *Цена продажи {idx+1} из {len(cart['items'])}*\n\n"
        f"📦 {item['name']}\n"
        f"   Закупка: *{p_s} ₽* × {item['qty']} шт.\n"
        f"   {cur}\n"
        f"{hist_line}\nВведи цену продажи или пропусти:",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup([[
            InlineKeyboardButton("➡️ Оставить", callback_data=f"sell:skip:{idx}"),
        ]])
    )
    context.user_data["awaiting"] = "sell_price"
    await _set_flow(context, context.bot, msg_obj.chat_id, [m.message_id])

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
    await _clear_flow(context, context.bot, msg_obj.chat_id)
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
    """Удаляет заказ. Работает и для заказов из Google Sheets (по номеру строки),
    и для локального JSON — раньше при работе через Sheets заказ "удалялся"
    только из JSON и появлялся снова при следующей загрузке."""
    q   = update.callback_query
    await q.answer()
    idx = int(q.data.split(":")[1])
    orders = load_orders_json()
    if not (0 <= idx < len(orders)):
        return
    removed = orders[idx]

    if "sheet_row" in removed:
        sheet = get_sheet()
        if sheet is not None:
            try:
                sheet.delete_rows(removed["sheet_row"])
            except Exception as e:
                logger.warning(f"delete sheet row: {e}")

    try:
        if os.path.exists(ORDERS_JSON):
            with open(ORDERS_JSON, encoding="utf-8") as f:
                local = json.load(f)
            local = [o for o in local
                     if not (o.get("dt") == removed.get("dt") and o.get("client") == removed.get("client"))]
            with open(ORDERS_JSON, "w", encoding="utf-8") as f:
                json.dump(local, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.warning(f"delete local order: {e}")

    client = removed.get("client", "?")
    dt     = removed.get("dt", "")[:10]
    remaining = load_orders_json()
    await q.message.reply_text(
        f"🗑 Заказ *{client}* от {dt} удалён. Осталось: {len(remaining)}",
        parse_mode="Markdown"
    )
    if remaining:
        await _show_orders_page(q.message, context, page=0)


# ─────────────────────────────────────────────────────────────────
# АНАЛИТИКА
# ─────────────────────────────────────────────────────────────────

def _stats_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("📅 День",   callback_data="stats:1"),
        InlineKeyboardButton("📆 Неделя", callback_data="stats:7"),
        InlineKeyboardButton("🗓 Месяц",  callback_data="stats:30"),
    ], [
        InlineKeyboardButton("📆 Свой период", callback_data="stats:custom"),
    ]])

async def cmd_stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    orders = load_orders_json()
    if not orders:
        await update.message.reply_text("📊 Нет данных. Подтверди хотя бы один заказ.")
        return
    await _clear_flow(context, context.bot, update.message.chat_id)
    await update.message.reply_text(
        "📊 *Аналитика — выбери период:*",
        parse_mode="Markdown",
        reply_markup=_stats_kb()
    )

async def cb_stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q   = update.callback_query
    await q.answer()
    val = q.data.split(":")[1]

    if val == "custom":
        context.user_data["awaiting"] = "stats_from"
        await q.message.reply_text(
            "Введи дату начала периода (`ДД.ММ.ГГГГ`), например `01.06.2026`:",
            parse_mode="Markdown"
        )
        return

    days   = int(val)
    labels = {1: "сегодня", 7: "за неделю", 30: "за месяц"}
    text   = analytics_for_period(load_orders_json(), days)
    await q.message.edit_text(
        f"📊 *Аналитика — {labels[days]}*\n\n{text}",
        parse_mode="Markdown",
        reply_markup=_stats_kb()
    )


# ─────────────────────────────────────────────────────────────────
# НОВЫЙ ЗАКАЗ — выбор из недавних клиентов + повтор заказа
# ─────────────────────────────────────────────────────────────────

async def _start_new_order(msg_obj, context):
    context.user_data["cart"] = {"client": "", "items": []}
    context.user_data.pop("pending", None)
    context.user_data.pop("awaiting", None)
    await _clear_flow(context, context.bot, msg_obj.chat_id)

    clients = get_known_clients()
    if not clients:
        context.user_data["awaiting"] = "client"
        await msg_obj.reply_text("🛒 Новый заказ.\n\n👤 Введи имя клиента:")
        return

    context.user_data["known_clients"] = clients
    btns = [[InlineKeyboardButton(c, callback_data=f"neword:pick:{i}")] for i, c in enumerate(clients)]
    btns.append([InlineKeyboardButton("➕ Другой клиент", callback_data="neword:other")])
    await msg_obj.reply_text(
        "🛒 *Новый заказ*\n\n👤 Выбери клиента или добавь нового:",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(btns)
    )

async def cmd_neworder(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await _start_new_order(update.message, context)

async def cb_neword(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q      = update.callback_query
    await q.answer()
    parts  = q.data.split(":")
    action = parts[1]
    cart   = get_cart(context)

    if action == "other":
        context.user_data["awaiting"] = "client"
        await q.message.reply_text("👤 Введи имя клиента:")
        return

    if action == "pick":
        idx = int(parts[2])
        clients = context.user_data.get("known_clients", [])
        if idx >= len(clients):
            await q.message.reply_text("Список устарел, начни заново: /neworder")
            return
        client = clients[idx]
        cart["client"] = client
        last = get_last_order_for_client(client)
        if last and last.get("items"):
            n_items = len(last["items"])
            await q.message.reply_text(
                f"👤 Клиент: *{client}*\n\nНайден прошлый заказ ({n_items} поз.). Повторить его или собрать новый?",
                parse_mode="Markdown",
                reply_markup=InlineKeyboardMarkup([[
                    InlineKeyboardButton(f"🔁 Повторить ({n_items} поз.)", callback_data="neword:repeat"),
                    InlineKeyboardButton("🆕 Собрать новый", callback_data="neword:fresh"),
                ]])
            )
        else:
            await q.message.reply_text(
                f"👤 Клиент: *{client}*\n\nИщи товары для заказа.",
                parse_mode="Markdown"
            )
        return

    if action == "repeat":
        client = cart.get("client", "")
        last = get_last_order_for_client(client)
        if not last or not last.get("items"):
            await q.message.reply_text("Не нашёл детальный прошлый заказ. Собери вручную.")
            return
        warnings = build_cart_from_last_order(cart, last)
        text = fmt_cart(cart, title="🔁 *Заказ восстановлен*")
        if warnings:
            text += "\n\n" + "\n".join(warnings)
        text += "\n\nПроверь количество/цены и жми /confirm, либо ищи товары чтобы добавить ещё."
        await q.message.reply_text(text, parse_mode="Markdown")
        return

    if action == "fresh":
        await q.message.reply_text(
            f"👤 Клиент: *{cart.get('client','')}*\n\nИщи товары для заказа.",
            parse_mode="Markdown"
        )
        return


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
        await _start_new_order(q.message, context)
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
            await _clear_flow(context, context.bot, q.message.chat_id)
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
            await _clear_flow(context, context.bot, q.message.chat_id)
            await q.message.reply_text(
                "📊 *Аналитика — выбери период:*",
                parse_mode="Markdown",
                reply_markup=_stats_kb()
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
    app.add_handler(CallbackQueryHandler(cb_neword,       pattern=r"^neword:"))
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
