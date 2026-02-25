# main.py - ربات مدیریت باشگاه ورزشی برای بله (نسخه نهایی کامل)
import os
import logging
import asyncio
import time
from datetime import datetime
from typing import Dict, List, Set, Tuple
from dotenv import load_dotenv
import aiohttp
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import jdatetime

# --------------------------- ENV LOAD ---------------------------
load_dotenv()
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

BALE_TOKEN = os.getenv("BALE_TOKEN")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
CREDENTIALS_JSON = os.getenv("CREDENTIALS_JSON", "credentials.json")
ADMIN_IDS = os.getenv("ADMIN_IDS", "")
PAYMENT_WALLET_TOKEN = os.getenv("PAYMENT_TOKEN", "")

ADMINS = set(int(x.strip()) for x in ADMIN_IDS.split(",") if x.strip())

# --------------------------- GOOGLE SHEETS ---------------------------
def gs_client():
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_JSON, scope)
    return gspread.authorize(creds)

def sheet_users():
    ss = gs_client().open_by_key(SPREADSHEET_ID)
    try:
        ws = ss.worksheet("users")
    except:
        ws = ss.add_worksheet("users", 2000, 10)
        ws.append_row(["user_id", "name", "category", "total_fee", "allowed_admin"])
    return ws

def sheet_sessions():
    ss = gs_client().open_by_key(SPREADSHEET_ID)
    try:
        ws = ss.worksheet("sessions")
    except:
        ws = ss.add_worksheet("sessions", 200, 10)
        ws.append_row(["session_id", "name", "price", "date", "attended_users"])
    return ws

def sheet_logs():
    ss = gs_client().open_by_key(SPREADSHEET_ID)
    try:
        ws = ss.worksheet("logs")
    except:
        ws = ss.add_worksheet("logs", 2000, 10)
        ws.append_row(["timestamp", "user_id", "session_id"])
    return ws

def sheet_payments():
    """شیت مخصوص ثبت تراکنش‌های پرداخت"""
    ss = gs_client().open_by_key(SPREADSHEET_ID)
    try:
        ws = ss.worksheet("payments")
    except:
        ws = ss.add_worksheet("payments", 2000, 10)
        ws.append_row(["timestamp", "user_id", "name", "amount", "payment_id", "status"])
    return ws

# --------------------------- UTILITIES ---------------------------
def is_admin(uid: int) -> bool:
    return uid in ADMINS

def get_users():
    try:
        ws = sheet_users()
        data = ws.get_all_values()
        if not data or len(data) < 2:
            return []
        headers = data[0]
        records = []
        for row in data[1:]:
            if any(row):
                record = {}
                for i, header in enumerate(headers):
                    record[header] = row[i] if i < len(row) else ''
                records.append(record)
        return records
    except Exception as e:
        logger.error(f"خطا در دریافت کاربران: {e}")
        return []

def get_sessions():
    try:
        ws = sheet_sessions()
        data = ws.get_all_values()
        if not data or len(data) < 2:
            return []
        headers = data[0]
        records = []
        for row in data[1:]:
            if any(row):
                record = {}
                for i, header in enumerate(headers):
                    record[header] = row[i] if i < len(row) else ''
                records.append(record)
        return records
    except Exception as e:
        logger.error(f"خطا در دریافت سانس‌ها: {e}")
        return []

def find_user_row(uid: str):
    """پیدا کردن سطر کاربر با آیدی داده شده، برگرداندن شماره سطر واقعی و دیکشنری داده‌ها"""
    ws = sheet_users()
    all_values = ws.get_all_values()
    if len(all_values) < 2:
        return None, None
    headers = all_values[0]
    try:
        user_id_col = headers.index("user_id")
    except ValueError:
        logger.error("ستون user_id در شیت users یافت نشد")
        return None, None
    for i, row in enumerate(all_values[1:], start=2):
        if len(row) > user_id_col and row[user_id_col].strip():
            if str(row[user_id_col]).strip() == str(uid).strip():
                record = {}
                for j, header in enumerate(headers):
                    record[header] = row[j] if j < len(row) else ''
                logger.info(f"کاربر {uid} در سطر {i} یافت شد")
                return i, record
    logger.warning(f"کاربر {uid} یافت نشد")
    return None, None

def find_session_row(session_id: str):
    ws = sheet_sessions()
    all_values = ws.get_all_values()
    if len(all_values) < 2:
        return None, None
    headers = all_values[0]
    try:
        session_id_col = headers.index("session_id")
    except ValueError:
        logger.error("ستون session_id در شیت sessions یافت نشد")
        return None, None
    for i, row in enumerate(all_values[1:], start=2):
        if len(row) > session_id_col and row[session_id_col].strip():
            if str(row[session_id_col]).strip() == str(session_id).strip():
                record = {}
                for j, header in enumerate(headers):
                    record[header] = row[j] if j < len(row) else ''
                return i, record
    return None, None

def set_user_fee(row_idx, amount):
    try:
        sheet_users().update_cell(row_idx, 4, amount)  # ستون 4 = total_fee
        logger.info(f"بدهی سطر {row_idx} به {amount} تغییر یافت")
    except Exception as e:
        logger.error(f"خطا در set_user_fee برای سطر {row_idx}: {e}")

def add_log(uid, session_id):
    ts = datetime.now().isoformat()
    sheet_logs().append_row([ts, uid, session_id])

def add_user_to_sheet(user_id: str, name: str, category: str, allowed_admin: str = "no"):
    ws = sheet_users()
    ws.append_row([user_id, name, category, 0, allowed_admin])

def add_session_to_sheet(session_id: int, name: str, price: int, date: str, attended_users: str = ""):
    ws = sheet_sessions()
    ws.append_row([session_id, name, price, date, attended_users])

def update_session_attendance(session_id: str, attended_users: str):
    idx, _ = find_session_row(session_id)
    if idx:
        ws = sheet_sessions()
        ws.update_cell(idx, 5, attended_users)
        return True
    return False

def update_user_fee(user_id: str, amount: int):
    idx, _ = find_user_row(user_id)
    if idx:
        set_user_fee(idx, amount)
        return True
    return False

def get_users_by_category(category: str):
    users = get_users()
    return [u for u in users if u['category'] == category]

def get_attended_users_for_session(session_id: str):
    idx, session_data = find_session_row(session_id)
    if idx and session_data:
        attended_users = session_data.get("attended_users", "")
        if attended_users:
            return set(attended_users.split(","))
    return set()

# --------------------------- CONVERSATION STATES ---------------------------
class ConversationState:
    NONE = 0
    ADD_USER_GET_ID = 1
    ADD_USER_GET_CATEGORY = 2
    ADD_USER_GET_NAME = 3
    SET_DEBT_SELECT_USER = 4
    SET_DEBT_GET_AMOUNT = 5
    ADD_SESSION_GET_DATE = 6
    ADD_SESSION_GET_NAME = 7
    ADD_SESSION_SELECT_CATEGORY = 8
    ADD_SESSION_SELECT_USERS = 9
    ADD_SESSION_GET_PRICE = 10
    ADD_ATTENDANCE_SELECT_SESSION = 11
    ADD_ATTENDANCE_SELECT_CATEGORY = 12
    ADD_ATTENDANCE_SELECT_USERS = 13
    ADD_ATTENDANCE_SELECT_USER = 14

# ذخیره حالت گفتگو برای هر کاربر
user_states: Dict[int, Dict] = {}

# --------------------------- BALE BOT CLASS ---------------------------
class BaleBot:
    def __init__(self, token: str):
        self.token = token
        self.base_url = f"https://tapi.bale.ai/bot{token}"
        self.session = None

    async def start(self):
        self.session = aiohttp.ClientSession()

    async def stop(self):
        if self.session:
            await self.session.close()

    async def send_message(self, chat_id: int, text: str, reply_markup=None):
        url = f"{self.base_url}/sendMessage"
        payload = {
            "chat_id": chat_id,
            "text": text,
            "parse_mode": "HTML"
        }
        if reply_markup:
            payload["reply_markup"] = reply_markup
        try:
            async with self.session.post(url, json=payload) as response:
                return await response.json()
        except Exception as e:
            logger.error(f"خطا در ارسال پیام: {e}")
            return None

    async def send_invoice(self, chat_id: int, title: str, description: str,
                          payload: str, prices: List[Dict], photo_url: str = None):
        url = f"{self.base_url}/sendInvoice"
        provider_token = PAYMENT_WALLET_TOKEN if PAYMENT_WALLET_TOKEN else "WALLET-TEST-1111111111111111"
        invoice_data = {
            "chat_id": chat_id,
            "title": title,
            "description": description,
            "payload": payload,
            "provider_token": provider_token,
            "prices": prices,
            "currency": "IRR"
        }
        if photo_url:
            invoice_data["photo_url"] = photo_url
        try:
            async with self.session.post(url, json=invoice_data) as response:
                return await response.json()
        except Exception as e:
            logger.error(f"خطا در ارسال صورتحساب: {e}")
            return None

    async def answer_pre_checkout_query(self, pre_checkout_query_id: str, ok: bool, error_message: str = None):
        url = f"{self.base_url}/answerPreCheckoutQuery"
        payload = {"pre_checkout_query_id": pre_checkout_query_id, "ok": ok}
        if not ok and error_message:
            payload["error_message"] = error_message
        try:
            async with self.session.post(url, json=payload) as response:
                return await response.json()
        except Exception as e:
            logger.error(f"خطا در پاسخ به پیش‌پرداخت: {e}")
            return None

    async def answer_callback(self, callback_id: str, text: str = None):
        url = f"{self.base_url}/answerCallbackQuery"
        payload = {"callback_query_id": callback_id}
        if text:
            payload["text"] = text
        try:
            async with self.session.post(url, json=payload) as response:
                return await response.json()
        except Exception as e:
            logger.error(f"خطا در پاسخ به callback: {e}")
            return None

    async def edit_message(self, chat_id: int, message_id: int, text: str, reply_markup=None):
        url = f"{self.base_url}/editMessageText"
        payload = {
            "chat_id": chat_id,
            "message_id": message_id,
            "text": text,
            "parse_mode": "HTML"
        }
        if reply_markup:
            payload["reply_markup"] = reply_markup
        try:
            async with self.session.post(url, json=payload) as response:
                return await response.json()
        except Exception as e:
            logger.error(f"خطا در ویرایش پیام: {e}")
            return None

# --------------------------- KEYBOARDS ---------------------------
def keyboard_categories():
    return {
        "inline_keyboard": [
            [{"text": "🏷️ شهید فهمیده", "callback_data": "cat_شهید فهمیده"}],
            [{"text": "🏷️ شهید دانشگر", "callback_data": "cat_شهید دانشگر"}],
            [{"text": "🏷️ شهید صدرزاده", "callback_data": "cat_شهید صدرزاده"}],
            [{"text": "🏷️ طلاب و دانشجویان", "callback_data": "cat_طلاب و دانشجویان"}],
            [{"text": "❌ انصراف", "callback_data": "cancel"}]
        ]
    }

def keyboard_session_categories():
    return {
        "inline_keyboard": [
            [{"text": "🏷️ شهید فهمیده", "callback_data": "session_cat_شهید فهمیده"}],
            [{"text": "🏷️ شهید دانشگر", "callback_data": "session_cat_شهید دانشگر"}],
            [{"text": "🏷️ شهید صدرزاده", "callback_data": "session_cat_شهید صدرزاده"}],
            [{"text": "🏷️ طلاب و دانشجویان", "callback_data": "session_cat_طلاب و دانشجویان"}],
            [{"text": "✅ اتمام انتخاب دسته‌بندی", "callback_data": "session_cat_done"}],
            [{"text": "❌ انصراف", "callback_data": "cancel"}]
        ]
    }

def keyboard_panel():
    return {
        "inline_keyboard": [
            [{"text": "👥 اضافه کردن کاربر", "callback_data": "add_user"}],
            [{"text": "💰 تغییر بدهی کاربر", "callback_data": "set_debt"}],
            [{"text": "🎯 اضافه کردن سانس", "callback_data": "add_session"}],
            [{"text": "📝 ثبت حضور سانس", "callback_data": "add_attendance"}],
            [{"text": "📋 لیست بدهکاران", "callback_data": "list_debts"}],
            [{"text": "🔄 بازگشت به منوی اصلی", "callback_data": "back_to_start"}]
        ]
    }

def keyboard_jalali_dates():
    today = jdatetime.datetime.now()
    keyboard = []
    for i in range(3, 0, -1):
        date = today - jdatetime.timedelta(days=i)
        date_str = date.strftime("%Y/%m/%d")
        keyboard.append([{"text": f"📅 {date_str} (گذشته)", "callback_data": f"jalali_date_{date_str}"}])
    today_str = today.strftime("%Y/%m/%d")
    keyboard.append([{"text": f"📅 {today_str} (امروز)", "callback_data": f"jalali_date_{today_str}"}])
    for i in range(1, 11):
        date = today + jdatetime.timedelta(days=i)
        date_str = date.strftime("%Y/%m/%d")
        keyboard.append([{"text": f"📅 {date_str}", "callback_data": f"jalali_date_{date_str}"}])
    keyboard.append([{"text": "📝 وارد کردن تاریخ دستی", "callback_data": "date_manual"}])
    keyboard.append([{"text": "❌ انصراف", "callback_data": "cancel"}])
    return {"inline_keyboard": keyboard}

def keyboard_users_list(users):
    keyboard = []
    for user in users:
        keyboard.append([
            {"text": f"👤 {user['name']} ({user['category']})",
             "callback_data": f"user_{user['user_id']}"}
        ])
    keyboard.append([{"text": "❌ انصراف", "callback_data": "cancel"}])
    return {"inline_keyboard": keyboard}

def keyboard_session_users(category: str, selected_users: Set[str]):
    users = get_users_by_category(category)
    keyboard = []
    for user in users:
        user_id_str = str(user['user_id'])
        if user_id_str in selected_users:
            emoji = "✅"
            callback = f"session_user_unselect_{user_id_str}"
        else:
            emoji = "◻️"
            callback = f"session_user_select_{user_id_str}"
        keyboard.append([
            {"text": f"{emoji} {user['name']}",
             "callback_data": callback}
        ])
    keyboard.append([
        {"text": "✅ انتخاب همه این دسته", "callback_data": "session_select_all_category"},
        {"text": "❌ پاک کردن این دسته", "callback_data": "session_clear_category"}
    ])
    keyboard.append([
        {"text": "➡️ انتخاب دسته‌بندی دیگر", "callback_data": "session_another_category"},
        {"text": "💾 ادامه و تعیین قیمت", "callback_data": "session_continue_price"}
    ])
    keyboard.append([{"text": "❌ انصراف", "callback_data": "cancel"}])
    return {"inline_keyboard": keyboard}

def keyboard_sessions():
    sessions = get_sessions()
    keyboard = []
    for session in sessions:
        keyboard.append([
            {"text": f"🎯 {session['name']} - {session['date']}",
             "callback_data": f"session_{session['session_id']}"}
        ])
    keyboard.append([{"text": "❌ انصراف", "callback_data": "cancel"}])
    return {"inline_keyboard": keyboard}

def keyboard_attendance_categories():
    return {
        "inline_keyboard": [
            [{"text": "🏷️ شهید فهمیده", "callback_data": "att_cat_شهید فهمیده"}],
            [{"text": "🏷️ شهید دانشگر", "callback_data": "att_cat_شهید دانشگر"}],
            [{"text": "🏷️ شهید صدرزاده", "callback_data": "att_cat_شهید صدرزاده"}],
            [{"text": "🏷️ طلاب و دانشجویان", "callback_data": "att_cat_طلاب و دانشجویان"}],
            [{"text": "✅ اتمام انتخاب دسته‌بندی", "callback_data": "att_cat_done"}],
            [{"text": "❌ انصراف", "callback_data": "cancel"}]
        ]
    }

def keyboard_attendance_users(session_id: str):
    attended_users = get_attended_users_for_session(session_id)
    all_users = get_users()
    keyboard = []
    for user in all_users:
        user_id_str = str(user['user_id'])
        if user_id_str in attended_users:
            emoji = "✅"
            callback = f"att_user_{user_id_str}_already"
        else:
            emoji = "◻️"
            callback = f"att_user_select_{session_id}_{user_id_str}"
        keyboard.append([
            {"text": f"{emoji} {user['name']} ({user['category']})",
             "callback_data": callback}
        ])
    keyboard.append([{"text": "✅ اتمام و ثبت", "callback_data": f"att_finish_{session_id}"}])
    keyboard.append([{"text": "❌ انصراف", "callback_data": "cancel"}])
    return {"inline_keyboard": keyboard}

# --------------------------- ADD USER CONVERSATION ---------------------------
async def start_add_user_conversation(bot: BaleBot, chat_id: int, user_id: int):
    user_states[user_id] = {"state": ConversationState.ADD_USER_GET_ID, "data": {}}
    await bot.send_message(chat_id,
        "👥 **اضافه کردن کاربر جدید**\n\n"
        "لطفاً آیدی عددی کاربر را در بله ارسال کنید:\n"
        "(برای انصراف /cancel را ارسال کنید)")

async def handle_add_user_id(bot: BaleBot, chat_id: int, user_id: int, text: str):
    if not text.isdigit():
        await bot.send_message(chat_id, "❌ آیدی باید یک عدد باشد! لطفاً دوباره وارد کنید:")
        return
    target_id = text.strip()
    idx, existing = find_user_row(target_id)
    if idx:
        await bot.send_message(chat_id, f"❌ کاربر با آیدی {target_id} قبلاً وجود دارد!")
        await cancel_conversation(bot, chat_id, user_id)
        return
    user_states[user_id]["data"]["target_id"] = target_id
    user_states[user_id]["state"] = ConversationState.ADD_USER_GET_CATEGORY
    await bot.send_message(chat_id,
        "✅ آیدی کاربر ثبت شد.\n\nلطفاً دسته کاربر را انتخاب کنید:",
        reply_markup=keyboard_categories())

async def handle_add_user_category(bot: BaleBot, chat_id: int, user_id: int, category: str):
    user_states[user_id]["data"]["category"] = category
    user_states[user_id]["state"] = ConversationState.ADD_USER_GET_NAME
    await bot.send_message(chat_id,
        f"✅ دسته کاربر: {category}\n\nلطفاً نام کامل کاربر را ارسال کنید:")

async def handle_add_user_name(bot: BaleBot, chat_id: int, user_id: int, name: str):
    data = user_states[user_id]["data"]
    add_user_to_sheet(data["target_id"], name, data["category"])
    del user_states[user_id]
    await bot.send_message(chat_id,
        f"✅ **کاربر با موفقیت اضافه شد!**\n\n"
        f"🆔 آیدی: {data['target_id']}\n"
        f"👤 نام: {name}\n"
        f"🏷️ دسته: {data['category']}\n\n"
        "برای بازگشت به پنل /panel را ارسال کنید.")

# --------------------------- SET DEBT CONVERSATION ---------------------------
async def start_set_debt_conversation(bot: BaleBot, chat_id: int, user_id: int):
    users = get_users()
    if not users:
        await bot.send_message(chat_id, "❌ هیچ کاربری در سیستم وجود ندارد!")
        return
    user_states[user_id] = {"state": ConversationState.SET_DEBT_SELECT_USER, "data": {}}
    await bot.send_message(chat_id,
        "💰 **تغییر بدهی کاربر**\n\nلطفاً کاربر مورد نظر را انتخاب کنید:",
        reply_markup=keyboard_users_list(users))

async def handle_set_debt_select_user(bot: BaleBot, chat_id: int, user_id: int, selected_user_id: str):
    idx, user_data = find_user_row(selected_user_id)
    user_states[user_id]["data"]["target_user_id"] = selected_user_id
    user_states[user_id]["data"]["target_user_name"] = user_data["name"]
    user_states[user_id]["data"]["current_debt"] = user_data.get("total_fee", 0)
    user_states[user_id]["state"] = ConversationState.SET_DEBT_GET_AMOUNT
    await bot.send_message(chat_id,
        f"👤 **{user_data['name']}**\n"
        f"🏷️ {user_data['category']}\n"
        f"💰 بدهی فعلی: {int(user_data.get('total_fee', 0)):,} تومان\n\n"
        "لطفاً مبلغ بدهی جدید را وارد کنید (فقط عدد):")

async def handle_set_debt_amount(bot: BaleBot, chat_id: int, user_id: int, amount_str: str):
    if not amount_str.isdigit():
        await bot.send_message(chat_id, "❌ مبلغ باید عدد باشد! لطفاً دوباره وارد کنید:")
        return
    amount = int(amount_str)
    data = user_states[user_id]["data"]
    update_user_fee(data["target_user_id"], amount)
    del user_states[user_id]
    await bot.send_message(chat_id,
        f"✅ **بدهی کاربر با موفقیت تغییر کرد!**\n\n"
        f"👤 {data['target_user_name']}\n"
        f"💰 بدهی جدید: {amount:,} تومان\n\n"
        "برای بازگشت به پنل /panel را ارسال کنید.")

# --------------------------- ADD SESSION CONVERSATION ---------------------------
async def start_add_session_conversation(bot: BaleBot, chat_id: int, user_id: int):
    user_states[user_id] = {
        "state": ConversationState.ADD_SESSION_GET_DATE,
        "data": {"selected_users": set(), "selected_categories": set()},
        "last_message_id": None
    }
    await bot.send_message(chat_id,
        "🎯 **اضافه کردن سانس جدید**\n\nلطفاً تاریخ سانس را انتخاب کنید (تاریخ شمسی):",
        reply_markup=keyboard_jalali_dates())

async def handle_add_session_date(bot: BaleBot, chat_id: int, user_id: int, date_str: str):
    if date_str == "manual":
        await bot.send_message(chat_id,
            "لطفاً تاریخ شمسی را به فرمت YYYY/MM/DD وارد کنید:\n"
            "مثال: ۱۴۰۳/۱۰/۱۵")
        return
    user_states[user_id]["data"]["date"] = date_str
    user_states[user_id]["state"] = ConversationState.ADD_SESSION_GET_NAME
    await bot.send_message(chat_id,
        f"✅ تاریخ: {date_str}\n\nلطفاً نام سانس را وارد کنید:\nمثال: سانس صبح")

async def handle_add_session_name(bot: BaleBot, chat_id: int, user_id: int, name: str):
    user_states[user_id]["data"]["name"] = name
    user_states[user_id]["state"] = ConversationState.ADD_SESSION_SELECT_CATEGORY
    await bot.send_message(chat_id,
        f"✅ نام سانس: {name}\n\nلطفاً دسته‌بندی شرکت‌کنندگان را انتخاب کنید:",
        reply_markup=keyboard_session_categories())

async def handle_add_session_select_category(bot: BaleBot, chat_id: int, user_id: int, category: str):
    data = user_states[user_id]["data"]
    if category not in data["selected_categories"]:
        data["selected_categories"].add(category)
    user_states[user_id]["state"] = ConversationState.ADD_SESSION_SELECT_USERS
    users = get_users_by_category(category)
    if not users:
        await bot.send_message(chat_id,
            f"❌ هیچ کاربری در دسته '{category}' وجود ندارد!\nلطفاً دسته دیگری انتخاب کنید.")
        return
    message_id = user_states[user_id].get("last_message_id")
    text = (f"🏷️ دسته: **{category}**\n"
            f"👥 تعداد کاربران: {len(users)} نفر\n\n"
            "✅ **کاربران را انتخاب کنید:**\n(برای انتخاب/لغو روی هر کاربر کلیک کنید)\n\n"
            "◻️ = انتخاب نشده\n✅ = انتخاب شده")
    if message_id:
        await bot.edit_message(chat_id, message_id, text,
                              reply_markup=keyboard_session_users(category, data["selected_users"]))
    else:
        result = await bot.send_message(chat_id, text,
                                       reply_markup=keyboard_session_users(category, data["selected_users"]))
        if result and result.get("result"):
            user_states[user_id]["last_message_id"] = result["result"]["message_id"]

async def handle_add_session_toggle_user(bot: BaleBot, chat_id: int, user_id: int, action: str, target_user_id: str):
    data = user_states[user_id]["data"]
    if action == "select":
        data["selected_users"].add(target_user_id)
    elif action == "unselect":
        if target_user_id in data["selected_users"]:
            data["selected_users"].remove(target_user_id)
    idx, user_data = find_user_row(target_user_id)
    if user_data:
        category = user_data['category']
        users_in_category = get_users_by_category(category)
        selected_in_category = sum(1 for u in users_in_category if str(u['user_id']) in data["selected_users"])
        message_id = user_states[user_id].get("last_message_id")
        text = (f"🏷️ دسته: **{category}**\n"
                f"✅ انتخاب‌شده در این دسته: {selected_in_category} از {len(users_in_category)} نفر\n"
                f"✅ مجموع انتخاب‌شده: {len(data['selected_users'])} نفر\n\n"
                "کاربران را انتخاب کنید:")
        if message_id:
            await bot.edit_message(chat_id, message_id, text,
                                  reply_markup=keyboard_session_users(category, data["selected_users"]))
        else:
            result = await bot.send_message(chat_id, text,
                                           reply_markup=keyboard_session_users(category, data["selected_users"]))
            if result and result.get("result"):
                user_states[user_id]["last_message_id"] = result["result"]["message_id"]

async def handle_add_session_select_all_category(bot: BaleBot, chat_id: int, user_id: int):
    data = user_states[user_id]["data"]
    if data["selected_categories"]:
        category = list(data["selected_categories"])[-1]
        users = get_users_by_category(category)
        for user in users:
            data["selected_users"].add(str(user['user_id']))
        message_id = user_states[user_id].get("last_message_id")
        text = (f"✅ همه کاربران دسته '{category}' انتخاب شدند!\n"
                f"تعداد: {len(users)} نفر\n"
                f"✅ مجموع انتخاب‌شده: {len(data['selected_users'])} نفر\n\n"
                "برای انتخاب دسته دیگر یا ادامه کلیک کنید.")
        if message_id:
            await bot.edit_message(chat_id, message_id, text,
                                  reply_markup=keyboard_session_users(category, data["selected_users"]))
        else:
            result = await bot.send_message(chat_id, text,
                                           reply_markup=keyboard_session_users(category, data["selected_users"]))
            if result and result.get("result"):
                user_states[user_id]["last_message_id"] = result["result"]["message_id"]

async def handle_add_session_clear_category(bot: BaleBot, chat_id: int, user_id: int):
    data = user_states[user_id]["data"]
    if data["selected_categories"]:
        category = list(data["selected_categories"])[-1]
        users_to_remove = []
        for uid in data["selected_users"]:
            idx, ud = find_user_row(uid)
            if ud and ud['category'] == category:
                users_to_remove.append(uid)
        for uid in users_to_remove:
            data["selected_users"].remove(uid)
        message_id = user_states[user_id].get("last_message_id")
        text = (f"❌ همه انتخاب‌های دسته '{category}' پاک شدند.\n"
                f"✅ مجموع انتخاب‌شده: {len(data['selected_users'])} نفر\n\n"
                "کاربران را انتخاب کنید:")
        if message_id:
            await bot.edit_message(chat_id, message_id, text,
                                  reply_markup=keyboard_session_users(category, data["selected_users"]))
        else:
            result = await bot.send_message(chat_id, text,
                                           reply_markup=keyboard_session_users(category, data["selected_users"]))
            if result and result.get("result"):
                user_states[user_id]["last_message_id"] = result["result"]["message_id"]

async def handle_add_session_another_category(bot: BaleBot, chat_id: int, user_id: int):
    user_states[user_id]["state"] = ConversationState.ADD_SESSION_SELECT_CATEGORY
    await bot.send_message(chat_id,
        "لطفاً دسته‌بندی دیگر را انتخاب کنید:",
        reply_markup=keyboard_session_categories())

async def handle_add_session_continue_price(bot: BaleBot, chat_id: int, user_id: int):
    data = user_states[user_id]["data"]
    if not data["selected_users"]:
        await bot.send_message(chat_id, "❌ هیچ کاربری انتخاب نشده است!")
        return
    user_states[user_id]["state"] = ConversationState.ADD_SESSION_GET_PRICE
    summary = f"✅ **خلاصه انتخاب‌ها:**\n\n"
    summary += f"📅 تاریخ: {data['date']}\n"
    summary += f"📝 نام سانس: {data['name']}\n"
    summary += f"🏷️ دسته‌بندی‌ها: {', '.join(data['selected_categories'])}\n"
    summary += f"👥 تعداد شرکت‌کنندگان: {len(data['selected_users'])} نفر\n\n"
    users_list = []
    for uid in data["selected_users"]:
        idx, ud = find_user_row(uid)
        if ud:
            users_list.append(f"• {ud['name']} ({ud['category']})")
    if users_list:
        summary += "📋 **لیست شرکت‌کنندگان:**\n" + "\n".join(users_list[:10])
        if len(users_list) > 10:
            summary += f"\n... و {len(users_list) - 10} نفر دیگر"
    summary += "\n\n💰 **لطفاً قیمت سانس را وارد کنید (فقط عدد):**\nمثال: 30000"
    await bot.send_message(chat_id, summary)

async def handle_add_session_price(bot: BaleBot, chat_id: int, user_id: int, price_str: str):
    if not price_str.isdigit():
        await bot.send_message(chat_id, "❌ قیمت باید عدد باشد! لطفاً دوباره وارد کنید:")
        return
    price = int(price_str)
    data = user_states[user_id]["data"]
    sessions = get_sessions()
    if sessions:
        session_ids = []
        for r in sessions:
            try:
                session_ids.append(int(r["session_id"]))
            except:
                pass
        next_id = max(session_ids) + 1 if session_ids else 1
    else:
        next_id = 1
    attended_users_str = ",".join(sorted(data["selected_users"]))
    add_session_to_sheet(next_id, data["name"], price, data["date"], attended_users_str)

    sent_messages = 0
    failed_messages = 0
    for uid in data["selected_users"]:
        idx, ud = find_user_row(uid)
        if idx:
            current_debt = int(ud.get("total_fee", 0))
            new_debt = current_debt + price
            set_user_fee(idx, new_debt)
            add_log(uid, next_id)
            try:
                await bot.send_message(
                    int(uid),
                    f"🎯 **شما در سانس جدید ثبت شدید**\n\n"
                    f"📝 نام سانس: {data['name']}\n"
                    f"📅 تاریخ: {data['date']}\n"
                    f"💰 قیمت سانس: {price:,} تومان\n\n"
                    f"💰 بدهی قبلی: {current_debt:,} تومان\n"
                    f"💰 بدهی جدید: {new_debt:,} تومان\n"
                    f"📈 افزایش بدهی: {price:,} تومان\n\n"
                    f"برای مشاهده و پرداخت بدهی از دستور /mydebt استفاده کنید."
                )
                sent_messages += 1
            except Exception as e:
                logger.error(f"خطا در ارسال پیام به کاربر {uid}: {e}")
                failed_messages += 1
    del user_states[user_id]
    result_text = (
        f"✅ **سانس با موفقیت اضافه شد!**\n\n"
        f"🆔 کد سانس: {next_id}\n"
        f"📝 نام: {data['name']}\n"
        f"📅 تاریخ: {data['date']}\n"
        f"💰 قیمت: {price:,} تومان\n"
        f"👥 تعداد شرکت‌کنندگان: {len(data['selected_users'])} نفر\n"
        f"💸 مجموع افزایش بدهی: {len(data['selected_users']) * price:,} تومان\n\n"
        f"📨 پیام‌های ارسالی: {sent_messages} موفق / {failed_messages} ناموفق\n"
        f"✅ بدهی هر کاربر به میزان {price:,} تومان افزایش یافت."
    )
    await bot.send_message(chat_id, result_text)

# --------------------------- ADD ATTENDANCE CONVERSATION ---------------------------
async def start_add_attendance_conversation(bot: BaleBot, chat_id: int, user_id: int):
    sessions = get_sessions()
    if not sessions:
        await bot.send_message(chat_id, "❌ هیچ سانسی وجود ندارد! ابتدا یک سانس اضافه کنید.")
        return
    user_states[user_id] = {
        "state": ConversationState.ADD_ATTENDANCE_SELECT_SESSION,
        "data": {"selected_users": set()}
    }
    await bot.send_message(chat_id,
        "📝 **ثبت حضور در سانس موجود**\n\n"
        "با این گزینه می‌توانید افراد جدید را به سانس‌های قبلی اضافه کنید.\n"
        "هزینه سانس به بدهی کاربر اضافه خواهد شد.\n\n"
        "لطفاً سانس مورد نظر را انتخاب کنید:",
        reply_markup=keyboard_sessions())

async def handle_add_attendance_select_session(bot: BaleBot, chat_id: int, user_id: int, session_id: str):
    idx, session_data = find_session_row(session_id)
    if not idx:
        await bot.send_message(chat_id, "❌ سانس مورد نظر یافت نشد!")
        await cancel_conversation(bot, chat_id, user_id)
        return
    user_states[user_id]["data"]["session_id"] = session_id
    user_states[user_id]["data"]["session_name"] = session_data["name"]
    user_states[user_id]["data"]["session_price"] = int(session_data.get("price", 0))
    user_states[user_id]["data"]["session_date"] = session_data["date"]
    user_states[user_id]["state"] = ConversationState.ADD_ATTENDANCE_SELECT_USERS
    attended_users = get_attended_users_for_session(session_id)
    text = (
        f"✅ سانس انتخاب شد:\n\n"
        f"📝 نام: {session_data['name']}\n"
        f"📅 تاریخ: {session_data['date']}\n"
        f"💰 قیمت: {session_data['price']:,} تومان\n"
        f"👥 حاضرین فعلی: {len(attended_users)} نفر\n\n"
        f"**لیست کاربران برای اضافه کردن به سانس:**\n"
        f"✅ = قبلاً در سانس ثبت شده\n"
        f"◻️ = می‌تواند انتخاب شود\n\n"
        f"کاربران مورد نظر را انتخاب کنید:"
    )
    await bot.send_message(chat_id, text, reply_markup=keyboard_attendance_users(session_id))

async def handle_add_attendance_select_user(bot: BaleBot, chat_id: int, user_id: int, session_id: str, target_user_id: str):
    data = user_states[user_id]["data"]
    attended_users = get_attended_users_for_session(session_id)
    if target_user_id in attended_users:
        await bot.answer_callback(f"کاربر قبلاً در این سانس ثبت شده است!")
        return
    if target_user_id in data["selected_users"]:
        data["selected_users"].remove(target_user_id)
    else:
        data["selected_users"].add(target_user_id)
    await bot.send_message(chat_id,
        f"✅ **انتخاب کاربران برای سانس:**\n"
        f"📝 نام سانس: {data['session_name']}\n"
        f"💰 قیمت: {data['session_price']:,} تومان\n"
        f"✅ انتخاب‌شده: {len(data['selected_users'])} نفر\n\n"
        f"کاربران مورد نظر را انتخاب کنید:",
        reply_markup=keyboard_attendance_users(session_id))

async def handle_add_attendance_finish(bot: BaleBot, chat_id: int, user_id: int, session_id: str):
    data = user_states[user_id]["data"]
    if not data["selected_users"]:
        await bot.send_message(chat_id, "❌ هیچ کاربری انتخاب نشده است!")
        return
    idx, session_data = find_session_row(session_id)
    if not idx:
        await bot.send_message(chat_id, "❌ سانس مورد نظر یافت نشد!")
        await cancel_conversation(bot, chat_id, user_id)
        return
    price = int(session_data.get("price", 0))
    attended_users = get_attended_users_for_session(session_id)
    new_attended_users = attended_users.copy()
    added_users = []
    failed_users = []
    for uid in data["selected_users"]:
        if uid not in new_attended_users:
            new_attended_users.add(uid)
            idx_user, user_data = find_user_row(uid)
            if idx_user:
                current_debt = int(user_data.get("total_fee", 0))
                new_debt = current_debt + price
                set_user_fee(idx_user, new_debt)
                add_log(uid, session_id)
                try:
                    await bot.send_message(
                        int(uid),
                        f"📝 **شما به سانس اضافه شدید**\n\n"
                        f"🎯 سانس: {session_data['name']}\n"
                        f"📅 تاریخ: {session_data['date']}\n"
                        f"💰 قیمت سانس: {price:,} تومان\n\n"
                        f"💰 بدهی قبلی: {current_debt:,} تومان\n"
                        f"💰 بدهی جدید: {new_debt:,} تومان\n"
                        f"📈 افزایش بدهی: {price:,} تومان\n\n"
                        f"برای مشاهده و پرداخت بدهی از دستور /mydebt استفاده کنید."
                    )
                    added_users.append(user_data['name'])
                except Exception as e:
                    logger.error(f"خطا در ارسال پیام به کاربر {uid}: {e}")
                    failed_users.append(user_data['name'] if user_data else uid)
            else:
                failed_users.append(uid)
    attended_users_str = ",".join(sorted(new_attended_users))
    update_session_attendance(session_id, attended_users_str)
    del user_states[user_id]
    result_text = (
        f"✅ **ثبت حضور با موفقیت انجام شد!**\n\n"
        f"🎯 سانس: {session_data['name']}\n"
        f"📅 تاریخ: {session_data['date']}\n"
        f"💰 قیمت: {price:,} تومان\n\n"
        f"👥 تعداد اضافه شده: {len(added_users)} نفر\n"
        f"💸 مجموع افزایش بدهی: {len(added_users) * price:,} تومان\n\n"
    )
    if added_users:
        result_text += f"✅ **کاربران اضافه شده:**\n"
        for name in added_users[:10]:
            result_text += f"• {name}\n"
        if len(added_users) > 10:
            result_text += f"... و {len(added_users) - 10} نفر دیگر\n"
    if failed_users:
        result_text += f"\n❌ **خطا در اضافه کردن:** {len(failed_users)} نفر\n"
    await bot.send_message(chat_id, result_text)

# --------------------------- HANDLE MYDEBT با پرداخت صورتحساب ---------------------------
async def handle_mydebt(bot: BaleBot, chat_id: int, user_id: int):
    idx, r = find_user_row(str(user_id))
    if not idx:
        await bot.send_message(chat_id, "شما در سیستم ثبت نشده‌اید. لطفا اول /start را بزنید.")
        return
    debt = int(r.get("total_fee", 0))
    name = r.get("name", "کاربر")
    category = r.get("category", "")
    if debt <= 0:
        await bot.send_message(chat_id,
            f"👤 **{name}**\n🏷️ {category}\n\n✅ شما هیچ بدهی ندارید!\n💰 بدهی فعلی: ۰ تومان")
        return
    title = f"پرداخت بدهی باشگاه - {name}"
    description = f"پرداخت بدهی باشگاه ورزشی\nکاربر: {name}\nدسته: {category}"
    payload = f"debt_payment_{user_id}_{int(datetime.now().timestamp())}"
    prices = [{"label": f"بدهی باشگاه ({debt:,} تومان)", "amount": debt * 10}]  # ریال
    result = await bot.send_invoice(chat_id=chat_id, title=title, description=description,
                                    payload=payload, prices=prices)
    if not result or not result.get("ok"):
        await bot.send_message(chat_id,
            f"❌ خطا در ایجاد درخواست پرداخت!\n\n👤 {name}\n🏷️ {category}\n💰 بدهی شما: {debt:,} تومان\n\nلطفاً با مدیر سیستم تماس بگیرید.")
        logger.error(f"خطا در ارسال صورتحساب: {result}")

# --------------------------- HANDLE PRE CHECKOUT QUERY ---------------------------
async def handle_pre_checkout_query(bot: BaleBot, pre_checkout_query: dict):
    query_id = pre_checkout_query["id"]
    user_id = pre_checkout_query["from"]["id"]
    amount = pre_checkout_query["total_amount"] / 10  # تومان
    payload = pre_checkout_query.get("invoice_payload", "")
    logger.info(f"pre_checkout_query از کاربر {user_id} به مبلغ {amount} تومان")
    idx, user_data = find_user_row(str(user_id))
    if not idx:
        await bot.answer_pre_checkout_query(query_id, False, "شما در سیستم ثبت نشده‌اید. لطفاً ابتدا /start را بزنید.")
        return
    current_debt = int(user_data.get("total_fee", 0))
    if current_debt <= 0:
        await bot.answer_pre_checkout_query(query_id, False, "شما هیچ بدهی ندارید!")
        return
    if abs(amount - current_debt) > 100:
        await bot.answer_pre_checkout_query(query_id, False,
                                            f"مبلغ پرداختی با بدهی شما ({current_debt:,} تومان) مطابقت ندارد.")
        return
    await bot.answer_pre_checkout_query(query_id, True)
    logger.info(f"pre_checkout_query تأیید شد")

# --------------------------- HANDLE SUCCESSFUL PAYMENT (با لاگ کامل) ---------------------------
async def handle_successful_payment(bot: BaleBot, chat_id: int, user_id: int, successful_payment: dict):
    total_amount = successful_payment["total_amount"] / 10
    payment_id = successful_payment.get("telegram_payment_charge_id", "")

    logger.info("="*50)
    logger.info(f"💰 پرداخت موفق دریافت شد:")
    logger.info(f"   کاربر: {user_id}")
    logger.info(f"   مبلغ: {total_amount} تومان")
    logger.info(f"   شناسه پرداخت: {payment_id}")
    logger.info("="*50)

    # مرحله 1: پیدا کردن کاربر
    logger.info("مرحله 1: جستجوی کاربر در شیت...")
    idx, user_data = find_user_row(str(user_id))
    if not idx:
        logger.error(f"کاربر {user_id} در شیت users یافت نشد!")
        await bot.send_message(chat_id,
            "❌ خطا در پردازش پرداخت! (کاربر یافت نشد)\nلطفاً با مدیر سیستم تماس بگیرید.")
        return
    logger.info(f"کاربر در سطر {idx} یافت شد: {user_data}")

    # مرحله 2: محاسبه بدهی جدید
    current_debt = int(user_data.get("total_fee", 0))
    new_debt = max(0, current_debt - total_amount)
    logger.info(f"مرحله 2: بدهی قبلی={current_debt}، بدهی جدید={new_debt}")

    # مرحله 3: به‌روزرسانی بدهی در شیت users
    logger.info(f"مرحله 3: به‌روزرسانی سطر {idx} با مقدار {new_debt}")
    try:
        set_user_fee(idx, new_debt)
        logger.info("✅ بدهی با موفقیت به‌روز شد")
    except Exception as e:
        logger.error(f"خطا در به‌روزرسانی بدهی: {e}")

    # مرحله 4: ثبت در شیت payments
    logger.info("مرحله 4: ثبت تراکنش در شیت payments")
    try:
        ws_pay = sheet_payments()
        ws_pay.append_row([
            datetime.now().isoformat(),
            str(user_id),
            user_data.get('name', ''),
            total_amount,
            payment_id,
            'completed'
        ])
        logger.info("✅ تراکنش در شیت payments ثبت شد")
    except Exception as e:
        logger.error(f"خطا در ثبت شیت payments: {e}")

    # مرحله 5: ارسال پیام تأیید به کاربر
    logger.info("مرحله 5: ارسال پیام تأیید به کاربر")
    confirm_text = (
        f"✅ **پرداخت با موفقیت انجام شد!**\n\n"
        f"💰 مبلغ پرداختی: {total_amount:,} تومان\n"
        f"💳 کد پیگیری: {payment_id}\n\n"
        f"👤 {user_data['name']}\n"
        f"🏷️ {user_data['category']}\n"
        f"💰 بدهی قبلی: {current_debt:,} تومان\n"
        f"💰 بدهی جدید: {new_debt:,} تومان\n"
        f"📉 کاهش: {total_amount:,} تومان\n\n"
        f"✅ تغییرات با موفقیت در سیستم ثبت شد."
    )
    try:
        await bot.send_message(chat_id, confirm_text)
        logger.info("✅ پیام تأیید ارسال شد")
    except Exception as e:
        logger.error(f"خطا در ارسال پیام تأیید: {e}")

    logger.info("پردازش پرداخت موفق با موفقیت پایان یافت")

# --------------------------- MAIN HANDLERS ---------------------------
async def handle_start(bot: BaleBot, chat_id: int, user_id: int, user_name: str):
    idx, user_data = find_user_row(str(user_id))
    if not idx:
        keyboard = {
            "inline_keyboard": [
                [{"text": "🏷️ شهید فهمیده", "callback_data": "reg_شهید فهمیده"}],
                [{"text": "🏷️ شهید دانشگر", "callback_data": "reg_شهید دانشگر"}],
                [{"text": "🏷️ شهید صدرزاده", "callback_data": "reg_شهید صدرزاده"}],
                [{"text": "🏷️ طلاب و دانشجویان", "callback_data": "reg_طلاب و دانشجویان"}]
            ]
        }
        text = (
            f"سلام {user_name} 🌟\n"
            "به ربات مدیریت باشگاه ورزشی خوش آمدید!\n\n"
            "لطفاً دسته خود را انتخاب کنید:\n\n"
            "📱 **همیشه در دسترس**\n"
            "برای مشاهده امکانات /help را بزنید"
        )
        await bot.send_message(chat_id, text, reply_markup=keyboard)
    else:
        text = (
            f"سلام {user_data['name']} 🌟\n"
            f"🏷️ دسته: {user_data['category']}\n\n"
            "📱 **همیشه در دسترس**\n\n"
            "🔹 /mydebt - مشاهده و پرداخت بدهی\n"
            "🔹 /panel - پنل مدیریت (مدیران)\n"
            "🔹 /help - راهنما\n\n"
            "برای شروع کار از دستورات بالا استفاده کنید."
        )
        await bot.send_message(chat_id, text)

async def handle_panel(bot: BaleBot, chat_id: int, user_id: int):
    if not is_admin(user_id):
        await bot.send_message(chat_id,
            "❌ شما دسترسی مدیریتی ندارید!\n\n"
            "فقط مدیران می‌توانند از پنل استفاده کنند.")
        return
    keyboard = keyboard_panel()
    await bot.send_message(chat_id,
        "👨‍💼 **پنل مدیریت**\n\n"
        "لطفاً عملیات مورد نظر را انتخاب کنید:\n\n"
        "🔹 برای انصراف از هر عملیات /cancel را بزنید",
        reply_markup=keyboard)

async def handle_help(bot: BaleBot, chat_id: int, user_id: int):
    idx, user_data = find_user_row(str(user_id))
    if idx and is_admin(user_id):
        text = (
            "📱 **راهنمای ربات مدیریت باشگاه**\n\n"
            "👤 **کاربران عادی:**\n"
            "🔹 /start - ثبت نام/ورود\n"
            "🔹 /mydebt - مشاهده و پرداخت بدهی\n"
            "🔹 /help - نمایش این راهنما\n\n"
            "👨‍💼 **مدیران:**\n"
            "🔹 /panel - پنل مدیریت\n"
            "🔹 /adduser - اضافه کردن کاربر\n"
            "🔹 /setdebt - تغییر بدهی کاربر\n"
            "🔹 /addsession - اضافه کردن سانس\n"
            "🔹 /addattendance - ثبت حضور سانس موجود\n\n"
            "💰 **سیستم پرداخت:**\n"
            "• پرداخت از طریق کیف پول بله\n"
            "• صورتحساب الکترونیکی\n"
            "• رسید خودکار\n\n"
            "📞 **پشتیبانی:**\n"
            "برای گزارش مشکل با مدیر سیستم تماس بگیرید."
        )
    else:
        text = (
            "📱 **راهنمای ربات مدیریت باشگاه**\n\n"
            "🔹 /start - ثبت نام/ورود\n"
            "🔹 /mydebt - مشاهده و پرداخت بدهی\n"
            "🔹 /help - نمایش این راهنما\n\n"
            "💰 پرداخت از طریق کیف پول بله\n"
            "📞 برای اطلاعات بیشتر با مدیر سیستم تماس بگیرید."
        )
    await bot.send_message(chat_id, text)

async def cancel_conversation(bot: BaleBot, chat_id: int, user_id: int):
    if user_id in user_states:
        del user_states[user_id]
    await bot.send_message(chat_id,
        "❌ عملیات لغو شد.\n\n"
        "برای بازگشت به پنل /panel را ارسال کنید.")

async def handle_callback(bot: BaleBot, callback: dict):
    callback_id = callback["id"]
    user_id = callback["from"]["id"]
    chat_id = callback["message"]["chat"]["id"]
    message_id = callback["message"]["message_id"]
    data = callback["data"]

    await bot.answer_callback(callback_id)

    if data.startswith("reg_"):
        category = data.replace("reg_", "")
        user_name = callback["from"].get("first_name", "کاربر")
        idx, existing = find_user_row(str(user_id))
        if idx:
            await bot.edit_message(chat_id, message_id,
                f"👤 **شما قبلاً ثبت‌نام کرده‌اید!**\n\n"
                f"نام: {existing['name']}\n"
                f"دسته: {existing['category']}\n\n"
                f"برای مشاهده بدهی /mydebt را بزنید.")
            return
        add_user_to_sheet(str(user_id), user_name, category)
        await bot.edit_message(chat_id, message_id,
            f"✅ **ثبت‌نام با موفقیت انجام شد!**\n\n"
            f"👤 نام: {user_name}\n"
            f"🏷️ دسته: {category}\n\n"
            f"🔹 /mydebt - مشاهده و پرداخت بدهی\n"
            f"🔹 /help - راهنما")

    elif data == "add_user":
        await start_add_user_conversation(bot, chat_id, user_id)
    elif data == "set_debt":
        await start_set_debt_conversation(bot, chat_id, user_id)
    elif data == "add_session":
        await start_add_session_conversation(bot, chat_id, user_id)
    elif data == "add_attendance":
        await start_add_attendance_conversation(bot, chat_id, user_id)
    elif data.startswith("cat_"):
        category = data.replace("cat_", "")
        if user_id in user_states and user_states[user_id]["state"] == ConversationState.ADD_USER_GET_CATEGORY:
            await handle_add_user_category(bot, chat_id, user_id, category)
    elif data.startswith("jalali_date_"):
        date_str = data.replace("jalali_date_", "")
        if user_id in user_states and user_states[user_id]["state"] == ConversationState.ADD_SESSION_GET_DATE:
            await handle_add_session_date(bot, chat_id, user_id, date_str)
    elif data.startswith("user_"):
        selected_user_id = data.replace("user_", "")
        if user_id in user_states and user_states[user_id]["state"] == ConversationState.SET_DEBT_SELECT_USER:
            await handle_set_debt_select_user(bot, chat_id, user_id, selected_user_id)
    elif data.startswith("session_cat_"):
        if data == "session_cat_done":
            if user_id in user_states:
                data_state = user_states[user_id]["data"]
                if not data_state["selected_categories"]:
                    await bot.send_message(chat_id, "❌ حداقل یک دسته‌بندی انتخاب کنید!")
                    return
                await handle_add_session_continue_price(bot, chat_id, user_id)
        else:
            category = data.replace("session_cat_", "")
            if user_id in user_states and user_states[user_id]["state"] in [ConversationState.ADD_SESSION_SELECT_CATEGORY,
                                                                           ConversationState.ADD_SESSION_SELECT_USERS]:
                await handle_add_session_select_category(bot, chat_id, user_id, category)
    elif data.startswith("session_user_select_"):
        target_user_id = data.replace("session_user_select_", "")
        if user_id in user_states and user_states[user_id]["state"] == ConversationState.ADD_SESSION_SELECT_USERS:
            await handle_add_session_toggle_user(bot, chat_id, user_id, "select", target_user_id)
    elif data.startswith("session_user_unselect_"):
        target_user_id = data.replace("session_user_unselect_", "")
        if user_id in user_states and user_states[user_id]["state"] == ConversationState.ADD_SESSION_SELECT_USERS:
            await handle_add_session_toggle_user(bot, chat_id, user_id, "unselect", target_user_id)
    elif data == "session_select_all_category":
        if user_id in user_states and user_states[user_id]["state"] == ConversationState.ADD_SESSION_SELECT_USERS:
            await handle_add_session_select_all_category(bot, chat_id, user_id)
    elif data == "session_clear_category":
        if user_id in user_states and user_states[user_id]["state"] == ConversationState.ADD_SESSION_SELECT_USERS:
            await handle_add_session_clear_category(bot, chat_id, user_id)
    elif data == "session_another_category":
        if user_id in user_states and user_states[user_id]["state"] == ConversationState.ADD_SESSION_SELECT_USERS:
            await handle_add_session_another_category(bot, chat_id, user_id)
    elif data == "session_continue_price":
        if user_id in user_states and user_states[user_id]["state"] == ConversationState.ADD_SESSION_SELECT_USERS:
            await handle_add_session_continue_price(bot, chat_id, user_id)
    elif data.startswith("session_"):
        session_id = data.replace("session_", "")
        if user_id in user_states and user_states[user_id]["state"] == ConversationState.ADD_ATTENDANCE_SELECT_SESSION:
            await handle_add_attendance_select_session(bot, chat_id, user_id, session_id)
    elif data.startswith("att_user_select_"):
        parts = data.split("_")
        if len(parts) >= 5:
            session_id = parts[3]
            target_user_id = parts[4]
            if user_id in user_states and user_states[user_id]["state"] == ConversationState.ADD_ATTENDANCE_SELECT_USERS:
                await handle_add_attendance_select_user(bot, chat_id, user_id, session_id, target_user_id)
    elif data.startswith("att_finish_"):
        session_id = data.replace("att_finish_", "")
        if user_id in user_states and user_states[user_id]["state"] == ConversationState.ADD_ATTENDANCE_SELECT_USERS:
            await handle_add_attendance_finish(bot, chat_id, user_id, session_id)
    elif data == "list_debts":
        users = get_users()
        if not users:
            await bot.edit_message(chat_id, message_id, "❌ هیچ کاربری ثبت نشده است!")
            return
        users_with_debt = [u for u in users if int(u.get("total_fee", 0)) > 0]
        if not users_with_debt:
            await bot.edit_message(chat_id, message_id, "✅ هیچ بدهکاری وجود ندارد!")
            return
        text = "📋 **لیست بدهکاران**\n\n"
        total_debt = 0
        for user in users_with_debt:
            debt = int(user.get("total_fee", 0))
            total_debt += debt
            text += (
                f"👤 **{user['name']}**\n"
                f"🏷️ {user['category']}\n"
                f"💰 {debt:,} تومان\n"
                f"――――――――――――\n"
            )
        text += f"\n💰 **مجموع بدهی‌ها: {total_debt:,} تومان**"
        await bot.edit_message(chat_id, message_id, text)
    elif data == "back_to_start":
        user_name = callback["from"].get("first_name", "کاربر")
        await handle_start(bot, chat_id, user_id, user_name)
    elif data == "cancel":
        await cancel_conversation(bot, chat_id, user_id)

async def handle_message(bot: BaleBot, message: dict):
    chat_id = message["chat"]["id"]
    user_id = message["from"]["id"]
    text = message.get("text", "").strip()

    if text.lower() in ["/cancel", "انصراف", "لغو"]:
        await cancel_conversation(bot, chat_id, user_id)
        return

    if user_id in user_states:
        state = user_states[user_id]["state"]
        if state == ConversationState.ADD_USER_GET_ID:
            await handle_add_user_id(bot, chat_id, user_id, text)
        elif state == ConversationState.ADD_USER_GET_NAME:
            await handle_add_user_name(bot, chat_id, user_id, text)
        elif state == ConversationState.SET_DEBT_GET_AMOUNT:
            await handle_set_debt_amount(bot, chat_id, user_id, text)
        elif state == ConversationState.ADD_SESSION_GET_NAME:
            await handle_add_session_name(bot, chat_id, user_id, text)
        elif state == ConversationState.ADD_SESSION_GET_PRICE:
            await handle_add_session_price(bot, chat_id, user_id, text)
        elif state == ConversationState.ADD_SESSION_GET_DATE and "/" in text and text.count("/") == 2:
            user_states[user_id]["data"]["date"] = text
            user_states[user_id]["state"] = ConversationState.ADD_SESSION_GET_NAME
            await bot.send_message(chat_id,
                f"✅ تاریخ: {text}\n\nلطفاً نام سانس را وارد کنید:")
        else:
            await bot.send_message(chat_id,
                "❌ ورودی نامعتبر!\nلطفاً دوباره تلاش کنید یا /cancel را بزنید.")
        return

    if not text.startswith("/"):
        await bot.send_message(chat_id,
            "🤖 **ربات مدیریت باشگاه ورزشی**\n\n"
            "برای شروع از دستورات زیر استفاده کنید:\n"
            "🔹 /start - ثبت نام/ورود\n"
            "🔹 /mydebt - مشاهده و پرداخت بدهی\n"
            "🔹 /panel - پنل مدیریت\n"
            "🔹 /help - راهنما")
        return

    parts = text.split()
    command = parts[0].lower()

    if command == "/start":
        user_name = message["from"].get("first_name", "کاربر")
        await handle_start(bot, chat_id, user_id, user_name)
    elif command == "/mydebt":
        await handle_mydebt(bot, chat_id, user_id)
    elif command == "/panel":
        await handle_panel(bot, chat_id, user_id)
    elif command == "/help":
        await handle_help(bot, chat_id, user_id)
    elif command == "/adduser":
        await start_add_user_conversation(bot, chat_id, user_id)
    elif command == "/setdebt":
        await start_set_debt_conversation(bot, chat_id, user_id)
    elif command == "/addsession":
        await start_add_session_conversation(bot, chat_id, user_id)
    elif command == "/addattendance":
        await start_add_attendance_conversation(bot, chat_id, user_id)
    else:
        await bot.send_message(chat_id,
            "❌ دستور ناشناخته!\n\n"
            "برای مشاهده دستورات معتبر /help را بزنید.")

# --------------------------- MAIN POLLING ---------------------------
async def poll_updates(bot: BaleBot):
    url = f"{bot.base_url}/getUpdates"
    offset = 0
    while True:
        try:
            params = {"offset": offset, "timeout": 30}
            async with bot.session.get(url, params=params) as response:
                data = await response.json()
                if data.get("ok"):
                    updates = data.get("result", [])
                    for update in updates:
                        offset = update["update_id"] + 1
                        if "message" in update:
                            if "successful_payment" in update["message"]:
                                chat_id = update["message"]["chat"]["id"]
                                user_id = update["message"]["from"]["id"]
                                await handle_successful_payment(bot, chat_id, user_id, update["message"]["successful_payment"])
                            else:
                                await handle_message(bot, update["message"])
                        elif "callback_query" in update:
                            await handle_callback(bot, update["callback_query"])
                        elif "pre_checkout_query" in update:
                            await handle_pre_checkout_query(bot, update["pre_checkout_query"])
                await asyncio.sleep(0.1)
        except Exception as e:
            logger.error(f"خطا در دریافت به‌روزرسانی‌ها: {e}")
            await asyncio.sleep(5)

async def main_async():
    if not BALE_TOKEN:
        print("❌ BALE_TOKEN یافت نشد!")
        return
    print(f"✅ توکن: {BALE_TOKEN[:15]}...")
    print("🤖 راه‌اندازی ربات بله...")
    bot = BaleBot(BALE_TOKEN)
    try:
        await bot.start()
        print("✅ Session شروع شد")
        test_url = f"{bot.base_url}/getMe"
        async with bot.session.get(test_url) as response:
            data = await response.json()
            if data.get("ok"):
                bot_info = data["result"]
                print(f"✅ متصل شد: {bot_info.get('first_name')}")
                print(f"👤 @{bot_info.get('username')}")
                if PAYMENT_WALLET_TOKEN:
                    if PAYMENT_WALLET_TOKEN.startswith("WALLET-TEST"):
                        print("💰 وضعیت پرداخت: تستی")
                    else:
                        print("💰 وضعیت پرداخت: واقعی")
                else:
                    print("⚠️  توکن پرداخت تنظیم نشده! از توکن تست استفاده می‌شود")
            else:
                print(f"❌ خطای اتحاد: {data}")
                return
        print("🔄 شروع دریافت پیام‌ها...")
        await poll_updates(bot)
    except KeyboardInterrupt:
        print("\n🤖 ربات متوقف شد")
    except Exception as e:
        print(f"❌ خطا: {e}")
    finally:
        await bot.stop()
        print("✅ Session بسته شد")

def main():
    asyncio.run(main_async())

if __name__ == "__main__":
    main()