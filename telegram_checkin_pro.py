# telegram_checkin_pro_v3.py
# 完整版（aiogram 3）：
# - 打卡 / 休息 / 回座统计 / 超时提醒
# - 多管理员设置面板（多 ID）
# - 管理员日志（写入 admin_logs）
# - 自动/手动 报表（Excel .xlsx，中文文件名，带群名）
# - 自动在首次使用时为群插入 settings 初始行
#
# 依赖:
# pip install aiogram==3.1.0 aiosqlite python-dotenv openpyxl apscheduler

import asyncio
import aiosqlite
import io
import os
import re
import logging
from datetime import datetime, timedelta, date, time
from typing import Optional

from dotenv import load_dotenv
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardMarkup, InlineKeyboardButton, InputFile
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.cron import CronTrigger

# ---------------------------
# 配置
# ---------------------------
load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN")
ADMIN_IDS = [int(x) for x in os.getenv("ADMIN_IDS", "").replace(" ", "").split(",") if x]

if not BOT_TOKEN:
    raise RuntimeError("请在 .env 中设置 BOT_TOKEN")

DB_PATH = "checkin_pro.db"
LOCAL_OFFSET = timedelta(hours=7)   # 印尼时区，可改
DAILY_REPORT_HOUR = 10
WEEKLY_REPORT_DAY = 0
WEEKLY_REPORT_HOUR = 10
MONTHLY_REPORT_DAY = 1
MONTHLY_REPORT_HOUR = 10
OVERTIME_REMINDER_INTERVAL = 3  # 分钟

BREAK_LIMITS = {
    "toilet_small": 5,
    "toilet_big": 10,
    "smoke": 5,
    "meal": 30,
}

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
pending_media_for_chat = {}  # chat_id -> state string

# ---------------------------
# DB 初始化（含 admin_logs）
# ---------------------------
async def init_db():
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute("""
            CREATE TABLE IF NOT EXISTS work_sessions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                chat_id INTEGER,
                start_time TEXT,
                end_time TEXT
            )
        """)
        await db.execute("""
            CREATE TABLE IF NOT EXISTS break_sessions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                chat_id INTEGER,
                type TEXT,
                start_time TEXT,
                end_time TEXT
            )
        """)
        await db.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                chat_id INTEGER PRIMARY KEY,
                reminder_text TEXT,
                reminder_media_file_id TEXT,
                weekly_report_enabled INTEGER DEFAULT 0,
                monthly_report_enabled INTEGER DEFAULT 0
            )
        """)
        await db.execute("""
            CREATE TABLE IF NOT EXISTS admin_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                chat_id INTEGER,
                admin_id INTEGER,
                action TEXT,
                details TEXT,
                created_at TEXT
            )
        """)
        await db.commit()
    logger.info("数据库初始化完成。")

# ---------------------------
# 时间与格式工具
# ---------------------------
def now_utc() -> datetime:
    return datetime.utcnow()

def now_local() -> datetime:
    return datetime.utcnow() + LOCAL_OFFSET

def to_str(dt: datetime) -> str:
    return dt.strftime("%Y-%m-%d %H:%M:%S")

def parse_str(s: Optional[str]) -> Optional[datetime]:
    if not s:
        return None
    try:
        return datetime.strptime(s, "%Y-%m-%d %H:%M:%S")
    except:
        return None

def fmt_hm_local(dt_utc: Optional[datetime]) -> str:
    if not dt_utc:
        return "-"
    return (dt_utc + LOCAL_OFFSET).strftime("%H:%M")

def today_local_date() -> date:
    return (datetime.utcnow() + LOCAL_OFFSET).date()

def minutes_between(a: Optional[datetime], b: Optional[datetime]) -> int:
    if not a or not b:
        return 0
    return max(0, int((b - a).total_seconds() // 60))

def fmt_minutes(m: int) -> str:
    if m >= 60:
        h = m // 60
        mm = m % 60
        return f"{h}小时{mm}分钟"
    return f"{m}分钟"

# ---------------------------
# 设置/日志辅助
# ---------------------------
async def ensure_settings(chat_id: int):
    """确保 settings 表存在该 chat_id 的行（首次使用自动插入）"""
    async with aiosqlite.connect(DB_PATH) as db:
        cur = await db.execute("SELECT 1 FROM settings WHERE chat_id = ?", (chat_id,))
        found = await cur.fetchone()
        if not found:
            await db.execute("INSERT INTO settings (chat_id, reminder_text, reminder_media_file_id, weekly_report_enabled, monthly_report_enabled) VALUES (?, ?, ?, 0, 0)",
                             (chat_id, None, None))
            await db.commit()

async def set_chat_setting(chat_id: int, key: str, value):
    await ensure_settings(chat_id)
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(f"UPDATE settings SET {key} = ? WHERE chat_id = ?", (value, chat_id))
        await db.commit()

async def get_chat_settings(chat_id: int):
    await ensure_settings(chat_id)
    async with aiosqlite.connect(DB_PATH) as db:
        cur = await db.execute("SELECT reminder_text, reminder_media_file_id, weekly_report_enabled, monthly_report_enabled FROM settings WHERE chat_id = ?", (chat_id,))
        row = await cur.fetchone()
    return {"reminder_text": row[0], "reminder_media_file_id": row[1], "weekly_report_enabled": row[2], "monthly_report_enabled": row[3]}

async def get_chats_with_setting_enabled(col_name: str):
    async with aiosqlite.connect(DB_PATH) as db:
        cur = await db.execute(f"SELECT chat_id FROM settings WHERE {col_name} = 1")
        rows = await cur.fetchall()
    return [r[0] for r in rows]

async def log_admin_action(chat_id: int, admin_id: int, action: str, details: str = ""):
    created_at = to_str(now_utc())
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute("INSERT INTO admin_logs (chat_id, admin_id, action, details, created_at) VALUES (?, ?, ?, ?, ?)",
                         (chat_id, admin_id, action, details, created_at))
        await db.commit()

# ---------------------------
# 打卡 / 休息 数据写入（均确保 settings 存在）
# ---------------------------
async def start_work(user_id: int, chat_id: int):
    await ensure_settings(chat_id)
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute("INSERT INTO work_sessions (user_id, chat_id, start_time) VALUES (?, ?, ?)",
                         (user_id, chat_id, to_str(now_utc())))
        await db.commit()

async def end_work(user_id: int, chat_id: int):
    await ensure_settings(chat_id)
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute("UPDATE work_sessions SET end_time = ? WHERE user_id=? AND chat_id=? AND end_time IS NULL",
                         (to_str(now_utc()), user_id, chat_id))
        await db.commit()

async def start_break(user_id: int, chat_id: int, btype: str):
    await ensure_settings(chat_id)
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute("INSERT INTO break_sessions (user_id, chat_id, type, start_time) VALUES (?, ?, ?, ?)",
                         (user_id, chat_id, btype, to_str(now_utc())))
        await db.commit()

async def end_break(user_id: int, chat_id: int):
    await ensure_settings(chat_id)
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute("UPDATE break_sessions SET end_time = ? WHERE user_id=? AND chat_id=? AND end_time IS NULL",
                         (to_str(now_utc()), user_id, chat_id))
        await db.commit()

# ---------------------------
# 菜单
# ---------------------------
def get_menu():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="🏁 上班打卡"), KeyboardButton(text="🏁 下班签退")],
            [KeyboardButton(text="🚶‍♂️ 小厕开始"), KeyboardButton(text="🚽 大厕开始")],
            [KeyboardButton(text="🚬 抽烟开始"), KeyboardButton(text="🍱 吃饭开始")],
            [KeyboardButton(text="💺 回座"), KeyboardButton(text="📊 今日统计")],
            [KeyboardButton(text="📈 排行榜"), KeyboardButton(text="⚙️ 设置")]
        ], resize_keyboard=True
    )

# ---------------------------
# Handlers: 基本交互
# ---------------------------
@dp.message(Command("start"))
async def cmd_start(message: types.Message):
    await ensure_settings(message.chat.id)
    await message.reply("欢迎使用打卡机器人。请通过菜单进行操作。", reply_markup=get_menu())

@dp.message(F.text == "🏁 上班打卡")
async def handler_start_work(message: types.Message):
    await start_work(message.from_user.id, message.chat.id)
    await message.reply(f"✅ 上班打卡成功！（{fmt_hm_local(now_utc())}）", reply_markup=get_menu())

@dp.message(F.text == "🏁 下班签退")
async def handler_end_work(message: types.Message):
    await end_work(message.from_user.id, message.chat.id)
    await message.reply(f"🕒 下班签退成功。（{fmt_hm_local(now_utc())}）", reply_markup=get_menu())

BREAK_LABELS = {
    "🚶‍♂️ 小厕开始": "toilet_small",
    "🚽 大厕开始": "toilet_big",
    "🚬 抽烟开始": "smoke",
    "🍱 吃饭开始": "meal"
}

@dp.message(F.text.in_(list(BREAK_LABELS.keys())))
async def handler_start_break(message: types.Message):
    btype = BREAK_LABELS[message.text]
    await start_break(message.from_user.id, message.chat.id, btype)
    limit = BREAK_LIMITS.get(btype, 5)
    settings = await get_chat_settings(message.chat.id)
    rtext = settings.get("reminder_text") or f"你已开始 {message.text.replace('开始','')} ，预计 {limit} 分钟。"
    await message.reply(f"{rtext}\n⏰ 开始时间：{fmt_hm_local(now_utc())}", reply_markup=get_menu())
    asyncio.create_task(break_overtime_watcher(message.from_user.id, message.chat.id, btype, now_utc()))

@dp.message(F.text == "💺 回座")
async def handler_return_seat(message: types.Message):
    user_id = message.from_user.id
    chat_id = message.chat.id
    now = now_utc()

    async with aiosqlite.connect(DB_PATH) as db:
        cur = await db.execute(
            "SELECT id, type, start_time FROM break_sessions WHERE user_id=? AND chat_id=? AND end_time IS NULL ORDER BY id DESC LIMIT 1",
            (user_id, chat_id)
        )
        row = await cur.fetchone()

    if not row:
        await message.reply(f"💼 欢迎回来！（{fmt_hm_local(now)}）", reply_markup=get_menu())
        return

    _, btype, start_s = row
    sdt = parse_str(start_s)
    used_mins = minutes_between(sdt, now)
    human = {"toilet_small": "小厕", "toilet_big": "大厕", "smoke": "抽烟", "meal": "吃饭"}.get(btype, btype)

    await end_break(user_id, chat_id)

    today = today_local_date()
    summary = await compute_daily_summary(user_id, chat_id, today)
    total_times = summary["total_leave_times"]
    total_minutes = summary["total_leave_minutes"]

    msg = (
        f"💼 欢迎回来！\n"
        f"🚶‍♂️ 本次 {human} 用时：{used_mins} 分钟\n"
        f"📅 今日累计离开 {total_times} 次，共 {fmt_minutes(total_minutes)}\n"
        f"（{fmt_hm_local(sdt)} ~ {fmt_hm_local(now)}）"
    )

    await message.reply(msg, reply_markup=get_menu())

@dp.message(Command("leaderboard"))
@dp.message(F.text == "📈 排行榜")
async def cmd_leaderboard(message: types.Message):
    chat_id = message.chat.id
    async with aiosqlite.connect(DB_PATH) as db:
        cur = await db.execute("SELECT DISTINCT user_id FROM work_sessions WHERE chat_id = ?", (chat_id,))
        rows = await cur.fetchall()
    users = [r[0] for r in rows]
    today = today_local_date()
    entries = []
    for uid in users:
        works, breaks = await get_day_intervals_for_user_in_chat(uid, chat_id, today)
        total_work = sum(minutes_between(s, e or now_utc()) for s, e in works if s)
        total_break = sum(minutes_between(s, e or now_utc()) for _, s, e in breaks if s)
        entries.append((uid, total_work - total_break, total_break))
    entries.sort(key=lambda x: x[1], reverse=True)
    lines = [f"🏆 本群今日排行榜（{today.isoformat()}）"]
    if not entries:
        lines.append("暂无数据。")
    else:
        pos = 1
        for uid, net_m, break_m in entries[:10]:
            try:
                member = await bot.get_chat_member(chat_id, uid)
                name = member.user.full_name or member.user.username or str(uid)
            except:
                name = str(uid)
            lines.append(f"{pos}. {name} — 工作 {fmt_minutes(net_m)}，休息 {fmt_minutes(break_m)}")
            pos += 1
    await message.reply("\n".join(lines), reply_markup=get_menu())

# ---------------------------
# 今日统计工具（跨天兼容）
# ---------------------------
async def get_day_intervals_for_user_in_chat(user_id: int, chat_id: int, target_date: date):
    local_start = datetime.combine(target_date, time(0, 0, 0))
    local_end = datetime.combine(target_date, time(23, 59, 59))
    utc_start = local_start - LOCAL_OFFSET
    utc_end = local_end - LOCAL_OFFSET
    async with aiosqlite.connect(DB_PATH) as db:
        cur = await db.execute(
            "SELECT start_time, end_time FROM work_sessions "
            "WHERE user_id=? AND chat_id=? AND start_time <= ? AND (end_time IS NULL OR end_time >= ?)",
            (user_id, chat_id, to_str(utc_end), to_str(utc_start))
        )
        work_rows = await cur.fetchall()
        cur = await db.execute(
            "SELECT type, start_time, end_time FROM break_sessions "
            "WHERE user_id=? AND chat_id=? AND start_time <= ? AND (end_time IS NULL OR end_time >= ?)",
            (user_id, chat_id, to_str(utc_end), to_str(utc_start))
        )
        break_rows = await cur.fetchall()
    works = [(parse_str(s), parse_str(e) if e else None) for s, e in work_rows]
    breaks = [(t, parse_str(s), parse_str(e) if e else None) for t, s, e in break_rows]
    return works, breaks

async def compute_daily_summary(user_id: int, chat_id: int, target_date: date):
    works, breaks = await get_day_intervals_for_user_in_chat(user_id, chat_id, target_date)
    total_work = sum(minutes_between(s, e or now_utc()) for s, e in works if s)
    total_break = sum(minutes_between(s, e or now_utc()) for _, s, e in breaks if s)
    counts = {"meal": 0, "toilet_small": 0, "toilet_big": 0, "smoke": 0}
    durations = {"meal": 0, "toilet_small": 0, "toilet_big": 0, "smoke": 0}
    for btype, s, e in breaks:
        end_t = e or now_utc()
        if btype in counts:
            counts[btype] += 1
            durations[btype] += minutes_between(s, end_t)
    total_leave_times = sum(counts.values())
    total_leave_minutes = sum(durations.values())
    return {
        "total_work": total_work,
        "total_break": total_break,
        "counts": counts,
        "durations": durations,
        "total_leave_times": total_leave_times,
        "total_leave_minutes": total_leave_minutes
    }

@dp.message(F.text.contains("今日统计"))
async def handler_today_summary(message: types.Message):
    user_id = message.from_user.id
    chat_id = message.chat.id
    today = today_local_date()
    try:
        summary = await compute_daily_summary(user_id, chat_id, today)
    except Exception as e:
        logger.exception("compute_daily_summary 出错")
        await message.reply(f"❌ 统计出错：{e}")
        return

    try:
        member = await bot.get_chat_member(chat_id, user_id)
        username = member.user.full_name or member.user.username or str(user_id)
    except:
        username = str(user_id)

    text = (
        f"📋 <b>当日工作总结</b>（{today.isoformat()}）\n"
        f"👤 用户：{username}\n\n"
        f"• 工作总计：{fmt_minutes(summary['total_work'])}\n"
        f"• 休息时间：{fmt_minutes(summary['total_break'])}\n"
        f"• 离开次数：{summary['total_leave_times']}\n\n"
        f"🍱 吃饭：{summary['counts']['meal']} 次（{fmt_minutes(summary['durations']['meal'])}）\n"
        f"🚻 厕所：{summary['counts']['toilet_small'] + summary['counts']['toilet_big']} 次（{fmt_minutes(summary['durations']['toilet_small'] + summary['durations']['toilet_big'])}）\n"
        f"🚬 抽烟：{summary['counts']['smoke']} 次\n"
    )

    await message.reply(text, parse_mode="HTML", reply_markup=get_menu())

# ---------------------------
# 管理面板（多管理员） + 管理日志写入
# ---------------------------
def is_admin(user_id: int) -> bool:
    return user_id in ADMIN_IDS

def get_admin_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📝 设置提醒文字", callback_data="admin:set_text")],
        [InlineKeyboardButton(text="🖼️ 上传提醒图片", callback_data="admin:set_media")],
        [InlineKeyboardButton(text="📅 切换周报", callback_data="admin:toggle_weekly")],
        [InlineKeyboardButton(text="🗓️ 切换月报", callback_data="admin:toggle_monthly")],
        [InlineKeyboardButton(text="🔄 重置排行榜", callback_data="admin:reset_leaderboard")],
        [InlineKeyboardButton(text="📤 手动发送日报", callback_data="admin:send_daily_report")]
    ])

@dp.message(F.text == "⚙️ 设置")
async def handler_settings(message: types.Message):
    if not is_admin(message.from_user.id):
        await message.reply("🚫 你不是管理员，无法访问设置菜单。")
        return
    await message.reply("⚙️ 管理员设置菜单：", reply_markup=get_admin_menu())

@dp.callback_query(F.data == "admin:set_text")
async def admin_set_text(call: types.CallbackQuery):
    if not is_admin(call.from_user.id):
        return await call.answer("没有权限", show_alert=True)
    await call.message.answer("请输入新的提醒文字（发送一条消息即可）:")
    pending_media_for_chat[call.message.chat.id] = "awaiting_text"

@dp.message(F.text & (F.chat.id.in_(pending_media_for_chat.keys())))
async def handle_admin_input(message: types.Message):
    chat_id = message.chat.id
    state = pending_media_for_chat.get(chat_id)
    if state == "awaiting_text":
        await set_chat_setting(chat_id, "reminder_text", message.text)
        await log_admin_action(chat_id, message.from_user.id, "set_reminder_text", message.text[:400])
        del pending_media_for_chat[chat_id]
        await message.reply("✅ 提醒文字已更新。")
    elif state == "awaiting_media":
        await message.reply("请上传一张图片而不是文字。")

@dp.callback_query(F.data == "admin:set_media")
async def admin_set_media(call: types.CallbackQuery):
    if not is_admin(call.from_user.id):
        return await call.answer("没有权限", show_alert=True)
    await call.message.answer("请发送一张图片作为提醒媒体：")
    pending_media_for_chat[call.message.chat.id] = "awaiting_media"

@dp.message(F.photo)
async def handle_admin_photo(message: types.Message):
    chat_id = message.chat.id
    if pending_media_for_chat.get(chat_id) == "awaiting_media":
        file_id = message.photo[-1].file_id
        await set_chat_setting(chat_id, "reminder_media_file_id", file_id)
        await log_admin_action(chat_id, message.from_user.id, "set_reminder_media", f"file_id:{file_id}")
        del pending_media_for_chat[chat_id]
        await message.reply("✅ 提醒图片已更新。")

@dp.callback_query(F.data == "admin:toggle_weekly")
async def admin_toggle_weekly(call: types.CallbackQuery):
    if not is_admin(call.from_user.id):
        return await call.answer("没有权限", show_alert=True)
    settings = await get_chat_settings(call.message.chat.id)
    new_value = 0 if settings["weekly_report_enabled"] else 1
    await set_chat_setting(call.message.chat.id, "weekly_report_enabled", new_value)
    await log_admin_action(call.message.chat.id, call.from_user.id, "toggle_weekly", f"set_to:{new_value}")
    status = "✅ 已开启" if new_value else "❌ 已关闭"
    await call.message.edit_text(f"📅 周报功能 {status}", reply_markup=get_admin_menu())

@dp.callback_query(F.data == "admin:toggle_monthly")
async def admin_toggle_monthly(call: types.CallbackQuery):
    if not is_admin(call.from_user.id):
        return await call.answer("没有权限", show_alert=True)
    settings = await get_chat_settings(call.message.chat.id)
    new_value = 0 if settings["monthly_report_enabled"] else 1
    await set_chat_setting(call.message.chat.id, "monthly_report_enabled", new_value)
    await log_admin_action(call.message.chat.id, call.from_user.id, "toggle_monthly", f"set_to:{new_value}")
    status = "✅ 已开启" if new_value else "❌ 已关闭"
    await call.message.edit_text(f"🗓️ 月报功能 {status}", reply_markup=get_admin_menu())

@dp.callback_query(F.data == "admin:reset_leaderboard")
async def admin_reset_leaderboard(call: types.CallbackQuery):
    if not is_admin(call.from_user.id):
        return await call.answer("没有权限", show_alert=True)
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute("DELETE FROM work_sessions WHERE chat_id = ?", (call.message.chat.id,))
        await db.execute("DELETE FROM break_sessions WHERE chat_id = ?", (call.message.chat.id,))
        await db.commit()
    await log_admin_action(call.message.chat.id, call.from_user.id, "reset_leaderboard", "cleared work_sessions and break_sessions")
    await call.message.answer("🔄 排行榜已重置！")
    await call.message.edit_text("操作完成 ✅", reply_markup=get_admin_menu())

@dp.callback_query(F.data == "admin:send_daily_report")
async def admin_send_daily_report(call: types.CallbackQuery):
    if not is_admin(call.from_user.id):
        return await call.answer("没有权限", show_alert=True)
    today = today_local_date()
    chat_id = call.message.chat.id
    await send_report_for_chat(chat_id, "daily", today)
    await log_admin_action(chat_id, call.from_user.id, "manual_send_daily", f"sent daily for {today.isoformat()}")
    await call.message.answer("📊 日报已发送给管理员。")

# ---------------------------
# 超时提醒 watcher
# ---------------------------
async def break_overtime_watcher(user_id: int, chat_id: int, btype: str, start_dt_utc: datetime):
    limit_minutes = BREAK_LIMITS.get(btype, 5)
    limit_dt = start_dt_utc + timedelta(minutes=limit_minutes)
    while True:
        await asyncio.sleep(OVERTIME_REMINDER_INTERVAL * 60)
        now = now_utc()
        async with aiosqlite.connect(DB_PATH) as db:
            cur = await db.execute("SELECT id FROM break_sessions WHERE user_id=? AND chat_id=? AND end_time IS NULL", (user_id, chat_id))
            row = await cur.fetchone()
        if not row:
            break
        if now >= limit_dt:
            settings = await get_chat_settings(chat_id)
            rtext = settings.get("reminder_text") or f"⚠️ <a href='tg://user?id={user_id}'>你</a> 已超时，请尽快回座。"
            try:
                media_file = settings.get("reminder_media_file_id")
                if media_file:
                    try:
                        await bot.send_photo(chat_id, media_file, caption=rtext, parse_mode="HTML")
                    except:
                        await bot.send_message(chat_id, rtext, parse_mode="HTML")
                else:
                    await bot.send_message(chat_id, rtext, parse_mode="HTML")
            except Exception as e:
                logger.exception(f"发送超时提醒失败: {e}")
            break  # 一次提醒后停止本次 watcher

# ---------------------------
# 报表：收集 / 生成 / 发送（Excel）
# ---------------------------
from aiogram.types import *
import BufferedInputFile

def safe_filename(s: str) -> str:
    # 移除文件名非法字符
    return re.sub(r'[\\/:"*?<>|]+', "_", s)

async def gather_users_in_chat(chat_id: int):
    async with aiosqlite.connect(DB_PATH) as db:
        cur = await db.execute("SELECT DISTINCT user_id FROM work_sessions WHERE chat_id = ?", (chat_id,))
        rows = await cur.fetchall()
    return [r[0] for r in rows]

async def get_work_range_for_user(user_id: int, chat_id: int, start_utc: datetime, end_utc: datetime):
    async with aiosqlite.connect(DB_PATH) as db:
        cur = await db.execute(
            "SELECT start_time, end_time FROM work_sessions WHERE user_id=? AND chat_id=? AND start_time <= ? AND (end_time IS NULL OR end_time >= ?)",
            (user_id, chat_id, to_str(end_utc), to_str(start_utc))
        )
        rows = await cur.fetchall()
    starts = []
    ends = []
    total_work = 0
    for s, e in rows:
        ps = parse_str(s)
        pe = parse_str(e) if e else None
        if ps:
            starts.append(ps)
        if pe:
            ends.append(pe)
        total_work += minutes_between(ps, pe or end_utc)
    first_start = min(starts) if starts else None
    last_end = max(ends) if ends else None
    return first_start, last_end, total_work

async def get_break_summary_for_user(user_id: int, chat_id: int, start_utc: datetime, end_utc: datetime):
    async with aiosqlite.connect(DB_PATH) as db:
        cur = await db.execute(
            "SELECT type, start_time, end_time FROM break_sessions WHERE user_id=? AND chat_id=? AND start_time <= ? AND (end_time IS NULL OR end_time >= ?)",
            (user_id, chat_id, to_str(end_utc), to_str(start_utc))
        )
        rows = await cur.fetchall()
    total_break = 0
    leave_count = 0
    for t, s, e in rows:
        ps = parse_str(s)
        pe = parse_str(e) if e else end_utc
        if ps:
            total_break += minutes_between(ps, pe)
            leave_count += 1
    return total_break, leave_count


async def send_report_for_chat(chat_id: int, period: str, base_date: date):
    # 计算 local_start / local_end
    if period == "daily":
        local_start = datetime.combine(base_date, time.min)
        local_end = datetime.combine(base_date, time.max)
        prefix = "日报"
    elif period == "weekly":
        start_local = base_date - timedelta(days=base_date.weekday())
        local_start = datetime.combine(start_local, time.min)
        local_end = local_start + timedelta(days=6, hours=23, minutes=59, seconds=59)
        prefix = "周报"
    elif period == "monthly":
        start_local = base_date.replace(day=1)
        if start_local.month == 12:
            next_month = start_local.replace(year=start_local.year + 1, month=1, day=1)
        else:
            next_month = start_local.replace(month=start_local.month + 1, day=1)
        local_start = datetime.combine(start_local, time.min)
        local_end = datetime.combine(next_month - timedelta(seconds=1), time.max)
        prefix = "月报"
    else:
        return

    utc_start = local_start - LOCAL_OFFSET
    utc_end = local_end - LOCAL_OFFSET

    users = await gather_users_in_chat(chat_id)
    if not users:
        logger.info(f"chat {chat_id} 没有用户数据，跳过 {period} 报表。")
        return

    rows = []
    for uid in users:
        try:
            member = await bot.get_chat_member(chat_id, uid)
            name = member.user.full_name or member.user.username or str(uid)
        except:
            name = str(uid)
        first_start, last_end, total_work = await get_work_range_for_user(uid, chat_id, utc_start, utc_end)
        total_break, leave_count = await get_break_summary_for_user(uid, chat_id, utc_start, utc_end)
        first_start_s = fmt_hm_local(first_start) if first_start else "-"
        last_end_s = fmt_hm_local(last_end) if last_end else "-"
        rows.append((name, first_start_s, last_end_s, total_work, total_break, leave_count))

    rows.sort(key=lambda x: x[3], reverse=True)

    # 生成 Excel 报表
    wb = Workbook()
    ws = wb.active
    ws.title = f"{prefix}"

    headers = ["姓名", "上班时间", "下班时间", "工作时间(文本)", "休息时间(文本)", "离开次数", "工作时间(分钟)", "休息时间(分钟)"]
    ws.append(headers)
    header_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
    header_font = Font(bold=True)
    align_center = Alignment(horizontal="center", vertical="center")
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center

    for name, start_s, end_s, work_m, break_m, leave_cnt in rows:
        ws.append([
            name,
            start_s,
            end_s,
            fmt_minutes(work_m),
            fmt_minutes(break_m),
            leave_cnt,
            work_m,
            break_m
        ])

    # 自动列宽
    for col in ws.columns:
        max_length = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_length + 3

    # 保存到内存
    file_bytes = io.BytesIO()
    wb.save(file_bytes)
    file_bytes.seek(0)
    bytes_data = file_bytes.getvalue()

    # 获取群名
    try:
        chat = await bot.get_chat(chat_id)
        chat_title = chat.title or "群名未知"
    except Exception:
        chat_title = "群名未知"

    fname_safe = safe_filename(f"{prefix}_{chat_title}_{base_date.isoformat()}.xlsx")
    caption = f"📤 [{chat_title}] (ID: {chat_id}) 的 {prefix}\n时区：UTC{int(LOCAL_OFFSET.total_seconds()/3600):+d}"

    # 发送给所有管理员
    for admin in ADMIN_IDS:
        try:
            buffered = BufferedInputFile(bytes_data, filename=fname_safe)
            await bot.send_document(admin, document=buffered, caption=caption)
            logger.info(f"✅ 已发送 {prefix} 给管理员 {admin}")
        except Exception as e:
            logger.warning(f"发送报表给管理员 {admin} 失败: {e}")


# ---------------------------
# 定时任务（apscheduler）
# ---------------------------
scheduler = AsyncIOScheduler(timezone="Asia/Jakarta")

@scheduler.scheduled_job(CronTrigger(hour=DAILY_REPORT_HOUR, minute=0))
async def scheduled_daily_report():
    today = today_local_date()
    async with aiosqlite.connect(DB_PATH) as db:
        cur = await db.execute("SELECT DISTINCT chat_id FROM work_sessions")
        rows = await cur.fetchall()
    for (chat_id,) in rows:
        await send_report_for_chat(chat_id, "daily", today)

@scheduler.scheduled_job(CronTrigger(day_of_week="mon", hour=WEEKLY_REPORT_HOUR, minute=0))
async def scheduled_weekly_report():
    today = today_local_date()
    chats = await get_chats_with_setting_enabled("weekly_report_enabled")
    for cid in chats:
        await send_report_for_chat(cid, "weekly", today)

@scheduler.scheduled_job(CronTrigger(day=MONTHLY_REPORT_DAY, hour=MONTHLY_REPORT_HOUR, minute=0))
async def scheduled_monthly_report():
    today = today_local_date()
    chats = await get_chats_with_setting_enabled("monthly_report_enabled")
    for cid in chats:
        await send_report_for_chat(cid, "monthly", today)

# 手动触发日报命令（管理员）
@dp.message(F.text.contains("手动发送日报"))
async def manual_daily_report(message: types.Message):
    if message.from_user.id not in ADMIN_IDS:
        await message.reply("🚫 你不是管理员，无权执行此操作。")
        return
    today = today_local_date()
    async with aiosqlite.connect(DB_PATH) as db:
        cur = await db.execute("SELECT DISTINCT chat_id FROM work_sessions")
        rows = await cur.fetchall()
    for (chat_id,) in rows:
        await send_report_for_chat(chat_id, "daily", today)
    await message.reply("✅ 日报已生成并发送给管理员。")


# ---------------------------
# 启动
# ---------------------------
async def main():
    await init_db()
    scheduler.start()
    logger.info("调度器已启动（日报/周报/月报）。")
    await dp.start_polling(bot)

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logger.info("已停止。")

