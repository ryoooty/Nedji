import asyncio
import json
import logging
import os
import subprocess
import sys
from math import ceil

import aiogram.exceptions
import psutil
import pyautogui
from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import Command, CommandStart
from aiogram.types import BotCommand, CallbackQuery, InlineKeyboardButton, Message, FSInputFile
from aiogram.utils.keyboard import InlineKeyboardBuilder
import keyboard
import ctypes
import time
import win32gui
import win32con
import win32process
import re
from pathlib import Path
import configparser
import urllib.parse
from typing import Optional, List

# ==========================
# БАЗОВАЯ НАСТРОЙКА
# ==========================

def get_app_dir() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

APP_DIR = get_app_dir()

CONFIG_PATH = APP_DIR / 'config.ini'
config = configparser.ConfigParser()

if not CONFIG_PATH.exists():
    logging.info("Файл config.ini не найден. Создаю шаблон...")
    config['Settings'] = {
        'TELEGRAM_BOT_TOKEN': 'YOUR_TOKEN_HERE',
        'USER_ID': '0',
        'DEFAULT_SEARCH_ENGINE': 'google',
        'PREFERRED_SEARCH_BROWSER_KEY': ''
    }
    with open(CONFIG_PATH, 'w', encoding='utf-8') as configfile:
        config.write(configfile)
    logging.info(f"Шаблон конфигурации создан в {CONFIG_PATH}. Пожалуйста, заполните его.")

try:
    config.read(CONFIG_PATH, encoding='utf-8')
    BOT_TOKEN = config.get('Settings', 'TELEGRAM_BOT_TOKEN')
    USER_ID = config.getint('Settings', 'USER_ID')
    DEFAULT_SEARCH_ENGINE = config.get('Settings', 'DEFAULT_SEARCH_ENGINE', fallback='google')
    PREFERRED_SEARCH_BROWSER_KEY = config.get('Settings', 'PREFERRED_SEARCH_BROWSER_KEY', fallback='').strip()
except (configparser.Error, ValueError) as e:
    logging.error(f"Ошибка чтения config.ini: {e}")
    sys.exit(1)

if BOT_TOKEN == "YOUR_TOKEN_HERE" or USER_ID == 0:
    logging.error("Пожалуйста, укажите ваш TELEGRAM_BOT_TOKEN и USER_ID в файле config.ini")
    sys.exit(1)

SEARCH_ENGINES = {
    'yandex': 'https://yandex.ru/search/?text=',
    'google': 'https://www.google.com/search?q=',
    'bing': 'https://www.bing.com/search?q='
}
if DEFAULT_SEARCH_ENGINE not in SEARCH_ENGINES:
    logging.warning(f"Недопустимое значение DEFAULT_SEARCH_ENGINE: {DEFAULT_SEARCH_ENGINE}. Используется 'google'.")
    DEFAULT_SEARCH_ENGINE = 'google'
    config.set('Settings', 'DEFAULT_SEARCH_ENGINE', 'google')
    with open(CONFIG_PATH, 'w', encoding='utf-8') as configfile:
        config.write(configfile)

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

user_data = {
    'mode': 'inline',
    'apps_page': 0,
    'combos_page': 0,
    'search_engine': DEFAULT_SEARCH_ENGINE,
    'preferred_search_browser_key': PREFERRED_SEARCH_BROWSER_KEY if PREFERRED_SEARCH_BROWSER_KEY else None
}
toggle_state = {}
active_inline_menus = {}

# Состояние записи экрана (через Xbox Game Bar)
record_state = {
    'active': False,
    'started_at': 0.0
}
last_clip_by_user: dict[int, Optional[Path]] = {}

# ==========================
# ДАННЫЕ ПРИЛОЖЕНИЙ / КОМБО
# ==========================

def load_data(filename: str):
    path = APP_DIR / filename
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        logging.error(f"Ошибка загрузки {filename}: {e}")
        return []

APPS_JSON_PATH = APP_DIR / 'apps.json'
COMBOS_JSON_PATH = APP_DIR / 'combos.json'
apps_data = load_data('apps.json')
combos_data = load_data('combos.json')

def has_access(message: types.Message):
    return message.from_user.id == USER_ID

# ==========================
# КЛАВИАТУРЫ
# ==========================

def get_main_keyboard():
    buttons = [
        [types.KeyboardButton(text="📱 Приложения")],
        [types.KeyboardButton(text="⌨️ Комбинации")],
        [types.KeyboardButton(text="🖥 Управление")]
    ]
    return types.ReplyKeyboardMarkup(keyboard=buttons, resize_keyboard=True)

def get_controls_keyboard():
    builder = InlineKeyboardBuilder()
    builder.button(text="Up", callback_data="media_page_up")
    builder.button(text="⬆️", callback_data="media_arrow_up")
    builder.button(text="Dn", callback_data="media_page_down")
    builder.button(text="⬅️", callback_data="media_arrow_left")
    builder.button(text="⬇️", callback_data="media_arrow_down")
    builder.button(text="➡️", callback_data="media_arrow_right")
    builder.button(text="⌨️", callback_data="media_switch_reply")
    builder.button(text="⎵", callback_data="media_space")
    builder.button(text="🔉", callback_data="media_volume_down")
    builder.button(text="🔇", callback_data="media_volume_mute")
    builder.button(text="🔊", callback_data="media_volume_up")
    # Управление треками (системные медиа-клавиши)
    builder.button(text="⏮", callback_data="media_prev")
    builder.button(text="⏯", callback_data="media_play_pause")
    builder.button(text="⏭", callback_data="media_next")
    builder.adjust(3, 3, 2, 3, 3)
    return builder.as_markup()

def get_controls_reply_keyboard():
    buttons = [
        ["Up", "⬆️", "Dn"],
        ["⬅️", "⬇️", "➡️"],
        ["⌨️", "⎵"],
        ["🔉", "🔇", "🔊"],
        ["⏮", "⏯", "⏭"]
    ]
    return types.ReplyKeyboardMarkup(
        keyboard=[[types.KeyboardButton(text=b) for b in row] for row in buttons],
        resize_keyboard=True
    )

# ==========================
# ХЕНДЛЕРЫ «УПРАВЛЕНИЕ»
# ==========================

@dp.message(F.text == "🖥 Управление")
async def show_controls(message: Message):
    if not has_access(message):
        return
    sent_message = await message.answer("Управление клавишами:", reply_markup=get_controls_keyboard())
    user_id = message.from_user.id
    category = 'media'
    if user_id not in active_inline_menus:
        active_inline_menus[user_id] = {'apps': [], 'combos': [], 'media': []}
    active_inline_menus[user_id][category].append(sent_message.message_id)
    if len(active_inline_menus[user_id][category]) > 2:
        oldest_msg_id = active_inline_menus[user_id][category].pop(0)
        try:
            await bot.edit_message_reply_markup(chat_id=message.chat.id, message_id=oldest_msg_id, reply_markup=None)
        except aiogram.exceptions.TelegramBadRequest:
            pass

@dp.callback_query(F.data == "media_switch_reply")
async def switch_to_reply_controls(callback: CallbackQuery):
    user_data['mode'] = 'media_reply'
    await callback.message.answer("Управление клавишами (reply):", reply_markup=get_controls_reply_keyboard())
    await callback.answer()

@dp.message(F.text.in_(["Up", "⬆️", "Dn", "⬅️", "⬇️", "➡️", "⎵", "🔉", "🔇", "🔊", "⏮", "⏯", "⏭"]))
async def handle_controls_reply(message: Message):
    if user_data.get('mode') != 'media_reply':
        return
    t = message.text
    try:
        if t == "Up":
            keyboard.send('page up')
        elif t == "Dn":
            keyboard.send('page down')
        elif t == "⬆️":
            keyboard.send('up')
        elif t == "⬇️":
            keyboard.send('down')
        elif t == "⬅️":
            keyboard.send('left')
        elif t == "➡️":
            keyboard.send('right')
        elif t == "⎵":
            keyboard.send('space')
        elif t == "🔉":
            volume_down()
        elif t == "🔊":
            volume_up()
        elif t == "🔇":
            volume_mute()
        elif t == "⏮":
            media_prev()
        elif t == "⏯":
            media_play_pause()
        elif t == "⏭":
            media_next()
    except Exception as e:
        logging.error(f"Ошибка обработки кнопки '{t}': {e}")

@dp.callback_query(F.data.startswith("media_"))
async def process_controls(callback: CallbackQuery):
    action = callback.data[len("media_"):]
    if action == "switch_reply":
        return
    try:
        if action == "arrow_left":
            keyboard.send('left')
        elif action == "arrow_right":
            keyboard.send('right')
        elif action == "arrow_up":
            keyboard.send('up')
        elif action == "arrow_down":
            keyboard.send('down')
        elif action == "page_up":
            keyboard.send('page up')
        elif action == "page_down":
            keyboard.send('page down')
        elif action == "space":
            keyboard.send('space')
        elif action == "volume_up":
            volume_up()
        elif action == "volume_down":
            volume_down()
        elif action == "volume_mute":
            volume_mute()
        elif action == "prev":
            media_prev()
        elif action == "play_pause":
            media_play_pause()
        elif action == "next":
            media_next()
        else:
            await callback.answer("Неизвестная команда.", show_alert=True)
            return
        await callback.answer()
    except Exception as e:
        await callback.answer(f"Ошибка: {e}", show_alert=True)

@dp.message(F.text == "⌨️")
async def switch_to_inline_controls(message: Message):
    if user_data.get('mode') == 'media_reply':
        user_data['mode'] = 'inline'
        await message.answer("Главное меню:", reply_markup=get_main_keyboard())
    else:
        await message.answer("Главное меню:", reply_markup=get_main_keyboard())

# ==========================
# БАЗОВЫЕ ХЕНдЛЕРЫ
# ==========================

@dp.message(CommandStart())
async def send_welcome(message: Message):
    if not has_access(message):
        await message.answer("У вас нет доступа.")
        return
    await message.answer("Бот для управления ПК запущен!", reply_markup=get_main_keyboard())
    await set_commands()

@dp.message(Command("end"))
async def end_bot(message: Message):
    if not has_access(message):
        return
    await message.answer("Бот остановлен.", reply_markup=types.ReplyKeyboardRemove())
    await dp.storage.close()
    await bot.session.close()
    await dp.stop_polling()

@dp.message(Command("reload"))
async def reload_data(message: Message):
    if not has_access(message):
        return
    global apps_data, combos_data
    apps_data = load_data('apps.json')
    combos_data = load_data('combos.json')
    toggle_state.clear()
    await message.answer("Данные из JSON файлов успешно обновлены.")

file_wait = {'type': None}

@dp.message(Command("editapps"))
async def edit_apps(message: Message):
    if not has_access(message):
        await message.answer("У вас нет доступа к редактированию файлов.")
        return
    try:
        await message.answer_document(types.FSInputFile(APPS_JSON_PATH), caption="Файл apps.json для редактирования")
    except Exception as e:
        await message.answer(f"Ошибка отправки файла: {e}")

@dp.message(Command("editcombos"))
async def edit_combos(message: Message):
    if not has_access(message):
        await message.answer("У вас нет доступа к редактированию файлов.")
        return
    try:
        await message.answer_document(types.FSInputFile(COMBOS_JSON_PATH), caption="Файл combos.json для редактирования")
    except Exception as e:
        await message.answer(f"Ошибка отправки файла: {e}")

@dp.message(Command("saveapps"))
async def wait_for_apps_file(message: Message):
    if not has_access(message):
        await message.answer("У вас нет доступа к редактированию файлов.")
        return
    file_wait['type'] = 'apps'
    await message.answer("Отправьте файл apps.json для замены.")

@dp.message(Command("savecombos"))
async def wait_for_combos_file(message: Message):
    if not has_access(message):
        await message.answer("У вас нет доступа к редактированию файлов.")
        return
    file_wait['type'] = 'combos'
    await message.answer("Отправьте файл combos.json для замены.")

@dp.message(F.document)
async def handle_json_upload(message: Message):
    """Принимаем только JSON (музыкальные загрузки удалены)."""
    if not has_access(message):
        return
    file = message.document
    if file_wait['type'] in ('apps', 'combos') and file.file_name.endswith('.json'):
        file_type = file_wait['type']
        file_wait['type'] = None
        filename = APPS_JSON_PATH if file_type == 'apps' else COMBOS_JSON_PATH
        try:
            await bot.download(file, destination=filename)
            with open(filename, 'r', encoding='utf-8') as f:
                new_data = json.load(f)
            global apps_data, combos_data
            if file_type == 'apps':
                apps_data = new_data
            else:
                combos_data = new_data
            await message.answer(f"Файл {filename.name} успешно обновлён!")
        except json.JSONDecodeError as e:
            await message.answer(f"Ошибка в JSON файле: {e}")
        except Exception as e:
            await message.answer(f"Ошибка сохранения файла: {e}")

# ==========================
# МЕНЮ «ПРИЛОЖЕНИЯ»
# ==========================

def get_apps_keyboard(page=0):
    builder = InlineKeyboardBuilder()
    items_per_page = 20
    visible_apps = [a for a in apps_data if a.get('show_in_menu', True)]
    start_index = page * items_per_page
    end_index = start_index + items_per_page
    current_apps = visible_apps[start_index:end_index]
    for app in current_apps:
        builder.button(text=app['name'], callback_data=f"app_toggle_{app['key']}")
    builder.adjust(2)
    total_pages = ceil(len(visible_apps) / items_per_page) if visible_apps else 1
    if total_pages > 1:
        prev_page = (page - 1 + total_pages) % total_pages
        next_page = (page + 1) % total_pages
        builder.row(
            InlineKeyboardButton(text="⬅️", callback_data=f"app_page_{prev_page}"),
            InlineKeyboardButton(text=f"{page + 1}/{total_pages}", callback_data="noop"),
            InlineKeyboardButton(text="➡️", callback_data=f"app_page_{next_page}")
        )
    return builder.as_markup()

@dp.message(F.text == "📱 Приложения")
async def show_apps(message: Message):
    if not has_access(message):
        return
    if not apps_data:
        await message.answer("Список приложений пуст. Добавьте их в `apps.json`.")
        return
    user_data['apps_page'] = 0
    sent_message = await message.answer("Выберите приложение:", reply_markup=get_apps_keyboard(user_data['apps_page']))
    user_id = message.from_user.id
    category = 'apps'
    if user_id not in active_inline_menus:
        active_inline_menus[user_id] = {'apps': [], 'combos': [], 'media': []}
    active_inline_menus[user_id][category].append(sent_message.message_id)
    if len(active_inline_menus[user_id][category]) > 2:
        oldest_msg_id = active_inline_menus[user_id][category].pop(0)
        try:
            await bot.edit_message_reply_markup(chat_id=message.chat.id, message_id=oldest_msg_id, reply_markup=None)
        except aiogram.exceptions.TelegramBadRequest:
            pass

@dp.callback_query(F.data.startswith("app_page_"))
async def process_app_page(callback: CallbackQuery):
    page = int(callback.data[len("app_page_" ):])
    user_data['apps_page'] = page
    try:
        await callback.message.edit_text("Выберите приложение:", reply_markup=get_apps_keyboard(page))
        user_id = callback.from_user.id
        category = 'apps'
        current_msg_id = callback.message.message_id
        if user_id in active_inline_menus and current_msg_id not in active_inline_menus[user_id][category]:
            active_inline_menus[user_id][category].append(current_msg_id)
            if len(active_inline_menus[user_id][category]) > 2:
                oldest_msg_id = active_inline_menus[user_id][category].pop(0)
                try:
                    await bot.edit_message_reply_markup(chat_id=callback.message.chat.id, message_id=oldest_msg_id, reply_markup=None)
                except aiogram.exceptions.TelegramBadRequest:
                    pass
    except aiogram.exceptions.TelegramBadRequest as e:
        if "message is not modified" not in str(e).lower():
            logging.warning(f"TelegramBadRequest при редактировании клавиатуры приложений: {e}")
    await callback.answer()

def _as_list(val):
    if val is None:
        return []
    if isinstance(val, list):
        return val
    if isinstance(val, str) and val.strip():
        return [val]
    return []

def _is_url(s: str) -> bool:
    return bool(re.match(r'^(https?|steam|tg)://', s, re.IGNORECASE))

def _resolve_path(p: str) -> str:
    pp = Path(p)
    return str(pp if pp.is_absolute() else (APP_DIR / pp))

def _open_with_shell(path: str):
    os.startfile(path)

def _run_exe(path: str, args: List[str]):
    subprocess.Popen([path] + args, shell=False)

@dp.callback_query(F.data.startswith("app_toggle_"))
async def toggle_app(callback: CallbackQuery):
    key = callback.data[len("app_toggle_"):]
    app_info = next((app for app in apps_data if app['key'] == key and app.get('show_in_menu', True)), None)
    if not app_info:
        await callback.answer("Приложение не найдено!", show_alert=True)
        return

    # 0) Steam по appid
    steam_appid = app_info.get('steam_appid')
    if steam_appid:
        try:
            _open_with_shell(f"steam://rungameid/{steam_appid}")
            await callback.answer()
            return
        except Exception as e:
            await callback.answer(f"Steam ошибка: {e}", show_alert=True)
            return

    path = str(app_info.get('path', '')).strip()
    args = _as_list(app_info.get('args')) or _as_list(app_info.get('arg'))
    is_app = str(app_info.get('is_app', 'y')).lower()

    # 1) Явный steam:// / tg:// / http(s)://
    if _is_url(path):
        try:
            _open_with_shell(path)
            await callback.answer()
            return
        except Exception as e:
            # Спец-fallback: Telegram не установлен → открыть сайт установки
            if path.lower().startswith("tg://"):
                try:
                    _open_with_shell("https://desktop.telegram.org")
                    await callback.answer("Telegram не найден — открыл страницу установки.")
                    return
                except Exception as e2:
                    await callback.answer(f"Ошибка открытия Telegram: {e2}", show_alert=True)
                    return
            await callback.answer(f"Ошибка URL: {e}", show_alert=True)
            return

    # 2) .url ярлык?
    if path.lower().endswith(".url"):
        try:
            _open_with_shell(_resolve_path(path))
            await callback.answer()
            return
        except Exception as e:
            await callback.answer(f"Ошибка ярлыка: {e}", show_alert=True)
            return

    # 3) Обычный EXE/файл
    resolved = _resolve_path(path)

    if is_app == 'n':
        try:
            if resolved.lower().endswith(".exe") and args:
                _run_exe(resolved, args)
            else:
                _open_with_shell(resolved)
            await callback.answer()
            return
        except Exception as e:
            await callback.answer(f"Ошибка запуска: {e}", show_alert=True)
            return

    # is_app == 'y' → показать/свернуть (если запущено), иначе — запустить
    state = toggle_state.get(key, 'minimized')
    try:
        if state == 'minimized':
            try:
                if not activate_app_window(app_info):
                    if resolved.lower().endswith(".exe"):
                        _run_exe(resolved, args)
                    else:
                        _open_with_shell(resolved)
            except Exception:
                _open_with_shell(resolved)
            toggle_state[key] = 'shown'
            await callback.answer()
        else:
            try:
                minimize_app_window(app_info)
            except Exception:
                pass
            toggle_state[key] = 'minimized'
            await callback.answer()
    except Exception as e:
        await callback.answer(f"Ошибка: {e}", show_alert=True)

    try:
        await callback.message.edit_reply_markup(reply_markup=get_apps_keyboard(user_data['apps_page']))
    except aiogram.exceptions.TelegramBadRequest as e:
        if "message is not modified" not in str(e).lower():
            logging.warning(f"TelegramBadRequest при редактировании клавиатуры приложений: {e}")

def activate_app_window(app_info):
    exe_name = app_info.get('exe') or app_info.get('path', '').split('\\')[-1]
    pids = [p.info['pid'] for p in psutil.process_iter(['pid', 'name'])
            if p.info['name'] and exe_name.lower() in p.info['name'].lower()]
    if not pids:
        return False
    for pid in pids:
        def callback(hwnd, _):
            _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
            if found_pid == pid and win32gui.IsWindowVisible(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                try:
                    win32gui.SetForegroundWindow(hwnd)
                except Exception:
                    pass
                return False
            return True
        win32gui.EnumWindows(callback, None)
    return True

def minimize_app_window(app_info):
    exe_name = app_info.get('exe') or app_info.get('path', '').split('\\')[-1]
    pids = [p.info['pid'] for p in psutil.process_iter(['pid', 'name'])
            if p.info['name'] and exe_name.lower() in p.info['name'].lower()]
    if not pids:
        return False
    for pid in pids:
        def callback(hwnd, _):
            _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
            if found_pid == pid:
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                return False
            return True
        win32gui.EnumWindows(callback, None)
    return True

# ==========================
# МЕНЮ «КОМБИНАЦИИ»
# ==========================

def get_combos_keyboard(page=0):
    builder = InlineKeyboardBuilder()
    items_per_page = 20
    visible_combos = [c for c in combos_data if c.get('show_in_menu', True)]
    start_index = page * items_per_page
    end_index = start_index + items_per_page
    current_combos = visible_combos[start_index:end_index]
    for combo in current_combos:
        builder.button(text=combo['name'], callback_data=f"combo_run_{combo['key']}")
    builder.adjust(2)
    total_pages = ceil(len(visible_combos) / items_per_page) if visible_combos else 1
    if total_pages > 1:
        prev_page = (page - 1 + total_pages) % total_pages
        next_page = (page + 1) % total_pages
        builder.row(
            InlineKeyboardButton(text="⬅️", callback_data=f"combo_page_{prev_page}"),
            InlineKeyboardButton(text=f"{page + 1}/{total_pages}", callback_data="noop"),
            InlineKeyboardButton(text="➡️", callback_data=f"combo_page_{next_page}")
        )
    return builder.as_markup()

@dp.message(F.text == "⌨️ Комбинации")
async def show_combos(message: Message):
    if not has_access(message):
        return
    if not combos_data:
        await message.answer("Список комбинаций пуст. Добавьте их в `combos.json`.")
        return
    user_data['combos_page'] = 0
    sent_message = await message.answer("Выберите комбинацию:", reply_markup=get_combos_keyboard(user_data['combos_page']))
    user_id = message.from_user.id
    category = 'combos'
    if user_id not in active_inline_menus:
        active_inline_menus[user_id] = {'apps': [], 'combos': [], 'media': []}
    active_inline_menus[user_id][category].append(sent_message.message_id)
    if len(active_inline_menus[user_id][category]) > 2:
        oldest_msg_id = active_inline_menus[user_id][category].pop(0)
        try:
            await bot.edit_message_reply_markup(chat_id=message.chat.id, message_id=oldest_msg_id, reply_markup=None)
        except aiogram.exceptions.TelegramBadRequest:
            pass

@dp.callback_query(F.data.startswith("combo_page_"))
async def process_combo_page(callback: CallbackQuery):
    page = int(callback.data[len("combo_page_"):])
    user_data['combos_page'] = page
    try:
        await callback.message.edit_text("Выберите комбинацию:", reply_markup=get_combos_keyboard(page))
        user_id = callback.from_user.id
        category = 'combos'
        current_msg_id = callback.message.message_id
        if user_id in active_inline_menus and current_msg_id not in active_inline_menus[user_id][category]:
            active_inline_menus[user_id][category].append(current_msg_id)
            if len(active_inline_menus[user_id][category]) > 2:
                oldest_msg_id = active_inline_menus[user_id][category].pop(0)
                try:
                    await bot.edit_message_reply_markup(chat_id=callback.message.chat.id, message_id=oldest_msg_id, reply_markup=None)
                except aiogram.exceptions.TelegramBadRequest:
                    pass
    except aiogram.exceptions.TelegramBadRequest as e:
        if "message is not modified" not in str(e).lower():
            logging.warning(f"TelegramBadRequest при редактировании клавиатуры комбинаций: {e}")
    await callback.answer()

def save_config_setting(section, key, value):
    try:
        config.set(section, key, str(value))
        with open(CONFIG_PATH, 'w', encoding='utf-8') as configfile:
            config.write(configfile)
    except Exception as e:
        logging.error(f"Ошибка сохранения настройки {section}.{key} в config.ini: {e}")

# ===== Спец-утилиты для записи экрана (Xbox Game Bar) =====

def _videos_dirs() -> List[Path]:
    # На локализованных Windows путь остаётся "Videos", отображаемое имя — «Видео»
    candidates = [
        Path.home() / "Videos",
        Path.home() / "Видео"
    ]
    return [p for p in candidates if p.exists()]

def _captures_dirs() -> List[Path]:
    out = []
    subs = ["Captures", "Клипы", "Game Clips", "Игровые клипы"]
    for vd in _videos_dirs():
        for sub in subs:
            p = vd / sub
            if p.exists():
                out.append(p)
        # На некоторых системах клипы пишутся прямо в Videos
        out.append(vd)
    # Уникализировать с сохранением порядка
    unique = []
    seen = set()
    for p in out:
        if str(p) not in seen:
            seen.add(str(p)); unique.append(p)
    return unique

def _find_latest_clip(since_ts: float) -> Optional[Path]:
    exts = {'.mp4', '.mov', '.mkv', '.avi'}
    newest: tuple[float, Path] | None = None
    for d in _captures_dirs():
        try:
            for f in d.glob("*"):
                if f.is_file() and f.suffix.lower() in exts:
                    mtime = f.stat().st_mtime
                    # небольшой зазор назад — чтобы учесть задержки финализации файла
                    if mtime >= since_ts - 60:
                        if newest is None or mtime > newest[0]:
                            newest = (mtime, f)
        except Exception:
            continue
    return newest[1] if newest else None

# ===== Основная логика выполнения комбинаций =====

@dp.callback_query(F.data.startswith("combo_run_"))
async def run_combo(callback: CallbackQuery):
    key = callback.data[len("combo_run_"):]
    combo_info = next((c for c in combos_data if c['key'] == key), None)
    if not combo_info:
        await callback.answer("Комбинация не найдена!", show_alert=True)
        return
    await callback.answer()
    try:
        # Спец-ветки
        if key == "screenshot":
            screenshot = pyautogui.screenshot()
            screenshot_path = APP_DIR / "screenshot.png"
            screenshot.save(screenshot_path)
            await bot.send_photo(chat_id=callback.from_user.id, photo=FSInputFile(screenshot_path))
            os.remove(screenshot_path)
            await callback.message.answer("Скриншот отправлен.")
            return

        if key == "screen_rec":
            # Тоггл записи Xbox Game Bar (Win+Alt+R).
            if not record_state['active']:
                pyautogui.hotkey('winleft', 'alt', 'r')
                record_state['active'] = True
                record_state['started_at'] = time.time()
                await callback.message.answer("🎥 Запись начата (Win+Alt+R). Повторное нажатие остановит запись.")
            else:
                pyautogui.hotkey('winleft', 'alt', 'r')
                record_state['active'] = False
                # Дадим системе дописать файл
                await asyncio.sleep(2.0)
                clip = _find_latest_clip(record_state['started_at'])
                user_id = callback.from_user.id
                last_clip_by_user[user_id] = clip
                if clip and clip.exists():
                    kb = InlineKeyboardBuilder()
                    kb.button(text="📤 Отправить в Telegram", callback_data="send_last_clip_yes")
                    kb.button(text="Оставить в папке", callback_data="send_last_clip_no")
                    kb.adjust(1, 1)
                    human_path = str(clip)
                    await callback.message.answer(
                        f"🟢 Запись остановлена.\nНашёл последний клип:\n<code>{human_path}</code>\nОтправить в Telegram?",
                        reply_markup=kb.as_markup(),
                        parse_mode="HTML"
                    )
                else:
                    # Не нашли клип — просто сообщим, где искать
                    dirs = _captures_dirs()
                    hint = "\n".join(str(d) for d in dirs) if dirs else "(Не удалось определить папку клипов)"
                    await callback.message.answer(
                        "🟡 Запись остановлена, но не удалось найти созданный файл автоматически.\n"
                        f"Проверьте папку клипов:\n{hint}"
                    )
            return

        if combo_info.get('type') == 'batch' and 'path' in combo_info:
            full_path = APP_DIR / combo_info['path']
            ext = full_path.suffix.lower()
            if ext in ['.bat', '.cmd']:
                subprocess.Popen(["cmd.exe", "/c", "start", "", str(full_path)], shell=True)
            elif ext == '.ps1':
                subprocess.Popen(["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(full_path)], shell=True)
            elif ext == '.py':
                subprocess.Popen(["cmd.exe", "/c", "start", "", "python", str(full_path)], shell=True)
            else:
                subprocess.Popen(["cmd.exe", "/c", "start", "", str(full_path)], shell=True)
            return

        elif combo_info.get('type') == 'set_search_browser':
            target_key = combo_info.get('target_browser_key')
            app_exists = any(app.get('key') == target_key for app in apps_data)
            if app_exists:
                user_data['preferred_search_browser_key'] = target_key
                save_config_setting('Settings', 'PREFERRED_SEARCH_BROWSER_KEY', target_key)
                browser_name = next((app['name'] for app in apps_data if app['key'] == target_key), target_key)
                await callback.message.answer(f"Выбран браузер для поиска: {browser_name}. Настройка сохранена.")
            else:
                await callback.message.answer("Браузер не найден в списке приложений!")
            return

        keys = combo_info.get('keys', [])
        if keys == ["alt_down"]:
            pyautogui.keyDown('alt'); return
        elif keys == ["alt_up"]:
            pyautogui.keyUp('alt'); return
        if keys == ["f"]:
            pyautogui.keyDown('f'); pyautogui.keyUp('f'); return

        if not keys:
            await callback.message.answer("Комбинация без клавиш не выполняет действий.")
            return

        layout = combo_info.get('layout', 'none')
        current_layout = get_current_layout()
        switched = False
        if layout != 'none' and layout != current_layout:
            switch_layout(layout); switched = True

        if keys == ['win', 'd']:
            show_desktop_toggle()
        elif 'win' in keys:
            pyautogui.hotkey(*keys)
        else:
            keyboard.send('+'.join(keys))

        if switched:
            switch_layout(current_layout)
    except Exception as e:
        await callback.message.answer(f"Ошибка выполнения: {e}")

# ===== Кнопки «Отправить клип в Telegram / Оставить в папке» =====

@dp.callback_query(F.data == "send_last_clip_yes")
async def send_last_clip_yes(callback: CallbackQuery):
    user_id = callback.from_user.id
    clip = last_clip_by_user.get(user_id)
    if not clip or not clip.exists():
        await callback.answer("Файл клипа не найден.", show_alert=True)
        return
    try:
        size_mb = clip.stat().st_size / (1024 * 1024)
        if size_mb > 2000:  # простая защита от слишком больших файлов
            await callback.message.answer(
                f"Файл слишком большой для отправки через бота ({size_mb:.0f} МБ). Оставляю в папке:\n<code>{clip}</code>",
                parse_mode="HTML"
            )
            await callback.answer()
            return
        await bot.send_video(chat_id=user_id, video=FSInputFile(clip), caption=f"🎥 Клип: {clip.name}")
        await callback.message.answer("Готово! Клип отправлен в Telegram.")
        await callback.answer()
    except Exception as e:
        await callback.message.answer(f"Не удалось отправить клип: {e}")
        await callback.answer()

@dp.callback_query(F.data == "send_last_clip_no")
async def send_last_clip_no(callback: CallbackQuery):
    user_id = callback.from_user.id
    clip = last_clip_by_user.get(user_id)
    if clip:
        await callback.message.answer(f"Оставил файл в папке:\n<code>{clip}</code>", parse_mode="HTML")
    else:
        await callback.message.answer("Оставил как есть. (Файл не определён)")
    await callback.answer()

# ==========================
# СИСТЕМНЫЕ ДЕЙСТВИЯ
# ==========================

async def set_commands():
    commands = [
        BotCommand(command="start", description="Запустить/перезапустить бота"),
        BotCommand(command="reload", description="Обновить данные из JSON"),
        BotCommand(command="end", description="Остановить бота"),
        BotCommand(command="editapps", description="Показать apps.json"),
        BotCommand(command="saveapps", description="Сохранить новый apps.json"),
        BotCommand(command="editcombos", description="Показать combos.json"),
        BotCommand(command="savecombos", description="Сохранить новый combos.json"),
        BotCommand(command="set_search_yandex", description="Использовать Яндекс для поиска"),
        BotCommand(command="set_search_google", description="Использовать Google для поиска"),
        BotCommand(command="set_search_bing", description="Использовать Bing для поиска")
    ]
    try:
        await bot.delete_my_commands()
        await bot.set_my_commands(commands)
    except Exception as e:
        logging.error(f"Ошибка при установке команд меню: {e}")

@dp.callback_query(F.data == "noop")
async def noop_callback(callback: CallbackQuery):
    await callback.answer()

@dp.message(Command("set_search_yandex"))
async def set_search_yandex(message: Message):
    if not has_access(message): return
    user_data['search_engine'] = 'yandex'
    save_config_setting('Settings', 'DEFAULT_SEARCH_ENGINE', 'yandex')
    await message.answer("Выбран поисковик: Яндекс. Настройка сохранена.")

@dp.message(Command("set_search_google"))
async def set_search_google(message: Message):
    if not has_access(message): return
    user_data['search_engine'] = 'google'
    save_config_setting('Settings', 'DEFAULT_SEARCH_ENGINE', 'google')
    await message.answer("Выбран поисковик: Google. Настройка сохранена.")

@dp.message(Command("set_search_bing"))
async def set_search_bing(message: Message):
    if not has_access(message): return
    user_data['search_engine'] = 'bing'
    save_config_setting('Settings', 'DEFAULT_SEARCH_ENGINE', 'bing')
    await message.answer("Выбран поисковик: Bing. Настройка сохранена.")

# ==========================
# КЛАВИАТУРНАЯ РАСКЛАДКА / ГРОМКОСТЬ / МЕДИА-КЛАВИШИ
# ==========================

LANGS = {'en': 0x409, 'ru': 0x419}

def get_current_layout():
    user32 = ctypes.WinDLL('user32', use_last_error=True)
    hwnd = user32.GetForegroundWindow()
    thread_id = user32.GetWindowThreadProcessId(hwnd, 0)
    layout_id = user32.GetKeyboardLayout(thread_id)
    lang = layout_id & 0xFFFF
    if lang == LANGS['en']:
        return 'en'
    elif lang == LANGS['ru']:
        return 'ru'
    else:
        return 'unknown'

def switch_layout(target):
    user32 = ctypes.WinDLL('user32', use_last_error=True)
    hwnd = user32.GetForegroundWindow()
    n = user32.GetKeyboardLayoutList(0, None)
    buf = (ctypes.c_ulong * n)()
    user32.GetKeyboardLayoutList(n, buf)
    target_hkl = None
    for hkl in buf:
        if (hkl & 0xFFFF) == LANGS[target]:
            target_hkl = hkl
            break
    if target_hkl:
        user32.ActivateKeyboardLayout(target_hkl, 0)
        time.sleep(0.1)

def volume_up():
    VK_VOLUME_UP = 0xAF
    user32 = ctypes.WinDLL('user32', use_last_error=True)
    user32.keybd_event(VK_VOLUME_UP, 0, 0, 0)
    user32.keybd_event(VK_VOLUME_UP, 0, 2, 0)

def volume_down():
    VK_VOLUME_DOWN = 0xAE
    user32 = ctypes.WinDLL('user32', use_last_error=True)
    user32.keybd_event(VK_VOLUME_DOWN, 0, 0, 0)
    user32.keybd_event(VK_VOLUME_DOWN, 0, 2, 0)

def volume_mute():
    VK_VOLUME_MUTE = 0xAD
    user32 = ctypes.WinDLL('user32', use_last_error=True)
    user32.keybd_event(VK_VOLUME_MUTE, 0, 0, 0)
    user32.keybd_event(VK_VOLUME_MUTE, 0, 2, 0)

def media_play_pause():
    VK_MEDIA_PLAY_PAUSE = 0xB3
    user32 = ctypes.WinDLL('user32', use_last_error=True)
    user32.keybd_event(VK_MEDIA_PLAY_PAUSE, 0, 0, 0)
    user32.keybd_event(VK_MEDIA_PLAY_PAUSE, 0, 2, 0)

def media_next():
    VK_MEDIA_NEXT_TRACK = 0xB0
    user32 = ctypes.WinDLL('user32', use_last_error=True)
    user32.keybd_event(VK_MEDIA_NEXT_TRACK, 0, 0, 0)
    user32.keybd_event(VK_MEDIA_NEXT_TRACK, 0, 2, 0)

def media_prev():
    VK_MEDIA_PREV_TRACK = 0xB1
    user32 = ctypes.WinDLL('user32', use_last_error=True)
    user32.keybd_event(VK_MEDIA_PREV_TRACK, 0, 0, 0)
    user32.keybd_event(VK_MEDIA_PREV_TRACK, 0, 2, 0)

def show_desktop_toggle():
    try:
        subprocess.run(
            ['powershell', '-Command', '(New-Object -ComObject Shell.Application).MinimizeAll()'],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
    except Exception as e:
        logging.error(f"Win+D via PowerShell error: {e}")

# ==========================
# ПРОЧИЕ СООБЩЕНИЯ / ПОИСК
# ==========================

@dp.message()
async def handle_other_messages(message: Message):
    if not has_access(message):
        return
    text = message.text or ""
    url_match = re.search(r'(https?://\S+)', text)
    if url_match:
        url = url_match.group(1)
        try:
            os.startfile(url)
            await message.reply(f"Ссылка открыта: {url}")
        except Exception as e:
            await message.reply(f"Ошибка открытия ссылки: {e}")
        return
    if text.strip() and not text.startswith('/'):
        search_engine_key = user_data.get('search_engine', 'google')
        search_template = SEARCH_ENGINES.get(search_engine_key, SEARCH_ENGINES['google'])
        query_encoded = urllib.parse.quote_plus(text)
        search_url = f"{search_template}{query_encoded}"
        preferred_browser_key = user_data.get('preferred_search_browser_key')
        if preferred_browser_key:
            browser_app_info = next((app for app in apps_data if app['key'] == preferred_browser_key), None)
            if browser_app_info and str(browser_app_info.get('is_app','y')).lower() == 'y':
                browser_path = _resolve_path(browser_app_info['path'])
                browser_args = _as_list(browser_app_info.get('args')) or _as_list(browser_app_info.get('arg'))
                try:
                    subprocess.Popen([browser_path] + browser_args + [search_url], shell=False)
                    await message.reply(f"Ищу в {browser_app_info['name']}: {text}")
                    return
                except Exception as e:
                    await message.reply(f"Ошибка открытия в {browser_app_info['name']}: {e}. Открываю в браузере по умолчанию.")
                    user_data['preferred_search_browser_key'] = None
                    save_config_setting('Settings', 'PREFERRED_SEARCH_BROWSER_KEY', '')
        try:
            os.startfile(search_url)
            await message.reply(f"Ищу в браузере по умолчанию: {text}")
        except Exception as e:
            await message.reply(f"Ошибка выполнения поиска: {e}")

# ==========================
# ЗАПУСК ПОЛЛИНГА
# ==========================

async def main():
    await set_commands()
    await dp.start_polling(bot)

if __name__ == '__main__':
    try:
        asyncio.run(main())
    except (KeyboardInterrupt, SystemExit):
        logging.info("Бот остановлен.")
