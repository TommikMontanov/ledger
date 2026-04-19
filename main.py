import asyncio
import os
import re
from aiohttp import web
from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, ReplyKeyboardMarkup, KeyboardButton, FSInputFile
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup

# --- ИМПОРТЫ ТВОИХ МОДУЛЕЙ ---
# Убедись, что папки и файлы существуют
from Oborotka.oborotka import process_oborotka_file
from ActSverka.actsverka import update_master_oborotka, process_reconciliation_acts
from FinHelp.finhelp import generate_finhelp_acts
from MaterialReport.material_logic import generate_material_report
from Svodka.svodka_logic import generate_svodka_4010  # Твоя новая функция

# Создаем необходимые папки
for folder in ["Oborotka", "ActSverka", "FinHelp", "MaterialReport", "Svodka4010"]:
    os.makedirs(folder, exist_ok=True)

# Инициализация бота
TOKEN = os.getenv('BOT_TOKEN')
if not TOKEN:
    print("❌ ОШИБКА: Переменная окружения BOT_TOKEN не установлена!")
bot = Bot(token=TOKEN)
dp = Dispatcher()


# ==========================================
# СОСТОЯНИЯ (FSM)
# ==========================================
class BotStates(StatesGroup):
    # Состояния для других модулей
    wait_ob_files = State()
    wait_ob_month = State()
    wait_act_csv = State()
    wait_act_code = State()
    wait_act_db = State()
    wait_finhelp_ob = State()
    wait_material_file = State()

    # Состояния для Сводки 4010
    wait_svodka_files = State()
    wait_svodka_month = State()
    wait_svodka_saldo = State()


# ==========================================
# КЛАВИАТУРЫ
# ==========================================
main_kb = ReplyKeyboardMarkup(keyboard=[
    [KeyboardButton(text="📊 Оборотка"), KeyboardButton(text="📑 Акт сверки")],
    [KeyboardButton(text="📈 Сводка по счетам 4010"), KeyboardButton(text="💰 Фин. помощь")],
    [KeyboardButton(text="📦 Материальный отчет")]
], resize_keyboard=True)

files_kb = ReplyKeyboardMarkup(keyboard=[
    [KeyboardButton(text="✅ Готово"), KeyboardButton(text="❌ Отмена")]
], resize_keyboard=True)

saldo_kb = ReplyKeyboardMarkup(keyboard=[
    [KeyboardButton(text="да"), KeyboardButton(text="нет")],
    [KeyboardButton(text="❌ Отмена")]
], resize_keyboard=True)

cancel_kb = ReplyKeyboardMarkup(keyboard=[[KeyboardButton(text="❌ Отмена")]], resize_keyboard=True)


# ==========================================
# ОБЩИЕ ХЕНДЛЕРЫ
# ==========================================
@dp.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    await message.answer("Привет! Я бухгалтерский помощник. Выберите нужный раздел:", reply_markup=main_kb)


@dp.message(F.text == "❌ Отмена")
async def cancel_action(message: Message, state: FSMContext):
    data = await state.get_data()
    # Удаляем временные файлы, если они были
    files = data.get('svodka_files_list', [])
    for f in files:
        if os.path.exists(f): os.remove(f)

    await state.clear()
    await message.answer("Действие отменено.", reply_markup=main_kb)


# ==========================================
# РАЗДЕЛ: СВОДКА 4010 (ОБНОВЛЕННЫЙ)
# ==========================================
@dp.message(F.text == "📈 Сводка по счетам 4010")
async def start_svodka_4010(message: Message, state: FSMContext):
    await state.update_data(svodka_files_list=[])
    await message.answer(
        "📊 **Раздел Сводка 4010**\n\n"
        "Отправьте мне **3 файла** (.xlsx):\n"
        "1. Исходник (Сводка по счетам)\n"
        "2. Общую сводку\n"
        "3. Реестр\n\n"
        "Названия файлов не важны, я распознаю их по содержанию.\n"
        "После загрузки нажмите '✅ Готово'.",
        reply_markup=files_kb,
        parse_mode="Markdown"
    )
    await state.set_state(BotStates.wait_svodka_files)


@dp.message(BotStates.wait_svodka_files, F.document)
async def receive_svodka_files(message: Message, state: FSMContext):
    if not message.document.file_name.endswith('.xlsx'):
        await message.answer("⚠️ Ошибка: принимаются только файлы формата .xlsx")
        return

    # Сохраняем файл во временную папку
    file_id = message.document.file_id
    path = os.path.join("Svodka4010", f"temp_{file_id}.xlsx")
    await bot.download(message.document, destination=path)

    data = await state.get_data()
    files_list = data.get('svodka_files_list', [])
    files_list.append(path)
    await state.update_data(svodka_files_list=files_list)

    await message.answer(f"📥 Файл №{len(files_list)} получен: `{message.document.file_name}`", parse_mode="Markdown")


@dp.message(BotStates.wait_svodka_files, F.text == "✅ Готово")
async def svodka_files_done(message: Message, state: FSMContext):
    data = await state.get_data()
    files = data.get('svodka_files_list', [])

    if len(files) != 3:
        await message.answer(
            f"⚠️ Вы загрузили {len(files)} файла(ов), а нужно ровно 3. Продолжайте загрузку или нажмите 'Отмена'.")
        return

    await message.answer("Введите номер месяца, с которого начинается отчет (например, 1 — Январь, 4 — Апрель):",
                         reply_markup=cancel_kb)
    await state.set_state(BotStates.wait_svodka_month)


@dp.message(BotStates.wait_svodka_month)
async def svodka_get_month(message: Message, state: FSMContext):
    month_text = message.text.strip()
    if not month_text.isdigit() or not (1 <= int(month_text) <= 12):
        await message.answer("🔢 Пожалуйста, введите число от 1 до 12.")
        return

    await state.update_data(sv_month=int(month_text) - 1)
    await message.answer("Вы уже переносили сальдо вручную в столбцы C и D исходника?", reply_markup=saldo_kb)
    await state.set_state(BotStates.wait_svodka_saldo)


@dp.message(BotStates.wait_svodka_saldo, F.text.in_({"да", "нет"}))
async def svodka_get_saldo(message: Message, state: FSMContext):
    is_transferred = (message.text == "да")
    data = await state.get_data()

    await message.answer("⚙️ Начинаю магию классификации и расчет... Это займет немного времени.", reply_markup=main_kb)

    # Выполняем тяжелую функцию в отдельном потоке, чтобы бот не завис
    success, result = await asyncio.to_thread(
        generate_svodka_4010,
        data['svodka_files_list'],
        data['sv_month'],
        is_transferred
    )

    # Удаляем временные файлы
    for f_path in data['svodka_files_list']:
        if os.path.exists(f_path): os.remove(f_path)

    if success:
        await message.answer_document(
            FSInputFile(result),
            caption="✅ Сводка 4010 готова!"
        )
        if os.path.exists(result): os.remove(result)
    else:
        await message.answer(f"❌ Произошла ошибка:\n{result}")

    await state.clear()


# ==========================================
# ЗАПУСК СЕРВЕРА (ДЛЯ RENDER / KEEP-ALIVE)
# ==========================================
async def handle(request):
    return web.Response(text="Bot is running!")


async def web_server():
    app = web.Application()
    app.router.add_get('/', handle)
    runner = web.AppRunner(app)
    await runner.setup()
    port = int(os.environ.get("PORT", 8080))
    site = web.TCPSite(runner, '0.0.0.0', port)
    await site.start()


async def main():
    # Запуск веб-сервера параллельно с ботом
    asyncio.create_task(web_server())
    print("🚀 Бот и Веб-сервер запущены...")
    await dp.start_polling(bot)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("⭕ Бот остановлен.")