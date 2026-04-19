import asyncio
import os
from aiohttp import web
from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, ReplyKeyboardMarkup, KeyboardButton, FSInputFile
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup

# Создаем папки
os.makedirs("Oborotka", exist_ok=True)
os.makedirs("ActSverka", exist_ok=True)
os.makedirs("FinHelp", exist_ok=True)
os.makedirs("MaterialReport", exist_ok=True)
os.makedirs("Svodka4010", exist_ok=True)  # НОВАЯ ПАПКА

from Oborotka.oborotka import process_oborotka_file
from ActSverka.actsverka import update_master_oborotka, process_reconciliation_acts
from FinHelp.finhelp import generate_finhelp_acts
from MaterialReport.material_logic import generate_material_report
from Svodka.svodka_logic import generate_svodka_4010  # ИМПОРТ 4010

bot = Bot(token=os.getenv('BOT_TOKEN'))
dp = Dispatcher()


# ==========================================
# СОСТОЯНИЯ (FSM)
# ==========================================
class BotStates(StatesGroup):
    wait_ob_files = State()
    wait_ob_month = State()
    wait_act_csv = State()
    wait_act_code = State()
    wait_act_db = State()
    wait_finhelp_ob = State()
    wait_material_file = State()

    # НОВЫЕ СОСТОЯНИЯ ДЛЯ СВОДКИ 4010
    wait_svodka_files = State()
    wait_svodka_month = State()
    wait_svodka_saldo = State()


# ==========================================
# КЛАВИАТУРЫ
# ==========================================
main_kb = ReplyKeyboardMarkup(keyboard=[
    [KeyboardButton(text="📊 Оборотка"), KeyboardButton(text="📑 Акт сверки")],
    [KeyboardButton(text="📈 Сводка по счетам 4010"), KeyboardButton(text="💰 Фин. помощь")],  # Изменено название кнопки
    [KeyboardButton(text="📦 Материальный отчет")]
], resize_keyboard=True)

# Клавиатура для сальдо
saldo_kb = ReplyKeyboardMarkup(keyboard=[
    [KeyboardButton(text="да"), KeyboardButton(text="нет")],
    [KeyboardButton(text="❌ Отмена")]
], resize_keyboard=True)

# ... (остальные клавиатуры остаются без изменений: act_kb, files_kb, cancel_kb, code_kb)
act_kb = ReplyKeyboardMarkup(keyboard=[
    [KeyboardButton(text="⚙️ Сделать Акт"), KeyboardButton(text="📂 Обновить базу")],
    [KeyboardButton(text="🔙 Назад")]
], resize_keyboard=True)
files_kb = ReplyKeyboardMarkup(keyboard=[
    [KeyboardButton(text="✅ Готово"), KeyboardButton(text="❌ Отмена")]
], resize_keyboard=True)
cancel_kb = ReplyKeyboardMarkup(keyboard=[[KeyboardButton(text="❌ Отмена")]], resize_keyboard=True)
code_kb = ReplyKeyboardMarkup(keyboard=[
    [KeyboardButton(text="4010"), KeyboardButton(text="6010")],
    [KeyboardButton(text="❌ Отмена")]
], resize_keyboard=True)


# ... (Остальные хендлеры: cmd_start, cancel_action, start_ob, start_act, start_finhelp, start_material_report - ОСТАЮТСЯ КАК ЕСТЬ)
# ВСТАВЬ ИХ СЮДА ИЗ ПРОШЛОГО КОДА


# ==========================================
# РАЗДЕЛ: СВОДКА 4010
# ==========================================
@dp.message(F.text == "📈 Сводка по счетам 4010")
async def start_svodka_4010(message: Message, state: FSMContext):
    # Теперь храним файлы просто списком
    await state.update_data(svodka_files_list=[])
    await message.answer(
        "Вы в разделе Сводка 4010.\n"
        "Отправьте мне любые 3 файла (.xlsx):\n"
        "• Исходник\n"
        "• Общую сводку\n"
        "• Реестр\n\n"
        "Порядок и названия файлов значения не имеют. "
        "Как загрузите все три, нажмите '✅ Готово'.",
        reply_markup=files_kb
    )
    await state.set_state(BotStates.wait_svodka_files)


@dp.message(BotStates.wait_svodka_files, F.document)
async def receive_svodka_files(message: Message, state: FSMContext):
    if not message.document.file_name.endswith('.xlsx'):
        await message.answer("Нужны только .xlsx файлы!")
        return

    path = os.path.join("Svodka4010", f"temp_{message.document.file_id}.xlsx")
    await bot.download(message.document, destination=path)

    data = await state.get_data()
    files_list = data.get('svodka_files_list', [])
    files_list.append(path)

    await state.update_data(svodka_files_list=files_list)

    count = len(files_list)
    await message.answer(f"📥 Файл №{count} загружен: `{message.document.file_name}`", parse_mode="Markdown")

    if count == 3:
        await message.answer("Отлично! Все 3 файла получены. Нажмите '✅ Готово'.")


@dp.message(BotStates.wait_svodka_files, F.text == "✅ Готово")
async def svodka_files_done(message: Message, state: FSMContext):
    data = await state.get_data()
    files_list = data.get('svodka_files_list', [])

    if len(files_list) != 3:
        await message.answer(f"Нужно загрузить ровно 3 файла. Сейчас загружено: {len(files_list)}")
        return

    await message.answer("Введите номер месяца, с которого начать (например, 04 для Апреля):", reply_markup=cancel_kb)
    await state.set_state(BotStates.wait_svodka_month)


@dp.message(BotStates.wait_svodka_month)
async def svodka_get_month(message: Message, state: FSMContext):
    try:
        month_idx = int(message.text.strip()) - 1
        if not (0 <= month_idx <= 11): raise ValueError
    except:
        await message.answer("Пожалуйста, введите число от 1 до 12.")
        return

    await state.update_data(sv_month=month_idx)
    await message.answer("Вы переносили сальдо в исходном файле в столбцы C и D?", reply_markup=saldo_kb)
    await state.set_state(BotStates.wait_svodka_saldo)


@dp.message(BotStates.wait_svodka_saldo, F.text.in_({"да", "нет"}))
async def svodka_get_saldo(message: Message, state: FSMContext):
    is_transferred = (message.text == "да")
    data = await state.get_data()

    await message.answer("⚙️ Запускаю умную классификацию и расчет... Это может занять до минуты.",
                         reply_markup=main_kb)

    # Вызываем новую функцию, передавая СПИСОК путей
    success, result = await asyncio.to_thread(
        generate_svodka_4010,
        data['svodka_files_list'],
        data['sv_month'],
        is_transferred
    )

    # Удаляем временные файлы
    for f_path in data['svodka_files_list']:
        if os.path.exists(f_path):
            try:
                os.remove(f_path)
            except:
                pass

    if success:
        await message.answer_document(FSInputFile(result))
        try:
            os.remove(result)
        except:
            pass
        await message.answer("✅ Сводка успешно сформирована!")
    else:
        # Если возникла ошибка в логике классификации, выводим её пользователю
        await message.answer(f"❌ Ошибка при обработке:\n{result}")

    await state.clear()


# ==========================================
# ЗАПУСК СЕРВЕРА (ДЛЯ RENDER)
# ==========================================
async def handle(request):
    return web.Response(text="Бот жив и работает!")


async def web_server():
    app = web.Application()
    app.router.add_get('/', handle)
    runner = web.AppRunner(app)
    await runner.setup()
    port = int(os.environ.get("PORT", 8080))
    site = web.TCPSite(runner, '0.0.0.0', port)
    await site.start()


async def main():
    print("Запускаю веб-сервер для Render...")
    await web_server()
    print("Бот успешно запущен!")
    await dp.start_polling(bot)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("Бот остановлен.")