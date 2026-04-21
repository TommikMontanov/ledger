import asyncio
import os
from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, ReplyKeyboardMarkup, KeyboardButton, FSInputFile
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiohttp import web

# Создаем папки при старте
os.makedirs("Oborotka", exist_ok=True)
os.makedirs("ActSverka", exist_ok=True)
os.makedirs("FinHelp", exist_ok=True)
os.makedirs("MaterialReport", exist_ok=True) # НОВАЯ ПАПКА

# Импортируем функции
from Oborotka.oborotka import process_oborotka_file
from ActSverka.actsverka import update_master_oborotka, process_reconciliation_acts
from FinHelp.finhelp import generate_finhelp_acts
from MaterialReport.material_logic import generate_material_report # НОВЫЙ ИМПОРТ


bot = Bot(token='8604173225:AAEeRQqv5-yj5Ygd4sr2MgSp1F6hwc7gTDo')
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
    wait_material_file = State() # СОСТОЯНИЕ ДЛЯ МАТ. ОТЧЕТА


# ==========================================
# КЛАВИАТУРЫ
# ==========================================
main_kb = ReplyKeyboardMarkup(keyboard=[
    [KeyboardButton(text="📊 Оборотка"), KeyboardButton(text="📑 Акт сверки")],
    [KeyboardButton(text="📈 Сводка по счетам"), KeyboardButton(text="💰 Фин. помощь")],
    [KeyboardButton(text="📦 Материальный отчет")] # НОВАЯ КНОПКА
], resize_keyboard=True)

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


# ==========================================
# ОБЩИЕ ОБРАБОТЧИКИ
# ==========================================
@dp.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    await message.answer("Привет! Выберите нужный раздел:", reply_markup=main_kb)

@dp.message(F.text == "🔙 Назад")
async def back_to_main(message: Message, state: FSMContext):
    await state.clear()
    await message.answer("Возврат в главное меню:", reply_markup=main_kb)

@dp.message(F.text == "❌ Отмена")
async def cancel_action(message: Message, state: FSMContext):
    await state.clear()
    await message.answer("Действие отменено. Возврат в меню.", reply_markup=main_kb)


# ==========================================
# РАЗДЕЛ: ОБОРОТКА
# ==========================================
@dp.message(F.text == "📊 Оборотка")
async def start_ob(message: Message, state: FSMContext):
    await state.update_data(files=[])
    await message.answer("Пришлите файлы банковской выписки (.xlsx).\nМожно отправить несколько штук. После загрузки нажмите '✅ Готово'.", reply_markup=files_kb)
    await state.set_state(BotStates.wait_ob_files)

@dp.message(BotStates.wait_ob_files, F.document)
async def get_ob_docs(message: Message, state: FSMContext):
    if not message.document.file_name.endswith('.xlsx'):
        await message.answer("Пожалуйста, отправьте файл в формате .xlsx")
        return
    path = os.path.join("Oborotka", f"temp_{message.from_user.id}_{message.document.file_id}.xlsx")
    await bot.download(message.document, destination=path)
    data = await state.get_data()
    files_list = data.get('files', [])
    files_list.append(path)
    await state.update_data(files=files_list)
    await message.answer(f"📥 Файл `{message.document.file_name}` загружен!", parse_mode="Markdown")

@dp.message(BotStates.wait_ob_files, F.text == "✅ Готово")
async def ob_done(message: Message, state: FSMContext):
    data = await state.get_data()
    if not data.get('files'):
        await message.answer("Вы не отправили ни одного файла!")
        return
    await message.answer("Введите номер месяца (например: 01, 02, 03... 12):")
    await state.set_state(BotStates.wait_ob_month)

@dp.message(BotStates.wait_ob_month)
async def ob_process(message: Message, state: FSMContext):
    month = message.text.strip()
    data = await state.get_data()
    await message.answer("⚙️ Начинаю обработку файлов...", reply_markup=main_kb)

    total_generated = 0
    for f in data['files']:
        try:
            res = await asyncio.to_thread(process_oborotka_file, f, month)
            for out in res:
                if os.path.exists(out):
                    await message.answer_document(FSInputFile(out))
                    os.remove(out)
                    total_generated += 1
        except Exception as e:
            await message.answer(f"❌ Ошибка при обработке: {e}")
        finally:
            if os.path.exists(f): os.remove(f)

    if total_generated == 0:
        await message.answer(f"В файлах не найдено данных за выбранный месяц ({month}).")
    await state.clear()


# ==========================================
# РАЗДЕЛ: АКТ СВЕРКИ
# ==========================================
@dp.message(F.text == "📑 Акт сверки")
async def start_act(message: Message):
    await message.answer("Вы в меню Актов сверки:", reply_markup=act_kb)

@dp.message(F.text == "📂 Обновить базу")
async def act_db_start(message: Message, state: FSMContext):
    await message.answer("Пришлите файл Оборотки (.xlsx) для загрузки в Базу Актов.\nПредыдущая база будет удалена и заменена новой.", reply_markup=cancel_kb)
    await state.set_state(BotStates.wait_act_db)

@dp.message(BotStates.wait_act_db, F.document)
async def act_db_get(message: Message, state: FSMContext):
    if not message.document.file_name.endswith('.xlsx'):
        await message.answer("Нужен только файл .xlsx!")
        return
    path = os.path.join("ActSverka", f"temp_db_{message.document.file_id}.xlsx")
    await bot.download(message.document, destination=path)
    await message.answer("⏳ Заменяю базу...")
    success, msg = await asyncio.to_thread(update_master_oborotka, path)
    await message.answer(msg, reply_markup=act_kb)
    if os.path.exists(path) and not success: os.remove(path)
    await state.clear()

@dp.message(F.text == "⚙️ Сделать Акт")
async def act_generate_start(message: Message, state: FSMContext):
    await state.update_data(saved_csvs=[])
    await message.answer("Отправьте счета-фактуры в формате **.csv**.\nМожно выделять и отправлять сразу пачку.\nКак закончите — нажмите '✅ Готово'.", reply_markup=files_kb, parse_mode="Markdown")
    await state.set_state(BotStates.wait_act_csv)

@dp.message(BotStates.wait_act_csv, F.document)
async def act_csv_receive(message: Message, state: FSMContext):
    if not message.document.file_name.lower().endswith('.csv'):
        await message.answer("Пожалуйста, отправьте файл формата .csv")
        return
    path = os.path.join("ActSverka", f"temp_csv_{message.from_user.id}_{message.document.file_id}.csv")
    await bot.download(message.document, destination=path)
    data = await state.get_data()
    csv_list = data.get('saved_csvs', [])
    csv_list.append(path)
    await state.update_data(saved_csvs=csv_list)
    await message.answer(f"📥 Счёт-фактура `{message.document.file_name}` загружена!", parse_mode="Markdown")

@dp.message(BotStates.wait_act_csv, F.text == "✅ Готово")
async def act_csv_done(message: Message, state: FSMContext):
    data = await state.get_data()
    saved_csvs = data.get('saved_csvs', [])

    if not saved_csvs:
        await message.answer("Вы не отправили ни одного CSV файла.")
        return

    await message.answer("Выберите код (4010 или 6010):", reply_markup=code_kb)
    await state.set_state(BotStates.wait_act_code)

@dp.message(BotStates.wait_act_code, F.text.in_({"4010", "6010"}))
async def act_process_code(message: Message, state: FSMContext):
    code = message.text.strip()
    data = await state.get_data()
    saved_csvs = data.get('saved_csvs', [])

    await message.answer(f"⚙️ Выбран код {code}. Склеиваю инвойсы и формирую Акты...", reply_markup=act_kb)

    try:
        output_files = await asyncio.to_thread(process_reconciliation_acts, saved_csvs, code)

        if output_files:
            for out_file in output_files:
                if os.path.exists(out_file):
                    await message.answer_document(FSInputFile(out_file))
                    os.remove(out_file)
            await message.answer("✅ Все Акты сверки успешно сформированы!")
        else:
            await message.answer("⚠️ Акты не сформированы. Возможно, в базе нет совпадений по ИНН.")

    except Exception as e:
        await message.answer(f"❌ Произошла ошибка при формировании: {e}")

    finally:
        for f in saved_csvs:
            if os.path.exists(f): os.remove(f)
        await state.clear()


# ==========================================
# РАЗДЕЛ: ФИН. ПОМОЩЬ
# ==========================================
@dp.message(F.text == "💰 Фин. помощь")
async def start_finhelp(message: Message, state: FSMContext):
    await message.answer(
        "Вы в разделе Фин. Помощи.\n"
        "Пожалуйста, отправьте файл Оборотки (.xlsx).\n\n"
        "⚠️ ВАЖНО: В файле должен быть строго **1 лист**!",
        reply_markup=cancel_kb
    )
    await state.set_state(BotStates.wait_finhelp_ob)

@dp.message(BotStates.wait_finhelp_ob, F.document)
async def process_finhelp_file(message: Message, state: FSMContext):
    if not message.document.file_name.endswith('.xlsx'):
        await message.answer("Пожалуйста, отправьте файл в формате .xlsx")
        return
        
    path = os.path.join("FinHelp", f"temp_ob_{message.from_user.id}.xlsx")
    await bot.download(message.document, destination=path)
    
    await message.answer("⚙️ Проверяю файл и формирую акты, подождите...", reply_markup=main_kb)
    
    success, result = await asyncio.to_thread(generate_finhelp_acts, path)
    
    if os.path.exists(path):
        os.remove(path)
        
    if success:
        await message.answer_document(FSInputFile(result))
        os.remove(result) 
        await message.answer("✅ Акты по фин. помощи успешно сформированы!")
    else:
        await message.answer(result)
        
    await state.clear()


# ==========================================
# РАЗДЕЛ: МАТЕРИАЛЬНЫЙ ОТЧЕТ
# ==========================================
@dp.message(F.text == "📦 Материальный отчет")
async def start_material_report(message: Message, state: FSMContext):
    await message.answer(
        "Вы в разделе Материальный отчет.\n"
        "Пожалуйста, отправьте файл с выгрузкой материалов (.xlsx).",
        reply_markup=cancel_kb
    )
    await state.set_state(BotStates.wait_material_file)

@dp.message(BotStates.wait_material_file, F.document)
async def process_material_file(message: Message, state: FSMContext):
    if not message.document.file_name.endswith('.xlsx'):
        await message.answer("Пожалуйста, отправьте файл в формате .xlsx")
        return
        
    path = os.path.join("MaterialReport", f"temp_mat_{message.from_user.id}_{message.document.file_id}.xlsx")
    await bot.download(message.document, destination=path)
    
    await message.answer("⚙️ Анализирую поставщиков и формирую отчет...", reply_markup=main_kb)
    
    # Запускаем логику материального отчета
    success, result = await asyncio.to_thread(generate_material_report, path)
    
    # Удаляем временный файл выгрузки
    if os.path.exists(path):
        os.remove(path)
        
    if success:
        await message.answer_document(FSInputFile(result))
        os.remove(result) 
        await message.answer("✅ Материальный отчет успешно сформирован!")
    else:
        await message.answer(result) # Выведет ошибку, если что-то не так
        
    await state.clear()


# ==========================================
# ЗАПУСК БОТА
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
        



    
