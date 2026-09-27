import os
import re
import json
import asyncio
import tempfile
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from dotenv import load_dotenv

from aiogram import Bot, Dispatcher, Router, F
from aiogram.types import Message, FSInputFile
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.filters import CommandStart

# Загрузка переменных окружения
load_dotenv()

BOT_TOKEN = os.getenv('BOT_TOKEN')
OPENROUTER_API_KEY = os.getenv('OPENROUTER_API_KEY')
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"

# !!! ОБНОВЛЕННАЯ МОДЕЛЬ !!!
# Если OpenRouter выдаст ошибку "model not found", замените на "deepseek/deepseek-chat"
MODEL_NAME = "deepseek/deepseek-v4.1-flash" 

# Инициализация бота и диспетчера
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
router = Router()
dp.include_router(router)

# --- FSM (Конечный автомат) для пошагового сбора данных ---
class DirectGenState(StatesGroup):
    theme = State()
    keywords = State()
    url = State()

# --- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ---

STOP_WORDS = {"бесплатно", "скачать", "видео", "торрент", "халява", "взлом", "кряк", "реферат"}

def clean_text(text: str, max_length: int) -> str:
    if not text:
        return ""
    # Удаляем опасные символы по краям
    cleaned = re.sub(r'^[\s\"\'\-\*\•\n]+|[\s\"\'\-\*\•\n]+$', '', str(text))
    # Нормализуем пробелы
    cleaned = re.sub(r'\s+', ' ', cleaned)
    # Строгая обрезка по символам (Яндекс.Директ считает именно символы)
    return cleaned[:max_length].strip()

def filter_stop_words(keywords_text: str) -> tuple[list[str], list[str]]:
    lines = [line.strip() for line in keywords_text.split('\n') if line.strip()]
    valid_keywords = []
    found_stops = []
    for line in lines:
        words = set(re.findall(r'\b\w+\b', line.lower()))
        intersection = words.intersection(STOP_WORDS)
        if intersection:
            found_stops.extend(list(intersection))
        else:
            valid_keywords.append(line)
    return valid_keywords, list(set(found_stops))

# Синхронная функция для запуска в asyncio.to_thread (не блокирует цикл событий бота)
def sync_generate_and_create_excel(theme: str, valid_keywords: list[str], url: str) -> str:
    # Берем первые 30 ключей для контекста, чтобы уложиться в лимиты токенов и не перегружать промпт
    keywords_context = ", ".join(valid_keywords[:30])
    
    prompt = f"""Ты - эксперт по Яндекс.Директ.
Тематика: {theme}
Ключевые слова: {keywords_context}
Ссылка: {url}

Верни СТРОГО валидный JSON-объект без markdown-оберток (без ```json и без пояснений):
{{
  "headlines": ["Заголовок 1", "Заголовок 2", "Заголовок 3", "Заголовок 4", "Заголовок 5", "Заголовок 6", "Заголовок 7"],
  "texts": ["Текст 1", "Текст 2", "Текст 3"],
  "refinements": ["Уточнение 1", "Уточнение 2", "Уточнение 3", "Уточнение 4"]
}}

ЖЕСТКИЕ ПРАВИЛА:
1. Заголовки: макс. 56 знаков. Цепляющие, с ключевыми словами.
2. Тексты: макс. 81 знак. Содержат призыв к действию (CTA).
3. Уточнения: макс. 25 знаков. Короткие преимущества (напр. "Гарантия 5 лет").
4. Не используй в начале или конце кавычки, тире, восклицательные знаки подряд.
5. Язык: русский."""

    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
        "HTTP-Referer": "https://t.me/your_bot", # Замените на реальное имя бота
        "X-Title": "Direct Adv Generator Bot"
    }
    
    payload = {
        "model": MODEL_NAME,
        "messages": [{"role": "user", "content": prompt}],
        "response_format": {"type": "json_object"} # DeepSeek отлично поддерживает этот формат
    }

    response = requests.post(OPENROUTER_URL, headers=headers, json=payload, timeout=45)
    response.raise_for_status()
    
    content = response.json()['choices'][0]['message']['content']
    
    # Страховка: очистка от markdown-оберток, если модель вдруг их добавит
    content = re.sub(r'^```json\s*|\s*```$', '', content.strip(), flags=re.MULTILINE)
    ai_data = json.loads(content)

    # Создание Excel файла
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
        wb = Workbook()
        ws = wb.active
        ws.title = "Объявления"
        
        headers_excel = [
            "Тип", "Название кампании", "Название группы", "Фраза", 
            "Заголовок", "Текст", "Ссылка", 
            "Уточнение 1", "Уточнение 2", "Уточнение 3", "Уточнение 4"
        ]
        ws.append(headers_excel)
        
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")

        headlines = ai_data.get('headlines', ["Нет заголовка"])
        texts = ai_data.get('texts', ["Нет текста"])
        refinements = ai_data.get('refinements', [""] * 4)

        # Разворачиваем ключи: каждая строка Excel = одна ключевая фраза
        for i, kw in enumerate(valid_keywords):
            h = clean_text(headlines[i % len(headlines)], 56)
            t = clean_text(texts[i % len(texts)], 81)
            
            row = [
                "Текстово-графическое объявление",
                clean_text(theme, 40),
                f"{clean_text(theme, 25)}_Гр1",
                kw,
                h,
                t,
                url,
                clean_text(refinements[0], 25) if len(refinements) > 0 else "",
                clean_text(refinements[1], 25) if len(refinements) > 1 else "",
                clean_text(refinements[2], 25) if len(refinements) > 2 else "",
                clean_text(refinements[3], 25) if len(refinements) > 3 else ""
            ]
            ws.append(row)

        # Автоширина столбцов для удобства просмотра в Excel
        for column in ws.columns:
            max_length = max((len(str(cell.value)) if cell.value else 0 for cell in column), default=0)
            ws.column_dimensions[column[0].column_letter].width = min(max_length + 2, 50)

        wb.save(tmp_file.name)
        return tmp_file.name

# --- ОБРАБОТЧИКИ (HANDLERS) ---

@router.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    text = (
        "👋 Привет! Я **Генератор объявлений Яндекс.Директ**.\n\n"
        "⚠️ *Важно:* Я использую ИИ (DeepSeek) для генерации. Всегда проверяйте готовые объявления перед загрузкой в Директ Коммандер.\n\n"
        "Давайте начнем! Напишите **тематику** или название вашей кампании (например: *Ремонт квартир в Москве*):"
    )
    await message.answer(text, parse_mode="Markdown")
    await state.set_state(DirectGenState.theme)

@router.message(DirectGenState.theme)
async def process_theme(message: Message, state: FSMContext):
    await state.update_data(theme=message.text.strip())
    await message.answer(
        "Отлично! Теперь отправьте список **ключевых запросов**.\n"
        "Пишите каждый запрос с новой строки.\n"
        "💡 *Совет:* Для лучшего качества и скорости отправляйте не более 50 ключей за один раз."
    )
    await state.set_state(DirectGenState.keywords)

@router.message(DirectGenState.keywords)
async def process_keywords(message: Message, state: FSMContext):
    valid_keywords, found_stops = filter_stop_words(message.text)
    
    if not valid_keywords:
        await message.answer(
            "❌ Все ключевые слова содержат запрещенные стоп-слова или введены некорректно. "
            "Пожалуйста, введите список ключевых запросов заново."
        )
        return

    await state.update_data(keywords=valid_keywords)
    
    if found_stops:
        await message.answer(
            f"⚠️ Обнаружены и автоматически удалены стоп-слова: `{', '.join(found_stops)}`.\n"
            f"К обработке принято {len(valid_keywords)} ключевых фраз."
        )
    
    await message.answer("Принято! Теперь отправьте **ссылку на сайт** (посадочную страницу), например: `https://example.com`")
    await state.set_state(DirectGenState.url)

@router.message(DirectGenState.url)
async def process_url(message: Message, state: FSMContext, bot: Bot):
    url = message.text.strip()
    if not url.startswith("http"):
        await message.answer("⚠️ Ссылка должна начинаться с http:// или https://. Попробуйте еще раз:")
        return

    data = await state.get_data()
    theme = data['theme']
    keywords = data['keywords']

    # Отправляем сообщение о начале генерации
    processing_msg = await message.answer("⏳ ИИ анализирует тематику и генерирует объявления... Это займет 10-20 секунд.")

    try:
        # Запускаем тяжелую операцию в отдельном потоке, чтобы бот не "зависал"
        file_path = await asyncio.to_thread(sync_generate_and_create_excel, theme, keywords, url)
        
        # Отправляем файл пользователю
        filename = f"Direct_{theme[:15].replace(' ', '_')}.xlsx"
        input_file = FSInputFile(file_path, filename=filename)
        
        caption = (
            f"✅ Готово!\n"
            f"Сгенерировано объявлений для {len(keywords)} фраз.\n"
            f"Файл полностью совместим с импортом в Директ Коммандер."
        )
        await message.answer_document(document=input_file, caption=caption)
        
        # Очищаем временный файл после успешной отправки (безопасность и экономия места)
        os.remove(file_path)
        
        # Предлагаем начать заново
        await message.answer("🔄 Хотите создать новую кампанию? Отправьте команду /start")
        
    except Exception as e:
        error_text = str(e)
        if "model not found" in error_text.lower():
            error_text = "Ошибка модели ИИ. Проверьте название модели в коде (возможно, стоит использовать 'deepseek/deepseek-chat')."
            
        await processing_msg.edit_text(f"❌ Произошла ошибка при генерации: {error_text}\nПопробуйте еще раз или сократите список ключей.")
    
    finally:
        await state.clear()

# --- ЗАПУСК ---
if __name__ == "__main__":
    import logging
    logging.basicConfig(level=logging.INFO)
    print(f"Бот запущен с моделью: {MODEL_NAME}")
    asyncio.run(dp.start_polling(bot))
