import asyncio
import logging
import os
import tempfile

from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.types import Message, FSInputFile
from dotenv import load_dotenv

from processor import process_docx

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

WELCOME_TEXT = (
    "Привет! Я подготавливаю тендерные технические спецификации.\n\n"
    "Отправь файл .docx с сырым ТЗ — я:\n"
    "• уберу фразы «не менее / не более»\n"
    "• переведу требования в жёсткую форму\n"
    "• поставлю маркеры [ВСТАВИТЬ...] для товаров\n\n"
    "Просто прикрепи .docx файл к сообщению."
)


@dp.message(Command("start"))
async def cmd_start(message: Message):
    await message.answer(WELCOME_TEXT)


@dp.message(F.document)
async def handle_document(message: Message):
    doc = message.document

    if not doc.file_name.lower().endswith(".docx"):
        await message.answer("Пожалуйста, отправь файл в формате .docx")
        return

    if doc.file_size > 20 * 1024 * 1024:
        await message.answer("Файл слишком большой. Максимум — 20 МБ.")
        return

    status = await message.answer("Обрабатываю файл...")

    with tempfile.TemporaryDirectory() as tmp:
        input_path = os.path.join(tmp, "input.docx")
        output_path = os.path.join(tmp, "output.docx")

        file_info = await bot.get_file(doc.file_id)
        await bot.download_file(file_info.file_path, destination=input_path)

        try:
            await status.edit_text("Анализирую и очищаю спецификацию...")

            process_docx(input_path, output_path, GEMINI_API_KEY)

            output_name = doc.file_name.replace(".docx", "_cleaned.docx")
            await status.edit_text("Готово!")
            await message.answer_document(
                FSInputFile(output_path, filename=output_name),
                caption=(
                    "Спецификация обработана.\n"
                    "Вставьте товары вместо маркеров [ВСТАВИТЬ...] и отправляйте на тендер."
                ),
            )

        except Exception as e:
            logger.error("Processing error", exc_info=True)
            await status.edit_text(
                "Ошибка при обработке файла. Попробуйте ещё раз."
            )


@dp.message()
async def handle_other(message: Message):
    await message.answer(WELCOME_TEXT)


async def main():
    logger.info("Bot starting...")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
