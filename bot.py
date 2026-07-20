#!/usr/bin/env python3
"""
PDF Converter Telegram Bot
Fayllarni (rasm, matn, hujjat) PDF ga aylantiradi.
"""

import os
import io
import logging
from datetime import datetime
from dotenv import load_dotenv

from PIL import Image
import reportlab
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import img2pdf

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ConversationHandler,
    filters,
    ContextTypes,
)

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

load_dotenv()

TOKEN = os.getenv("BOT_TOKEN", "YOUR_BOT_TOKEN_HERE")

WAITING_FILE = 1

ALLOWED_EXTENSIONS = {
    "image": [".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp", ".tiff"],
    "document": [".txt", ".docx", ".doc", ".odt", ".rtf"],
}


def is_image(filename):
    ext = os.path.splitext(filename.lower())[1]
    return ext in ALLOWED_EXTENSIONS["image"]


def is_document(filename):
    ext = os.path.splitext(filename.lower())[1]
    return ext in ALLOWED_EXTENSIONS["document"]


def images_to_pdf(image_paths, output_path):
    valid_paths = []
    for p in image_paths:
        if os.path.exists(p):
            valid_paths.append(p)

    if not valid_paths:
        return None

    if len(valid_paths) == 1:
        img = Image.open(valid_paths[0])
        if img.mode == "RGBA":
            img = img.convert("RGB")
            tmp_path = valid_paths[0] + ".tmp.jpg"
            img.save(tmp_path, "JPEG")
            valid_paths[0] = tmp_path

    with open(output_path, "wb") as f:
        f.write(img2pdf.convert(valid_paths))
    return output_path


def image_to_pdf_with_reportlab(image_path, output_path):
    img = Image.open(image_path)
    img_width, img_height = img.size

    c = canvas.Canvas(output_path, pagesize=A4)
    page_width, page_height = A4

    max_width = page_width - 2 * inch
    max_height = page_height - 2 * inch

    ratio = min(max_width / img_width, max_height / img_height)
    new_width = img_width * ratio
    new_height = img_height * ratio

    x = (page_width - new_width) / 2
    y = (page_height - new_height) / 2

    if img.mode == "RGBA":
        img = img.convert("RGB")
        tmp_path = image_path + ".tmp.jpg"
        img.save(tmp_path, "JPEG")
        c.drawImage(tmp_path, x, y, new_width, new_height)
        os.remove(tmp_path)
    else:
        c.drawImage(ImageReader(img), x, y, new_width, new_height)

    c.save()
    return output_path


def text_to_pdf(text_content, output_path):
    c = canvas.Canvas(output_path, pagesize=A4)
    page_width, page_height = A4

    text_obj = c.beginText(inch, page_height - inch)
    text_obj.setFont("Helvetica", 12)
    text_obj.setLeading(14)

    lines = text_content.split("\n")
    for line in lines:
        if text_obj.getY() < inch:
            c.drawText(text_obj)
            c.showPage()
            text_obj = c.beginText(inch, page_height - inch)
            text_obj.setFont("Helvetica", 12)
            text_obj.setLeading(14)
        text_obj.textLine(line)

    c.drawText(text_obj)
    c.save()
    return output_path


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    welcome = (
        f"Salom, {user.first_name}! \n\n"
        "Men PDF Converter botman.\n\n"
        "Menga quyidagi turdagi fayllarni yuboring:\n"
        "  - Rasmlar (jpg, png, bmp, gif, webp)\n"
        "  - Matn fayllari (.txt)\n\n"
        "Men ularni PDF formatiga aylantiraman!\n\n"
        "Bir nechta rasm yuborsangiz, ular bitta PDF ga birlashtiriladi.\n\n"
        "Buyruqlar:\n"
        "/start - Botni qayta ishga tushirish\n"
        "/help - Yordam\n"
        "/convert - Faylni PDF ga aylantirish"
    )
    await update.message.reply_text(welcome)


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    help_text = (
        "PDF Converter Bot - Yordam\n\n"
        "Qanday ishlaydi:\n"
        "1. Menga rasm yoki matn fayli yuboring\n"
        "2. Bot avtomatik ravishda PDF yaratadi\n"
        "3. Tayyor PDF faylni sizga qaytaradi\n\n"
        "Qo'llab-quvvatlanadigan formatlar:\n"
        "  Rasmlar: JPG, JPEG, PNG, BMP, GIF, WebP, TIFF\n"
        "  Hujjatlar: TXT\n\n"
        "Bir nechta rasm yuborishingiz mumkin — ular bitta PDF bo'ladi.\n"
        "Album sifatida yuboring."
    )
    await update.message.reply_text(help_text)


async def convert_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Iltimos, PDF ga aylantirish uchun rasm yoki matn faylini yuboring."
    )
    return WAITING_FILE


async def _process_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message.photo:
        photo = update.message.photo[-1]
        file = await photo.get_file()
        ext = ".jpg"
    elif update.message.document:
        document = update.message.document
        file = await document.get_file()
        ext = os.path.splitext(document.file_name or "file")[1]
    else:
        await update.message.reply_text("Fayl topilmadi.")
        return

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    input_path = f"temp_{timestamp}{ext}"
    output_path = f"output_{timestamp}.pdf"

    status_msg = await update.message.reply_text("Fayl yuklanmoqda...")

    try:
        await file.download_to_drive(input_path)
        await status_msg.edit_text("PDF yaratilmoqda...")

        if is_image(input_path):
            try:
                image_to_pdf_with_reportlab(input_path, output_path)
            except Exception:
                images_to_pdf([input_path], output_path)
        elif ext == ".txt":
            with open(input_path, "r", encoding="utf-8") as f:
                text_content = f.read()
            text_to_pdf(text_content, output_path)
        else:
            await status_msg.edit_text(
                f"Bu format hozircha qo'llab-quvvatlanmaydi: {ext}"
            )
            return

        with open(output_path, "rb") as pdf_file:
            await update.message.reply_document(
                document=pdf_file,
                filename=f"converted_{timestamp}.pdf",
                caption="PDF tayyor!",
            )
        await status_msg.delete()

    except Exception as e:
        logger.error(f"Xatolik: {e}")
        await status_msg.edit_text(
            f"Xatolik yuz berdi: {str(e)[:200]}\n"
            "Iltimos, boshqa fayl yuboring."
        )
    finally:
        for f in [input_path, output_path]:
            if os.path.exists(f):
                os.remove(f)


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await _process_file(update, context)
    return ConversationHandler.END


async def handle_photo_standalone(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await _process_file(update, context)


async def handle_album_photos(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.message.photo:
        return

    if "album_photos" not in context.user_data:
        context.user_data["album_photos"] = []
        context.user_data["album_processed"] = False

    photo = update.message.photo[-1]
    file = await photo.get_file()

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    input_path = f"album_{timestamp}_{photo.file_id[-8:]}.jpg"
    await file.download_to_drive(input_path)
    context.user_data["album_photos"].append(input_path)

    if not context.user_data.get("album_timer_set"):
        context.user_data["album_timer_set"] = True
        context.job_queue.run_once(
            process_album, 2.0, context=update.effective_chat.id, name="album_timer"
        )


async def process_album(context: ContextTypes.DEFAULT_TYPE):
    chat_id = context.job.context
    photos = context.user_data.get("album_photos", [])

    if not photos or context.user_data.get("album_processed"):
        return

    context.user_data["album_processed"] = True

    status_msg = await context.bot.send_message(
        chat_id, f"{len(photos)} ta rasm PDF ga aylantirilmoqda..."
    )

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = f"album_{timestamp}.pdf"

    try:
        images_to_pdf(photos, output_path)

        with open(output_path, "rb") as pdf_file:
            await context.bot.send_document(
                chat_id,
                pdf_file,
                filename=f"album_{timestamp}.pdf",
                caption=f"{len(photos)} tadan tashkil topgan PDF tayyor!",
            )
        await status_msg.delete()
    except Exception as e:
        logger.error(f"Album xatolik: {e}")
        await status_msg.edit_text("PDF yaratishda xatolik yuz berdi.")
    finally:
        for p in photos:
            if os.path.exists(p):
                os.remove(p)
        if os.path.exists(output_path):
            os.remove(output_path)

    context.user_data["album_photos"] = []
    context.user_data["album_processed"] = False
    context.user_data["album_timer_set"] = False


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    if not document:
        return

    filename = document.file_name or "file"
    ext = os.path.splitext(filename.lower())[1]

    if ext == ".txt":
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        input_path = f"doc_{timestamp}{ext}"
        output_path = f"doc_{timestamp}.pdf"

        status_msg = await update.message.reply_text("Fayl yuklanmoqda...")

        try:
            file = await document.get_file()
            await file.download_to_drive(input_path)
            await status_msg.edit_text("PDF yaratilmoqda...")

            with open(input_path, "r", encoding="utf-8") as f:
                text_content = f.read()

            text_to_pdf(text_content, output_path)

            with open(output_path, "rb") as pdf_file:
                await update.message.reply_document(
                    document=pdf_file,
                    filename=f"{os.path.splitext(filename)[0]}.pdf",
                    caption="PDF tayyor!",
                )
            await status_msg.delete()
        except Exception as e:
            logger.error(f"Document xatolik: {e}")
            await status_msg.edit_text(f"Xatolik: {str(e)[:200]}")
        finally:
            for f in [input_path, output_path]:
                if os.path.exists(f):
                    os.remove(f)
    elif is_image(filename):
        await _process_file(update, context)
    else:
        await update.message.reply_text(
            f"Bu format hozircha qo'llab-quvvatlanmaydi: {ext}\n"
            "Qo'llab-quvvatlanadigan formatlar: rasm (jpg, png, bmp, gif, webp) va .txt"
        )


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Iltimos, PDF ga aylantirish uchun rasm yoki fayl yuboring.\n"
        "Yordam uchun /help buyrug'ini ishlating."
    )


def main():
    if TOKEN == "YOUR_BOT_TOKEN_HERE":
        print("=" * 50)
        print("BOT_TOKEN o'zgartirilishi kerak!")
        print("Quidagi usullardan birini ishlating:")
        print("1. Environment variable: set BOT_TOKEN=your_token")
        print("2. bot.py faylida TOKEN o'rniga tokeningizni yozing")
        print("=" * 50)
        return

    app = Application.builder().token(TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("convert", convert_command)],
        states={
            WAITING_FILE: [
                MessageHandler(filters.PHOTO, handle_photo),
                MessageHandler(filters.Document.ALL, handle_document),
            ],
        },
        fallbacks=[CommandHandler("cancel", lambda u, c: ConversationHandler.END)],
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(conv_handler)
    app.add_handler(
        MessageHandler(filters.PHOTO & ~filters.COMMAND, handle_photo_standalone)
    )
    app.add_handler(
        MessageHandler(filters.Document.ALL & ~filters.COMMAND, handle_document)
    )
    app.add_handler(
        MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message)
    )

    print("Bot ishga tushdi...")
    print("To'xtatish uchun Ctrl+C bosing")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
