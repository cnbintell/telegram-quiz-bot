import os
import io
import logging
import google.generativeai as genai
from pypdf import PdfReader
from docx import Document
from pptx import Presentation
import openpyxl
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

# Logging setup
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)

# Environment Variables
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
ADMIN_ID = int(os.environ.get("ADMIN_ID", "0"))

if GEMINI_API_KEY:
    genai.configure(api_key=GEMINI_API_KEY)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_name = update.effective_user.first_name
    await update.message.reply_text(
        f"မင်္ဂလာပါ {user_name}!\n\n"
        f"🤖 Gemini AI စွမ်းအားသုံး Multimodal Quiz Bot မှ ကြိုဆိုပါတယ်။\n"
        f"📌 Admin အနေဖြင့် PDF, Word, PowerPoint, Excel, Image နှင့် Video ဖိုင်များကို ပို့ပေး၍ မေးခွန်းများ ထုတ်ယူနိုင်ပါသည်။"
    )

def extract_text_from_docx(file_bytes):
    doc = Document(io.BytesIO(file_bytes))
    return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])

def extract_text_from_pptx(file_bytes):
    prs = Presentation(io.BytesIO(file_bytes))
    text = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text.append(shape.text)
    return "\n".join(text)

def extract_text_from_excel(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    text = []
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(values_only=True):
            row_text = [str(cell) for cell in row if cell is not None]
            if row_text:
                text.append(" | ".join(row_text))
    return "\n".join(text)

async def handle_media_and_docs(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if user_id != ADMIN_ID:
        await update.message.reply_text("⛔ ဒီ Feature ကို Admin သာ အသုံးပြုခွင့်ရှိပါတယ်။")
        return

    if not GEMINI_API_KEY:
        await update.message.reply_text("❌ GEMINI_API_KEY မရှိသေးပါ။ GitHub Secrets ထဲတွင် ထည့်သွင်းပေးပါ။")
        return

    status_msg = await update.message.reply_text("⏳ ဖိုင်ကို ဖတ်ရှုပြီး Gemini AI ဖြင့် မေးခွန်းများ ထုတ်ယူနေပါတယ်...")

    try:
        model = genai.GenerativeModel('gemini-1.5-flash')
        prompt = """
        ပေးပို့ထားသော အချက်အလက်များ/ဖိုင်ကို အခြေခံ၍ Multiple Choice Quiz မေးခွန်း (၅) ခု ထုတ်ပေးပါ။
        မေးခွန်းတစ်ခုစီအတွက် ရွေးချယ်စရာ (A, B, C, D) နှင့် မှန်ကန်သော အဖြေကို ရှင်းလင်းချက်နှင့်တကွ မြန်မာလို ဖော်ပြပေးပါ။
        """

        # 1. Image ဖိုင်များ (Photo/Image attachment)
        if update.message.photo:
            photo = update.message.photo[-1]
            file = await context.bot.get_file(photo.file_id)
            file_bytes = await file.download_as_bytearray()
            
            image_part = {"mime_type": "image/jpeg", "data": bytes(file_bytes)}
            response = model.generate_content([prompt, image_part])

        # 2. Document & Video ဖိုင်များ
        elif update.message.document or update.message.video:
            media = update.message.document or update.message.video
            file_name = getattr(media, 'file_name', 'file').lower()
            file = await context.bot.get_file(media.file_id)
            file_bytes = await file.download_as_bytearray()

            # Word (.docx)
            if file_name.endswith('.docx'):
                text = extract_text_from_docx(file_bytes)
                response = model.generate_content([f"{prompt}\n\nစာသားများ:\n{text[:4000]}"])

            # PowerPoint (.pptx)
            elif file_name.endswith('.pptx'):
                text = extract_text_from_pptx(file_bytes)
                response = model.generate_content([f"{prompt}\n\nစာသားများ:\n{text[:4000]}"])

            # Excel (.xlsx)
            elif file_name.endswith('.xlsx'):
                text = extract_text_from_excel(file_bytes)
                response = model.generate_content([f"{prompt}\n\nစာသားများ:\n{text[:4000]}"])

            # PDF (.pdf)
            elif file_name.endswith('.pdf'):
                pdf_reader = PdfReader(io.BytesIO(file_bytes))
                text = "\n".join([page.extract_text() for page in pdf_reader.pages if page.extract_text()])
                response = model.generate_content([f"{prompt}\n\nစာသားများ:\n{text[:4000]}"])

            # Image Documents (.jpg, .png, .webp)
            elif file_name.endswith(('.jpg', '.jpeg', '.png', '.webp')):
                mime_type = "image/png" if file_name.endswith('.png') else "image/jpeg"
                image_part = {"mime_type": mime_type, "data": bytes(file_bytes)}
                response = model.generate_content([prompt, image_part])

            # Video (.mp4, .mov, .avi)
            elif file_name.endswith(('.mp4', '.mov', '.avi')) or update.message.video:
                mime_type = media.mime_type or "video/mp4"
                video_part = {"mime_type": mime_type, "data": bytes(file_bytes)}
                response = model.generate_content([prompt, video_part])

            else:
                await status_msg.edit_text("⚠️ လက်ခံ၍ မရသော File Format ဖြစ်ပါသည်။ (.pdf, .docx, .pptx, .xlsx, .png, .jpg, .mp4 တို့ကိုသာ ပို့ပေးပါ)")
                return

        else:
            await status_msg.edit_text("⚠️ ကျေးဇူးပြု၍ ဖိုင်၊ ပုံ သို့မဟုတ် ဗီဒီယို ပို့ပေးပါ။")
            return

        if response and response.text:
            await status_msg.edit_text(f"✅ **ထုတ်ယူရရှိသော မေးခွန်းများ:**\n\n{response.text}")
        else:
            await status_msg.edit_text("❌ Gemini AI ထံမှ မေးခွန်းများ ထုတ်ယူ၍ မရရှိပါ။")

    except Exception as e:
        await status_msg.edit_text(f"❌ Error ဖြစ်ပွားပါသည်: {str(e)}")

def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.PHOTO | filters.Document.ALL | filters.VIDEO, handle_media_and_docs))
    
    print("Bot is running...")
    app.run_polling()

if __name__ == '__main__':
    main()
