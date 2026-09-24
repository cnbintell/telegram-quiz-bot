import os
import io
import logging
import requests
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

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_name = update.effective_user.first_name
    await update.message.reply_text(
        f"မင်္ဂလာပါ {user_name}!\n\n"
        f"🤖 Multimodal Quiz Bot မှ ကြိုဆိုပါတယ်။\n"
        f"📌 Admin အနေဖြင့် မေးခွန်းများ ထုတ်ယူရန် ဖိုင် သို့မဟုတ် ပုံများကို ပို့ပေးနိုင်ပါသည်။"
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

def get_available_model_name(api_key):
    """API Key အောက်မှာ အမှန်တကယ် သုံးလို့ရတဲ့ Gemini Model နာမည်ကို အလိုအလျောက် ရှာဖွေပေးခြင်း"""
    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models?key={api_key}"
        res = requests.get(url, timeout=10)
        if res.status_code == 200:
            models_data = res.json().get('models', [])
            for m in models_data:
                name = m.get('name', '')
                supported_methods = m.get('supportedGenerationMethods', [])
                if "generateContent" in supported_methods and "gemini" in name:
                    return name # e.g. "models/gemini-1.5-flash"
    except Exception:
        pass
    return "models/gemini-1.5-flash"

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
        # အသုံးပြုနိုင်သော Model ကို အလိုအလျောက် ရွေးချယ်ခြင်း
        model_path = get_available_model_name(GEMINI_API_KEY)
        api_url = f"https://generativelanguage.googleapis.com/v1beta/{model_path}:generateContent?key={GEMINI_API_KEY}"

        prompt = """
        ပေးပို့ထားသော စာသား/အချက်အလက်များကို အခြေခံ၍ Multiple Choice Quiz မေးခွန်း (၅) ခု ထုတ်ပေးပါ။
        မေးခွန်းတစ်ခုစီအတွက် ရွေးချယ်စရာ (A, B, C, D) နှင့် မှန်ကန်သော အဖြေကို ရှင်းလင်းချက်နှင့်တကွ မြန်မာလို ဖော်ပြပေးပါ။
        """

        extracted_text = ""

        if update.message.photo:
            await status_msg.edit_text("⚠️ ဓာတ်ပုံများအတွက် စာသားပြောင်းလဲခြင်းကို လုပ်ဆောင်နေပါသည်...")
            return

        elif update.message.document:
            media = update.message.document
            file_name = getattr(media, 'file_name', 'file').lower()
            file = await context.bot.get_file(media.file_id)
            file_bytes = await file.download_as_bytearray()

            if file_name.endswith('.docx'):
                extracted_text = extract_text_from_docx(file_bytes)
            elif file_name.endswith('.pptx'):
                extracted_text = extract_text_from_pptx(file_bytes)
            elif file_name.endswith('.xlsx'):
                extracted_text = extract_text_from_excel(file_bytes)
            elif file_name.endswith('.pdf'):
                pdf_reader = PdfReader(io.BytesIO(file_bytes))
                extracted_text = "\n".join([page.extract_text() for page in pdf_reader.pages if page.extract_text()])

        if not extracted_text.strip():
            await status_msg.edit_text("❌ ဖိုင်ထဲမှ စာသားများကို ဖတ်ရှု၍ မရရှိပါ။")
            return

        payload = {
            "contents": [
                {
                    "parts": [{"text": f"{prompt}\n\nစာသားများ:\n{extracted_text[:4000]}"}]
                }
            ]
        }

        headers = {"Content-Type": "application/json"}
        response = requests.post(api_url, headers=headers, json=payload, timeout=60)
        res_json = response.json()

        if response.status_code == 200:
            generated_text = res_json['candidates'][0]['content']['parts'][0]['text']
            await status_msg.edit_text(f"✅ **ထုတ်ယူရရှိသော မေးခွန်းများ:**\n\n{generated_text}")
        else:
            err_msg = res_json.get('error', {}).get('message', 'Unknown Error')
            await status_msg.edit_text(f"❌ Gemini API Error: {err_msg}")

    except Exception as e:
        await status_msg.edit_text(f"❌ Error ဖြစ်ပွားပါသည်: {str(e)}")

def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.PHOTO | filters.Document.ALL, handle_media_and_docs))
    
    print("Bot is running...")
    app.run_polling()

if __name__ == '__main__':
    main()
