import os
import io
import logging
import requests
from pypdf import PdfReader
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
        f"🤖 ကျွန်တော်ကတော့ Gemini AI စွမ်းအားသုံး Quiz Bot ဖြစ်ပါတယ်။\n"
        f"📌 Admin အနေဖြင့် မေးခွန်းများ ထုတ်ယူရန် PDF ဖိုင်ကို ပို့ပေးနိုင်ပါသည်။"
    )

async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    # Admin စစ်ဆေးခြင်း
    if user_id != ADMIN_ID:
        await update.message.reply_text("⛔ ဒီ Feature ကို Admin သာ အသုံးပြုခွင့်ရှိပါတယ်။")
        return

    if not GEMINI_API_KEY:
        await update.message.reply_text("❌ GEMINI_API_KEY မရှိသေးပါ။ GitHub Secrets ထဲတွင် ထည့်သွင်းပေးပါ။")
        return

    document = update.message.document
    if not document.file_name.lower().endswith('.pdf'):
        await update.message.reply_text("⚠️ ကျေးဇူးပြု၍ PDF ဖိုင်ကိုသာ ပို့ပေးပါ။")
        return

    status_msg = await update.message.reply_text("⏳ PDF ဖိုင်ကို ဖတ်ရှုပြီး Gemini AI ဖြင့် မေးခွန်းများ ထုတ်ယူနေပါတယ်...")

    try:
        # Telegram မှ ဖိုင်ဒေါင်းလုဒ်ဆွဲခြင်း
        file = await context.bot.get_file(document.file_id)
        file_bytes = await file.download_as_bytearray()
        
        # PDF မှ စာသားဖတ်ခြင်း
        pdf_reader = PdfReader(io.BytesIO(file_bytes))
        text_content = ""
        for page in pdf_reader.pages:
            text_content += page.extract_text() or ""

        if not text_content.strip():
            await status_msg.edit_text("❌ PDF ထဲတွင် စာသားများ ဖတ်မရပါ။ Scan ဖတ်ထားသော ပုံရိပ်များ ဖြစ်နိုင်ပါသည်။")
            return

        # Gemini REST API သို့ တိုက်ရိုက် Request ပို့ခြင်း
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={GEMINI_API_KEY}"
        
        prompt = f"""
        အောက်ပါ သင်ခန်းစာ စာသားများကို အခြေခံ၍ Multiple Choice Quiz မေးခွန်း (၅) ခု ထုတ်ပေးပါ။
        မေးခွန်းတစ်ခုစီအတွက် ရွေးချယ်စရာ (A, B, C, D) နှင့် မှန်ကန်သော အဖြေကို ရှင်းလင်းချက်နှင့်တကွ မြန်မာလို ဖော်ပြပေးပါ။

        စာသားများ-
        {text_content[:4000]}
        """

        payload = {
            "contents": [{
                "parts": [{"text": prompt}]
            }]
        }

        response = requests.post(url, json=payload)
        res_json = response.json()

        if response.status_code == 200:
            generated_text = res_json['candidates'][0]['content']['parts'][0]['text']
            await status_msg.edit_text(f"✅ **PDF မှ ထုတ်ယူရရှိသော မေးခွန်းများ:**\n\n{generated_text}")
        else:
            error_msg = res_json.get('error', {}).get('message', 'Unknown Error')
            await status_msg.edit_text(f"❌ Gemini API Error: {error_msg}")

    except Exception as e:
        await status_msg.edit_text(f"❌ Error ဖြစ်ပွားပါသည်: {str(e)}")

def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_pdf))
    
    print("Bot is running...")
    app.run_polling()

if __name__ == '__main__':
    main()
