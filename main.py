import os
import io
import json
import logging
import requests
from pypdf import PdfReader
from docx import Document
from pptx import Presentation
import openpyxl
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, filters, ContextTypes

# Logging Setup
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)

# Environment Variables
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
ADMIN_ID = int(os.environ.get("ADMIN_ID", "0"))

# File Text Memory Storage
file_text_cache = {}

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_name = update.effective_user.first_name

    if user_id == ADMIN_ID:
        await update.message.reply_text(
            f"👋 **မင်္ဂလာပါ Admin {user_name}!**\n\n"
            f"⚙️ **Admin Control Panel:**\n"
            f"သင်ခန်းစာ ဖိုင်များ (PDF, Word, PPT, Excel) ကို ပို့ပေးပါ။\n"
            f"လှပဆွဲဆောင်မှုရှိသော Interactive Quiz မေးခွန်းများကို အလိုအလျောက် ထုတ်ပေးပါမည်။ 🎯"
        )
    else:
        await update.message.reply_text(
            f"👋 **မင်္ဂလာပါ {user_name}!**\n\n"
            f"🎯 **Daily Quiz Bot မှ ကြိုဆိုပါတယ်။**\n"
            f"ဒီ Bot မှာ တက်လာတဲ့ Quiz မေးခွန်းလေးတွေကို ဖြေဆိုပြီး မိမိ၏ အသိပညာကို စိန်ခေါ်စမ်းသပ်နိုင်ပါတယ်! ✨"
        )

# Document Extraction
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

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id

    if user_id != ADMIN_ID:
        await update.message.reply_text("⛔ ဒီ Feature ကို Admin သာ အသုံးပြုနိုင်ပါသည်။")
        return

    if not GEMINI_API_KEY:
        await update.message.reply_text("❌ GEMINI_API_KEY မရှိသေးပါ။")
        return

    media = update.message.document
    if not media:
        await update.message.reply_text("⚠️ ကျေးဇူးပြု၍ ဖိုင်ပို့ပေးပါ။")
        return

    file_name = getattr(media, 'file_name', 'file').lower()
    status_msg = await update.message.reply_text("⏳ **ဖိုင်ကို ဖတ်ရှုစစ်ဆေးနေပါသည်...** 📄")

    try:
        file = await context.bot.get_file(media.file_id)
        file_bytes = await file.download_as_bytearray()
        extracted_text = ""

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

        file_text_cache[user_id] = extracted_text[:5000]

        # Attractive Menu Options
        keyboard = [
            [
                InlineKeyboardButton("🟢 မေးခွန်း ၃ ခု (Quick Quiz)", callback_data="gen_3_easy"),
                InlineKeyboardButton("🟡 မေးခွန်း ၅ ခု (Standard)", callback_data="gen_5_medium"),
            ],
            [
                InlineKeyboardButton("🔴 မေးခွန်း ၁၀ ခု (Challenge)", callback_data="gen_10_hard"),
            ]
        ]
        reply_markup = InlineKeyboardMarkup(keyboard)

        await status_msg.edit_text(
            "🎉 **ဖိုင်ကို အောင်မြင်စွာ ဖတ်ရှုပြီးပါပြီ!**\n\n"
            "⚙️ **မည်သည့် Quiz အမျိုးအစား ထုတ်ယူလိုသနည်း?**",
            reply_markup=reply_markup
        )

    except Exception as e:
        await status_msg.edit_text(f"❌ Error ဖြစ်ပွားပါသည်: {str(e)}")

async def handle_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    user_id = query.from_user.id
    data = query.data

    if user_id != ADMIN_ID:
        await query.message.reply_text("⛔ Admin သာ အသုံးပြုခွင့်ရှိသည်။")
        return

    if data.startswith("gen_"):
        parts = data.split("_")
        count = int(parts[1])
        difficulty = parts[2]

        extracted_text = file_text_cache.get(user_id, "")
        if not extracted_text:
            await query.message.reply_text("❌ ဖိုင်အချက်အလက် မရှိတော့ပါ။ ဖိုင်ပြန်ပို့ပေးပါ။")
            return

        await query.edit_message_text(f"🤖 **Gemini AI မှ Attractive Quiz ({count}) ခု ဖန်တီးပေးနေပါပြီ...** ✨")

        prompt = f"""
        အောက်ပါ စာသားများကို အခြေခံ၍ စိတ်ဝင်စားဖွယ်ရာ Multiple Choice Quiz မေးခွန်း ({count}) ခု ထုတ်ပေးပါ။
        ခက်ခဲမှုအဆင့်: {difficulty}
        
        အဖြေများကို JSON Format အတိအကျဖြင့်သာ ပြန်ပေးပါ။
        JSON Format ပုံစံ:
        [
          {{
            "question": "မေးခွန်းစာသား",
            "options": ["A ရွေးချယ်စရာ", "B ရွေးချယ်စရာ", "C ရွေးချယ်စရာ", "D ရွေးချယ်စရာ"],
            "correct_index": 0,
            "explanation": "💡 အဖြေမှန်၏ ရှင်းလင်းချက်"
          }}
        ]
        correct_index သည် မှန်ကန်သော အဖြေ၏ Index (0 မှ 3 အထိ) ဖြစ်ရမည်။
        """

        model_name = "gemini-3.6-flash"
        api_url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={GEMINI_API_KEY}"

        payload = {
            "contents": [{"parts": [{"text": f"{prompt}\n\nစာသားများ:\n{extracted_text}"}]}]
        }

        try:
            res = requests.post(api_url, headers={"Content-Type": "application/json"}, json=payload, timeout=60)
            res_json = res.json()

            if res.status_code == 200:
                raw_text = res_json['candidates'][0]['content']['parts'][0]['text']
                clean_json = raw_text.replace("```json", "").replace("```", "").strip()
                quiz_data = json.loads(clean_json)

                context.user_data['latest_quiz'] = quiz_data

                result_text = f"✨ **ဆွဲဆောင်မှုရှိသော Quiz ({len(quiz_data)}) ခု အဆင်သင့်ဖြစ်ပါပြီ!**\n\n"
                for idx, q in enumerate(quiz_data, 1):
                    result_text += f"❓ **{idx}. {q['question']}**\n"
                    for opt_idx, opt in enumerate(q['options']):
                        result_text += f"   {'🔹🔸🔺▪️'[opt_idx]} {opt}\n"
                    result_text += f"✅ **အဖြေမှန်:** {'ABCD'[q['correct_index']]}\n\n"

                publish_keyboard = InlineKeyboardMarkup([
                    [InlineKeyboardButton("🚀 Interactive Quiz Poll အဖြစ် ပို့မည် 🚀", callback_data="publish_poll")]
                ])

                await query.message.reply_text(result_text, reply_markup=publish_keyboard)

            else:
                err_msg = res_json.get('error', {}).get('message', 'Unknown Error')
                await query.message.reply_text(f"❌ Gemini API Error: {err_msg}")

        except Exception as e:
            await query.message.reply_text(f"❌ Error ဖြစ်ပွားပါသည်: {str(e)}")

    elif data == "publish_poll":
        quiz_data = context.user_data.get('latest_quiz', [])
        if not quiz_data:
            await query.message.reply_text("❌ ပို့ရန် မေးခွန်း မရှိပါ။")
            return

        chat_id = query.message.chat_id
        for q in quiz_data:
            # Native Interactive Telegram Quiz Poll Sending
            await context.bot.send_poll(
                chat_id=chat_id,
                question=f"❓ {q['question']}"[:300],
                options=[opt[:100] for opt in q['options']],
                type="quiz",
                correct_option_id=q['correct_index'],
                explanation=q.get('explanation', '')[:200],
                is_anonymous=False
            )
        await query.message.reply_text("🎉 **Interactive Quiz Poll များကို အောင်မြင်စွာ တင်ပြီးပါပြီ!**")

def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(CallbackQueryHandler(handle_callback))

    print("Bot is running...")
    app.run_polling()

if __name__ == '__main__':
    main()
