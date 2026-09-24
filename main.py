import os
import io
import json
import time
import logging
import requests
from pypdf import PdfReader
from docx import Document
from pptx import Presentation
import openpyxl
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler, CallbackQueryHandler,
    ConversationHandler, filters, ContextTypes
)

# Logging Setup
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)

# Environment Variables
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
ADMIN_ID = int(os.environ.get("ADMIN_ID", "0"))

# File Text Memory Storage
file_text_cache = {}

# Conversation States for Custom Quiz Creation
QUESTION, OPTIONS, CORRECT_OPTION, EXPLANATION = range(4)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_name = update.effective_user.first_name

    if user_id == ADMIN_ID:
        await update.message.reply_text(
            f"👋 **မင်္ဂလာပါ Admin {user_name}!**\n\n"
            f"⚙️ **Admin Control Panel:**\n"
            f"၁။ ဖိုင်များ (PDF, Word, PPT, Excel) ပို့ပေးပြီး AI ဖြင့် မေးခွန်းထုတ်နိုင်ပါသည်။\n"
            f"၂။ ကိုယ်တိုင် မေးခွန်းရေးသားရန် /createquiz ကို နှိပ်ပါ။ ✍️"
        )
    else:
        await update.message.reply_text(
            f"👋 **မင်္ဂလာပါ {user_name}!**\n\n"
            f"🎯 **Daily Quiz Bot မှ ကြိုဆိုပါတယ်။**\n"
            f"ဒီ Bot မှာ တက်လာတဲ့ Quiz မေးခွန်းလေးတွေကို ဖြေဆိုပြီး မိမိ၏ အသိပညာကို စိန်ခေါ်စမ်းသပ်နိုင်ပါတယ်! ✨"
        )

# Document Extraction Functions
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

# AI Callback Handler
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

        max_retries = 3
        res_json = None
        
        for attempt in range(max_retries):
            try:
                res = requests.post(api_url, headers={"Content-Type": "application/json"}, json=payload, timeout=60)
                res_json = res.json()
                if res.status_code == 200 or 'error' not in res_json:
                    break
                time.sleep(3)
            except Exception:
                time.sleep(3)

        try:
            if res_json and 'candidates' in res_json:
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
                err_msg = res_json.get('error', {}).get('message', 'Google Gemini Server ခေတ္တခဏ အလုပ်များနေပါသည်။ ခဏစောင့်ပြီး ပြန်လည် စမ်းသပ်ပေးပါ။')
                await query.message.reply_text(f"⚠️ Gemini API Error: {err_msg}")

        except Exception as e:
            await query.message.reply_text(f"❌ Error ဖြစ်ပွားပါသည်: {str(e)}")

    elif data == "publish_poll":
        quiz_data = context.user_data.get('latest_quiz', [])
        if not quiz_data:
            await query.message.reply_text("❌ ပို့ရန် မေးခွန်း မရှိပါ။")
            return

        chat_id = query.message.chat_id
        for q in quiz_data:
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

# Custom Quiz Creation Flow
async def create_quiz_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id != ADMIN_ID:
        await update.message.reply_text("⛔ ဒီ Feature ကို Admin သာ အသုံးပြုနိုင်ပါသည်။")
        return ConversationHandler.END

    await update.message.reply_text("✍️ **ကိုယ်ပိုင် မေးခွန်း ဖန်တီးခြင်း**\n\nကျေးဇူးပြု၍ **မေးခွန်း စာသား** ကို ရေးပို့ပေးပါ (ပယ်ဖျက်ရန် /cancel ကို နှိပ်ပါ) -")
    return QUESTION

async def set_question(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['custom_q'] = update.message.text
    await update.message.reply_text(
        "📝 **ရွေးချယ်စရာ (၄) ခု ရေးပေးပါ**\n\n"
        "စာကြောင်း တစ်ကြောင်းစီ ခွဲ၍ (၄) ကြောင်း ရေးပေးပါ။ ဥပမာ -\n"
        "ရွေးချယ်စရာ A\n"
        "ရွေးချယ်စရာ B\n"
        "ရွေးချယ်စရာ C\n"
        "ရွေးချယ်စရာ D"
    )
    return OPTIONS

async def set_options(update: Update, context: ContextTypes.DEFAULT_TYPE):
    options = [line.strip() for line in update.message.text.split("\n") if line.strip()]
    if len(options) < 2 or len(options) > 10:
        await update.message.reply_text("⚠️ ရွေးချယ်စရာ (၂) ခုမှ (၁၀) ခုအထိသာ ထည့်သွင်းပေးပါ။ ပြန်လည် ပို့ပေးပါ -")
        return OPTIONS

    context.user_data['custom_opts'] = options

    keyboard = []
    row = []
    for idx, opt in enumerate(options):
        row.append(InlineKeyboardButton(f"အဖြေ {idx+1}: {opt[:10]}", callback_data=f"correct_{idx}"))
        if len(row) == 2:
            keyboard.append(row)
            row = []
    if row:
        keyboard.append(row)

    await update.message.reply_text("✅ **မည်သည့် ရွေးချယ်စရာက အဖြေမှန် ဖြစ်သနည်း?**", reply_markup=InlineKeyboardMarkup(keyboard))
    return CORRECT_OPTION

async def set_correct_option(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    correct_idx = int(query.data.split("_")[1])
    context.user_data['custom_correct'] = correct_idx

    await query.edit_message_text("💡 **အဖြေမှန်၏ ရှင်းလင်းချက် (Explanation)** ကို ရေးပေးပါ (မထည့်ချင်ပါက `-` သို့မဟုတ် `Skip` ဟု ရေးပါ) -")
    return EXPLANATION

async def set_explanation(update: Update, context: ContextTypes.DEFAULT_TYPE):
    exp = update.message.text
    if exp.lower() in ['-', 'skip']:
        exp = ""

    question = context.user_data.get('custom_q')
    options = context.user_data.get('custom_opts')
    correct_idx = context.user_data.get('custom_correct')

    # Send Quiz Poll Directly
    await context.bot.send_poll(
        chat_id=update.effective_chat.id,
        question=f"❓ {question}"[:300],
        options=[opt[:100] for opt in options],
        type="quiz",
        correct_option_id=correct_idx,
        explanation=exp[:200],
        is_anonymous=False
    )

    await update.message.reply_text("🎉 **သင်ကိုယ်တိုင် ဖန်တီးထားသော Quiz Poll ကို အောင်မြင်စွာ တင်လိုက်ပါပြီ!**\nနောက်ထပ် ထပ်မံ ဖန်တီးလိုပါက /createquiz ကို နှိပ်ပါ၊")
    return ConversationHandler.END

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("❌ မေးခွန်း ဖန်တီးခြင်းကို ပယ်ဖျက်လိုက်ပါပြီ။")
    return ConversationHandler.END

def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler('createquiz', create_quiz_start)],
        states={
            QUESTION: [MessageHandler(filters.TEXT & ~filters.COMMAND, set_question)],
            OPTIONS: [MessageHandler(filters.TEXT & ~filters.COMMAND, set_options)],
            CORRECT_OPTION: [CallbackQueryHandler(set_correct_option, pattern="^correct_")],
            EXPLANATION: [MessageHandler(filters.TEXT & ~filters.COMMAND, set_explanation)],
        },
        fallbacks=[CommandHandler('cancel', cancel)]
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(conv_handler)
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(CallbackQueryHandler(handle_callback))

    print("Bot is running...")
    app.run_polling()

if __name__ == '__main__':
    main()
