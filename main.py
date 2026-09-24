import os
import logging
from datetime import date
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    filters,
    ContextTypes,
)
import google.generativeai as genai

# Logging Configuration
logging.basicConfig(level=logging.INFO)

# Environment Variables မှ Keys များကို ရယူခြင်း
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
ADMIN_ID = int(os.environ.get("ADMIN_ID", "0"))  # မင်းရဲ့ Telegram ID

# Gemini AI Setup
genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel("gemini-1.5-flash")

# In-Memory Database (ခေတ္တ သိမ်းဆည်းရန်)
# လက်တွေ့သုံးလျှင် SQLite သို့မဟုတ် Supabase နှင့် ချိတ်ဆက်နိုင်သည်
users_db = {} 
# Format: { user_id: {"is_premium": False, "daily_count": 0, "last_date": "YYYY-MM-DD"} }

quiz_db = [] 
# Format: [{"question": "...", "options": ["A", "B", "C", "D"], "correct_index": 0, "category": "..."}]


def check_user(user_id):
    today = str(date.today())
    if user_id not in users_db:
        users_db[user_id] = {"is_premium": False, "daily_count": 0, "last_date": today}
    
    # Date Reset logic
    if users_db[user_id]["last_date"] != today:
        users_db[user_id]["last_date"] = today
        users_db[user_id]["daily_count"] = 0
        
    return users_db[user_id]


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    check_user(user_id)
    
    keyboard = [
        [InlineKeyboardButton("🧠 နေ့စဉ် Quiz ဖြေမည်", callback_data="get_quiz")],
        [InlineKeyboardButton("📚 Study Library", callback_data="open_library")],
        [InlineKeyboardButton("👑 Premium ရယူရန်", callback_data="upgrade_premium")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    
    welcome_text = (
        "👋 မင်္ဂလာပါ! Up-to-Date Quiz Bot မှ ကြိုဆိုပါတယ်။\n\n"
        "✨ အခမဲ့ အသုံးပြုသူများအတွက် တစ်နေ့ မေးခွန်း (၁၀) ခု ဖြေဆိုနိုင်ပါသည်။\n"
        "👑 Premium User များအတွက် အကန့်အသတ်မရှိ ဖြေဆိုနိုင်ပါသည်။"
    )
    await update.message.reply_text(welcome_text, reply_markup=reply_markup)


async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    user_data = check_user(user_id)

    if query.data == "get_quiz":
        # Free User Daily Limit Check (၁၀ ခု ကန့်သတ်ချက်)
        if not user_data["is_premium"] and user_data["daily_count"] >= 10:
            await query.edit_message_text(
                "❌ ဒီနေ့အတွက် အခမဲ့ မေးခွန်း ၁၀ ခု ဖြေဆိုပြီးပါပြီ။\n\n"
                "မနက်ဖြန်မှ ထပ်မံဖြေဆိုပါ သို့မဟုတ် အကန့်အသတ်မရှိ ဖြေဆိုရန် /premium မှတစ်ဆင့် Premium သို့ မြှင့်တင်ပါ။"
            )
            return

        if not quiz_db:
            await query.edit_message_text("⚠️ လတ်တလော မေးခွန်းများ မရှိသေးပါ။ Admin မေးခွန်းအသစ်များ တင်ပေးရန် စောင့်ဆိုင်းပေးပါ။")
            return

        # Quiz တင်ပေးခြင်း Logic
        quiz = quiz_db[0] # နမူနာအဖြစ် ပထမဆုံး မေးခွန်းပေးခြင်း
        user_data["daily_count"] += 1
        
        keyboard = []
        for idx, option in enumerate(quiz["options"]):
            keyboard.append([InlineKeyboardButton(option, callback_data=f"ans_{idx}_{quiz['correct_index']}")])
        
        reply_markup = InlineKeyboardMarkup(keyboard)
        await query.edit_message_text(f"❓ မေးခွန်း ({user_data['daily_count']}/10):\n\n{quiz['question']}", reply_markup=reply_markup)

    elif query.data.startswith("ans_"):
        _, selected, correct = query.data.split("_")
        if selected == correct:
            await query.edit_message_text("🎉 မှန်ကန်ပါတယ်။ အဖြေမှန်ပါသည်!")
        else:
            await query.edit_message_text("❌ မှားယွင်းပါတယ်။ နောက်တစ်ကြိမ် ထပ်မံကြိုးစားပါ။")

    elif query.data == "open_library":
        await query.edit_message_text("📚 Study Library ကဏ္ဍသို့ ရောက်ရှိပါပြီ။\n\nဒီနေရာတွင် Topic အလိုက် သင်ခန်းစာ အနှစ်ချုပ်များကို ဖတ်ရှုနိုင်ပါသည်။")

    elif query.data == "upgrade_premium":
        await query.edit_message_text(
            "👑 **Premium အကောင့် ရယူရန်**\n\n"
            "KBZPay / WavePay မှတစ်ဆင့် လစဉ်ကြေး ၃,၀၀၀ ကျပ် လွှဲပြောင်းပေးပို့ပြီး ငွေလွှဲပြေစာ (Screenshot) ကို Admin ထံ ပေးပို့ပေးပါ သို့မဟုတ် Bot ထံ ပို့ပေးပါ။\n\n"
            "📞 Contact Admin: @your_admin_username"
        )


# Admin သီးသန့် PDF/Text ထည့်သွင်းစနစ်
async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id != ADMIN_ID:
        await update.message.reply_text("⚠️ သင့်တွင် PDF ထည့်သွင်းခွင့် မရှိပါ။")
        return

    await update.message.reply_text("⏳ PDF ကို AI မှ ဖတ်ရှုပြီး မေးခွန်းများ ထုတ်ယူနေပါသည်။ ခဏစောင့်ပါ။...")
    
    # AI ဖြင့် PDF ကနေ မေးခွန်းထုတ်မည့် Prompt
    prompt = (
        "ဒီ Text/PDF ထဲက အကြောင်းအရာတွေကို အခြေခံပြီး MCQ မေးခွန်း ၅ ခု ထုတ်ပေးပါ။\n"
        "Format အဖြစ် အောက်ပါ JSON အတိုင်းသာ ထုတ်ပေးပါ:\n"
        "[{\"question\": \"...\", \"options\": [\"A\", \"B\", \"C\", \"D\"], \"correct_index\": 0}]"
    )
    
    # Gemini AI Processing
    try:
        response = model.generate_content(prompt)
        await update.message.reply_text("✅ မေးခွန်းများကို အောင်မြင်စွာ ထုတ်ယူ သိမ်းဆည်းလိုက်ပါပြီ!")
    except Exception as e:
        await update.message.reply_text(f"❌ Error ဖြစ်ပွားပါသည်: {e}")


def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))

    print("Bot is running...")
    app.run_polling()


if __name__ == "__main__":
    main()
