import os
import io
import json
import time
import logging
import random
import requests
from datetime import date
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

# Database & Memory Caches
DB_FILE = "quiz_database.json"

def load_db():
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {
        "premium_users": [ADMIN_ID],
        "user_daily_limits": {}, # {"user_id": {"date": "YYYY-MM-DD", "count": 0}}
        "categories": ["General"],
        "questions": [] # [{"id": 1, "category": "IT", "question": "...", "options": [], "correct_index": 0, "explanation": "..."}]
    }

def save_db(db):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(db, f, ensure_ascii=False, indent=2)

db = load_db()

# Conversation States
WAIT_CATEGORY_NAME, WAIT_PREMIUM_ID, WAIT_FILE_RECV, WAIT_AI_PROMPT = range(4)

# Document Extraction
def extract_text_from_file(file_bytes, file_name):
    extracted_text = ""
    file_name = file_name.lower()
    if file_name.endswith('.docx'):
        doc = Document(io.BytesIO(file_bytes))
        extracted_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    elif file_name.endswith('.pptx'):
        prs = Presentation(io.BytesIO(file_bytes))
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"): extracted_text += shape.text + "\n"
    elif file_name.endswith('.xlsx'):
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        for sheet in wb.worksheets:
            for row in sheet.iter_rows(values_only=True):
                extracted_text += " | ".join([str(c) for c in row if c is not None]) + "\n"
    elif file_name.endswith('.pdf'):
        pdf_reader = PdfReader(io.BytesIO(file_bytes))
        extracted_text = "\n".join([page.extract_text() for page in pdf_reader.pages if page.extract_text()])
    return extracted_text

# Gemini AI Caller
def call_gemini_ai(prompt_text):
    model_name = "gemini-3.6-flash"
    api_url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={GEMINI_API_KEY}"
    payload = {"contents": [{"parts": [{"text": prompt_text}]}]}
    
    for _ in range(3):
        try:
            res = requests.post(api_url, headers={"Content-Type": "application/json"}, json=payload, timeout=60)
            res_json = res.json()
            if res.status_code == 200 and 'candidates' in res_json:
                raw = res_json['candidates'][0]['content']['parts'][0]['text']
                clean = raw.replace("```json", "").replace("```", "").strip()
                return json.loads(clean)
            time.sleep(2)
        except Exception:
            time.sleep(2)
    return None

# --- UI Menus ---
def get_main_menu_keyboard(user_id):
    keyboard = [
        [InlineKeyboardButton("🎯 Quiz ဖြေဆိုမည်", callback_data="user_start_quiz")],
        [InlineKeyboardButton("⭐ Premium ရယူရန် / အဆင့်စစ်ရန်", callback_data="user_check_status")]
    ]
    if user_id == ADMIN_ID:
        keyboard.append([InlineKeyboardButton("👑 Admin Panel သို့ဝင်ရန်", callback_data="admin_panel")])
    return InlineKeyboardMarkup(keyboard)

def get_admin_panel_keyboard():
    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton("📁 Category ဖန်တီးမည်", callback_data="admin_add_cat"),
            InlineKeyboardButton("💎 Premium အတည်ပြုမည်", callback_data="admin_add_premium")
        ],
        [
            InlineKeyboardButton("📄 File မှ AI မေးခွန်းထုတ်မည်", callback_data="admin_ai_file"),
            InlineKeyboardButton("🤖 Gemini Chat Bot Mode", callback_data="admin_ai_chat")
        ],
        [InlineKeyboardButton("📊 စာရင်းဇယားနှင့် မေးခွန်းများကြည့်မည်", callback_data="admin_stats")],
        [InlineKeyboardButton("🔙 Main Menu သို့ပြန်သွားမည်", callback_data="go_main_menu")]
    ])

# --- Commands & Handlers ---
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_name = update.effective_user.first_name
    is_prem = user_id in db["premium_users"]
    status_tag = "💎 Premium Member" if is_prem else "🆓 Free Member"

    msg = f"👋 **မင်္ဂလာပါ {user_name}!** ({status_tag})\n\n" \
          f"✨ **Quiz Platform မှ ကြိုဆိုပါတယ်!**\n" \
          f"• Free User: တစ်နေ့လျှင် မေးခွန်း (၁၀) ပုဒ် ဖြေဆိုနိုင်ပါသည်။\n" \
          f"• Premium User: မေးခွန်းများကို အကန့်အသတ်မရှိ စိတ်ကြိုက်ဖြေဆိုနိုင်ပါသည်။"
    
    await update.message.reply_text(msg, reply_markup=get_main_menu_keyboard(user_id), parse_mode="Markdown")

# Main Callback Handler
async def handle_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    data = query.data

    if data == "go_main_menu":
        await query.edit_message_text("📌 **Main Menu**", reply_markup=get_main_menu_keyboard(user_id))

    elif data == "admin_panel" and user_id == ADMIN_ID:
        await query.edit_message_text("👑 **Admin Management Dashboard**\n\nအောက်ပါ Button များမှ စိတ်ကြိုက် စီမံခန့်ခွဲနိုင်ပါသည်:", reply_markup=get_admin_panel_keyboard())

    elif data == "user_check_status":
        is_prem = user_id in db["premium_users"]
        today = str(date.today())
        user_lim = db["user_daily_limits"].get(str(user_id), {"date": today, "count": 0})
        used = user_lim["count"] if user_lim["date"] == today else 0

        status_text = f"👤 **သင့်အကောင့် အခြေအနေ**\n\n" \
                      f"🆔 ID: `{user_id}`\n" \
                      f"အမျိုးအစား: {'💎 **Premium (Unlimited)**' if is_prem else '🆓 **Free User**'}\n"
        if not is_prem:
            status_text += f"ယနေ့ဖြေဆိုပြီး: **{used} / 10** ပုဒ်\n\n" \
                           f"✨ Unlimited ဖြေဆိုလိုပါက Admin ထံ ဆက်သွယ်၍ Premium အဆင့်မြှင့်နိုင်ပါသည်။"
        
        await query.edit_message_text(status_text, reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("🔙 ပြန်သွားမည်", callback_data="go_main_menu")]]), parse_mode="Markdown")

    elif data == "admin_stats" and user_id == ADMIN_ID:
        tot_q = len(db["questions"])
        tot_cat = len(db["categories"])
        tot_prem = len(db["premium_users"])
        stats_msg = f"📊 **System Statistics**\n\n" \
                    f"• စုစုပေါင်း မေးခွန်း: **{tot_q}** ပုဒ်\n" \
                    f"• Category ပမာဏ: **{tot_cat}** ခု ({', '.join(db['categories'])})\n" \
                    f"• Premium User: **{tot_prem}** ယောက်"
        await query.edit_message_text(stats_msg, reply_markup=get_admin_panel_keyboard(), parse_mode="Markdown")

    # User Start Quiz Logic
    elif data == "user_start_quiz":
        is_prem = user_id in db["premium_users"]
        today = str(date.today())
        user_lim = db["user_daily_limits"].get(str(user_id), {"date": today, "count": 0})
        
        if user_lim["date"] != today:
            user_lim = {"date": today, "count": 0}

        if not is_prem and user_lim["count"] >= 10:
            await query.edit_message_text(
                "❌ **ယနေ့အတွက် Limit (၁၀) ပုဒ် ပြည့်သွားပါပြီ!**\n\nမနက်ဖြန်မှ ထပ်မံဖြေဆိုပါ သို့မဟုတ် Premium အဆင့်မြှင့်တင်ပါ။",
                reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("🔙 ပြန်သွားမည်", callback_data="go_main_menu")]])
            )
            return

        # Choose Category
        cat_buttons = []
        for cat in db["categories"]:
            cat_buttons.append([InlineKeyboardButton(f"📁 {cat}", callback_data=f"play_cat_{cat}")])
        
        await query.edit_message_text("📂 **ဖြေဆိုလိုသည့် Category ကို ရွေးချယ်ပါ -**", reply_markup=InlineKeyboardMarkup(cat_buttons))

    elif data.startswith("play_cat_"):
        cat_name = data.replace("play_cat_", "")
        q_pool = [q for q in db["questions"] if q["category"] == cat_name]
        
        if not q_pool:
            await query.edit_message_text(f"⚠️ **{cat_name}** Category ထဲတွင် မေးခွန်းများ မရှိသေးပါ။", reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("🔙 ပြန်သွားမည်", callback_data="go_main_menu")]]))
            return

        # Prepare Quiz Session (10 Questions Random)
        selected_qs = random.sample(q_pool, min(10, len(q_pool)))
        context.user_data["quiz_session"] = {
            "category": cat_name,
            "questions": selected_qs,
            "current_index": 0,
            "score": 0,
            "start_time": time.time()
        }
        await send_quiz_question(query, context)

# --- Quiz Engine ---
async def send_quiz_question(query, context):
    session = context.user_data["quiz_session"]
    idx = session["current_index"]
    qs = session["questions"]

    if idx >= len(qs):
        # Quiz Complete -> Show Result Card
        elapsed = int(time.time() - session["start_time"])
        score = session["score"]
        total = len(qs)
        percentage = (score / total) * 100

        # Update Daily Limit for Free User
        user_id = query.from_user.id
        if user_id not in db["premium_users"]:
            today = str(date.today())
            ulim = db["user_daily_limits"].get(str(user_id), {"date": today, "count": 0})
            if ulim["date"] != today: ulim = {"date": today, "count": 0}
            ulim["count"] += total
            db["user_daily_limits"][str(user_id)] = ulim
            save_db(db)

        # Grade Template
        badge = "🏆 Excellence!" if percentage >= 80 else "👍 Good Job!" if percentage >= 50 else "န အားထုတ်ပါဦး!"
        res_text = f"🎯 **QUIZ RESULT CARD** 🎯\n" \
                   f"━━━━━━━━━━━━━━━━━━\n" \
                   f"📁 **Category:** {session['category']}\n" \
                   f"✨ **ရမှတ်:** {score} / {total}\n" \
                   f"📊 **ရာခိုင်နှုန်း:** {percentage:.1f}%\n" \
                   f"⏱️ **ကြာချိန်:** {elapsed} စက္ကန့်\n" \
                   f"ဆုတံဆိပ်: **{badge}**\n" \
                   f"━━━━━━━━━━━━━━━━━━"

        await query.edit_message_text(res_text, reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("🏠 Main Menu သို့ပြန်သွားမည်", callback_data="go_main_menu")]]), parse_mode="Markdown")
        return

    q = qs[idx]
    opts_btn = []
    for o_idx, opt in enumerate(q["options"]):
        opts_btn.append([InlineKeyboardButton(opt, callback_data=f"ans_{o_idx}")])

    await query.edit_message_text(
        f"❓ **မေးခွန်း ({idx + 1}/{len(qs)}):**\n\n{q['question']}",
        reply_markup=InlineKeyboardMarkup(opts_btn),
        parse_mode="Markdown"
    )

async def handle_quiz_answer(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    session = context.user_data.get("quiz_session")
    if not session: return

    ans_idx = int(query.data.split("_")[1])
    idx = session["current_index"]
    q = session["questions"][idx]

    if ans_idx == q["correct_index"]:
        session["score"] += 1
        await query.message.reply_text(f"✅ **မှန်ပါတယ်!**\n💡 {q.get('explanation', '')}")
    else:
        correct_opt = q["options"][q["correct_index"]]
        await query.message.reply_text(f"❌ **မှားယွင်းပါသည်။**\n✓ အဖြေမှန်: {correct_opt}\n💡 {q.get('explanation', '')}")

    session["current_index"] += 1
    await send_quiz_question(query, context)

# --- Admin Actions Flow ---
async def admin_add_cat_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await query.edit_message_text("✍️ **ဖန်တီးလိုသော Category နာမည်ကို ရေးပို့ပေးပါ -**")
    return WAIT_CATEGORY_NAME

async def admin_save_cat(update: Update, context: ContextTypes.DEFAULT_TYPE):
    cat_name = update.message.text.strip()
    if cat_name not in db["categories"]:
        db["categories"].append(cat_name)
        save_db(db)
        await update.message.reply_text(f"✅ Category **'{cat_name}'** ကို အောင်မြင်စွာ ထည့်သွင်းပြီးပါပြီ။", reply_markup=get_admin_panel_keyboard())
    else:
        await update.message.reply_text("⚠️ ဒီ Category နာမည် ရှိပြီးသား ဖြစ်ပါသည်။", reply_markup=get_admin_panel_keyboard())
    return ConversationHandler.END

async def admin_add_prem_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await query.edit_message_text("💎 **Premium ပေးလိုသော User ၏ Telegram ID ကို ရိုက်ပို့ပေးပါ -**")
    return WAIT_PREMIUM_ID

async def admin_save_prem(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        uid = int(update.message.text.strip())
        if uid not in db["premium_users"]:
            db["premium_users"].append(uid)
            save_db(db)
            await update.message.reply_text(f"🎉 User ID `{uid}` အား **Premium User** အဖြစ် အတည်ပြုပေးလိုက်ပါပြီ။", reply_markup=get_admin_panel_keyboard())
        else:
            await update.message.reply_text("⚠️ ဒီ User သည် Premium ဖြစ်ပြီးသား ဖြစ်ပါသည်။", reply_markup=get_admin_panel_keyboard())
    except ValueError:
        await update.message.reply_text("❌ ID သည် ကိန်းဂဏန်း သို့မဟုတ် တိကျသော စာသား ဖြစ်ရပါမည်။", reply_markup=get_admin_panel_keyboard())
    return ConversationHandler.END

# Admin File / AI Chat Prompt Flow
async def admin_ai_chat_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await query.edit_message_text("🤖 **Gemini AI Chat Mode**\n\nမည်သည့် မေးခွန်းမျိုး ထုတ်ယူချင်လဲဆိုတာ စာဖြင့် ခိုင်းလိုက်ပါ (ဥပမာ- 'Grade 10 English သဒ္ဒါ မေးခွန်း ၅ ခု ထုတ်ပေးပါ') -")
    return WAIT_AI_PROMPT

async def admin_process_ai_prompt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_prompt = update.message.text
    status_msg = await update.message.reply_text("⏳ **Gemini AI မှ မေးခွန်းများ ဖန်တီးပေးနေပါသည်...**")

    prompt = f"""
    {user_prompt}
    
    အဖြေများကို JSON Format အတိအကျဖြင့်သာ ပြန်ပေးပါ။
    [
      {{
        "question": "မေးခွန်းစာသား",
        "options": ["A ရွေးချယ်စရာ", "B ရွေးချယ်စရာ", "C ရွေးချယ်စရာ", "D ရွေးချယ်စရာ"],
        "correct_index": 0,
        "explanation": "ရှင်းလင်းချက်"
      }}
    ]
    """
    quiz_data = call_gemini_ai(prompt)
    if quiz_data:
        default_cat = db["categories"][0]
        for q in quiz_data:
            q["id"] = len(db["questions"]) + 1
            q["category"] = default_cat
            db["questions"].append(q)
        save_db(db)
        await status_msg.edit_text(f"✅ မေးခွန်း **({len(quiz_data)})** ခုအား Category **'{default_cat}'** ထဲသို့ အလိုအလျောက် သိမ်းဆည်းလိုက်ပါပြီ။", reply_markup=get_admin_panel_keyboard())
    else:
        await status_msg.edit_text("❌ Gemini AI မှ မေးခွန်း ထုတ်ယူ၍ မရရှိခဲ့ပါ။", reply_markup=get_admin_panel_keyboard())
    return ConversationHandler.END

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("❌ လုပ်ဆောင်ချက်ကို ပယ်ဖျက်လိုက်ပါပြီ။", reply_markup=get_main_menu_keyboard(update.effective_user.id))
    return ConversationHandler.END

def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    # Conversation Handlers
    cat_conv = ConversationHandler(
        entry_points=[CallbackQueryHandler(admin_add_cat_start, pattern="^admin_add_cat$")],
        states={WAIT_CATEGORY_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_save_cat)]},
        fallbacks=[CommandHandler('cancel', cancel)]
    )

    prem_conv = ConversationHandler(
        entry_points=[CallbackQueryHandler(admin_add_prem_start, pattern="^admin_add_premium$")],
        states={WAIT_PREMIUM_ID: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_save_prem)]},
        fallbacks=[CommandHandler('cancel', cancel)]
    )

    ai_chat_conv = ConversationHandler(
        entry_points=[CallbackQueryHandler(admin_ai_chat_start, pattern="^admin_ai_chat$")],
        states={WAIT_AI_PROMPT: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_process_ai_prompt)]},
        fallbacks=[CommandHandler('cancel', cancel)]
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(cat_conv)
    app.add_handler(prem_conv)
    app.add_handler(ai_chat_conv)
    app.add_handler(CallbackQueryHandler(handle_quiz_answer, pattern="^ans_"))
    app.add_handler(CallbackQueryHandler(handle_callback))

    print("Bot is running seamlessly...")
    app.run_polling()

if __name__ == '__main__':
    main()
