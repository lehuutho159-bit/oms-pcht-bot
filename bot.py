import os
import glob
import datetime
import pandas as pd
from telegram import Update, ReplyKeyboardMarkup, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes, MessageHandler, CallbackQueryHandler, filters

TOKEN = "9098342871:AAH589i1e876M3_NYgvlue2YSG3ZSthjbFM"

DATA_DIR = r"C:\Users\LH Tho\Desktop\OMS\DATA"

def get_main_keyboard():
    keyboard = [
        ["Báo cáo thực hiện ĐTC PCHT", "💡 Biện pháp & Kiến nghị"],
        ["Xuất Excel OMS", "Gửi file Excel"],
        ["📊 Theo dõi bất thường", "Trạng thái"]
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Xin chào! Hệ thống Bot PCHT OMS đã sẵn sàng hoạt động.",
        reply_markup=get_main_keyboard()
    )

def load_data_directly():
    """Đọc trực tiếp 3 file từ thư mục DATA"""
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR, exist_ok=True)
        return None, None, None
        
    excel_files = [f for f in glob.glob(os.path.join(DATA_DIR, "*.xls*")) if not os.path.basename(f).startswith("~$")]
    
    file_thc, file_kh, file_sc = None, None, None
    for f in excel_files:
        fname = os.path.basename(f).lower()
        if "chung" in fname:
            file_thc = f
        elif "ke" in fname or "hoach" in fname or "kế hoạch" in fname:
            file_kh = f
        elif "su" in fname or "co" in fname or "sự cố" in fname:
            file_sc = f
            
    if not file_thc and len(excel_files) >= 1: file_thc = excel_files[0]
    if not file_kh and len(excel_files) >= 2: file_kh = excel_files[1]
    if not file_sc and len(excel_files) >= 3: file_sc = excel_files[2]
    
    df_thc, df_kh, df_sc = None, None, None
    try:
        if file_thc and os.path.exists(file_thc):
            df_thc = pd.read_excel(file_thc, sheet_name=0, header=None)
        if file_kh and os.path.exists(file_kh):
            df_kh = pd.read_excel(file_kh, sheet_name=0, header=None)
        if file_sc and os.path.exists(file_sc):
            df_sc = pd.read_excel(file_sc, sheet_name=0, header=None)
    except Exception as e:
        print(f"Lỗi đọc file: {e}")
        
    return df_thc, df_kh, df_sc

def find_month_columns(df_raw, target_month):
    for r_idx in [6, 7, 5]:
        if r_idx < len(df_raw):
            row_vals = df_raw.iloc[r_idx].values
            for c_idx, val in enumerate(row_vals):
                try:
                    if int(val) == target_month:
                        return c_idx, c_idx + 1
                except:
                    pass
    month_mapping = {6: 6, 7: 9, 8: 12, 9: 15, 10: 18, 11: 21, 12: 24}
    base_col = month_mapping.get(target_month, 21)
    return base_col, base_col + 1

def generate_report(mode='thang'):
    """Báo cáo chi tiết phân tách rạch ròi 3 bộ chỉ tiêu với đúng quy tắc số lượng đơn vị và chỉ tiêu"""
    df_thc, df_kh, df_sc = load_data_directly()
    if df_thc is None:
        return f"❌ Không tìm thấy file dữ liệu trong thư mục `{DATA_DIR}`!"
    
    target_month = datetime.datetime.now().month
    th_col, kh_col = (3, 4) if mode == 'luy_ke' else find_month_columns(df_thc, target_month)
    title = f"BÁO CÁO LŨY KẾ NĂM - OMS PCHT" if mode == 'luy_ke' else f"BÁO CÁO THÁNG {target_month} - OMS PCHT"
    
    # Định nghĩa quy tắc cho từng bộ chỉ tiêu: (Tên bộ, DataFrame, danh sách chỉ tiêu được phép tính)
    datasets = [
        ("THC lưới điện phân phối", df_thc, ['MAIFI', 'SAIDI', 'SAIFI']),     # 14 đơn vị x 3 chỉ tiêu = 42
        ("Sự cố trước miễn trừ", df_sc, ['MAIFI', 'SAIDI', 'SAIFI']),        # 14 đơn vị x 3 chỉ tiêu = 42
        ("Kế hoạch trước miễn trừ", df_kh, ['SAIDI', 'SAIFI'])               # 14 đơn vị x 2 chỉ tiêu = 28 (không tính MAIFI)
    ]
    
    category_reports = []
    
    for cat_name, df_raw, allowed_indicators in datasets:
        if df_raw is None:
            category_reports.append(f"📁 **Bộ chỉ tiêu: {cat_name}**\n  ❌ (Chưa có dữ liệu file)\n")
            continue
            
        try:
            df = df_raw.iloc[8:].reset_index(drop=True)
            red_list = []
            yellow_list = []
            cat_total = 0
            cat_good = 0
            current_unit = "Công ty Điện lực Hà Tĩnh"
            
            for index, row in df.iterrows():
                col1 = row.get(1)
                chi_tieu = str(row.get(2, '')).strip()
                
                if pd.notna(col1) and str(col1).strip() != '':
                    current_unit = str(col1).strip()
                
                # Chỉ xét các chỉ tiêu thuộc danh mục cho phép của bộ đó
                if chi_tieu not in allowed_indicators:
                    continue
                
                try:
                    th = float(row.get(th_col, 0))
                    kh = float(row.get(kh_col, 0))
                except:
                    continue
                    
                if kh == 0:
                    continue
                    
                cat_total += 1
                rate = (th / kh) * 100
                item_str = f"{current_unit} - {chi_tieu}: TH {th} / KH {kh} ({rate:.1f}%)"
                
                if rate >= 100.0:
                    excess = round(th - kh, 4)
                    red_list.append(f"• {item_str}, vượt KH {excess}")
                elif 90.0 <= rate < 100.0:
                    diff = round(kh - th, 4)
                    yellow_list.append(f"• {item_str}, còn {diff}")
                else:
                    cat_good += 1
                    
            cat_rate = (cat_good / cat_total * 100) if cat_total > 0 else 0
            
            cat_msg = f"📁 **Bộ chỉ tiêu: {cat_name}**\n"
            cat_msg += f"  • Tổng số: {cat_total} chỉ tiêu | Đạt tốt: {cat_good} ({cat_rate:.1f}%) | Vượt KH: {len(red_list)} | Cảnh báo: {len(yellow_list)}\n"
            
            if red_list:
                cat_msg += "  🔴 Vượt KH (>= 100%):\n    " + "\n    ".join(red_list) + "\n"
            if yellow_list:
                cat_msg += "  🟡 Cảnh báo (90% - < 100%):\n    " + "\n    ".join(yellow_list) + "\n"
            if not red_list and not yellow_list:
                cat_msg += "  🟢 Đạt tốt tất cả các chỉ tiêu (< 90%)\n"
                
            category_reports.append(cat_msg + "\n")
        except Exception as e:
            category_reports.append(f"📁 **Bộ chỉ tiêu: {cat_name}**\n  ⚠️ Lỗi đọc file ({str(e)})\n")

    msg = f"=== {title} ===\n\n"
    msg += "📋 **THỐNG KÊ CHI TIẾT TỪNG BỘ CHỈ TIÊU (CÔNG TY & 13 ĐIỆN LỰC):**\n\n"
    msg += "".join(category_reports)
    
    return msg

def get_recommendation_report():
    """Module Biện pháp & Kiến nghị chuẩn xác theo mẫu yêu cầu"""
    df_thc, _, _ = load_data_directly()
    if df_thc is None:
        return "Không tìm thấy file dữ liệu Tổng hợp chung trong thư mục DATA!"
        
    try:
        target_month = datetime.datetime.now().month
        th_col, kh_col = find_month_columns(df_thc, target_month)
        
        df = df_thc.iloc[8:].reset_index(drop=True)
        units_status = {}
        current_unit = "Công ty Điện lực Hà Tĩnh"
        
        for index, row in df.iterrows():
            col1 = row.get(1)
            chi_tieu = str(row.get(2, '')).strip()
            
            if pd.notna(col1) and str(col1).strip() != '':
                current_unit = str(col1).strip()
            
            if chi_tieu not in ['MAIFI', 'SAIDI', 'SAIFI']:
                continue
            
            try:
                th = float(row.get(th_col, 0))
                kh = float(row.get(kh_col, 0))
            except:
                continue
                
            if kh == 0:
                continue
                
            rate = (th / kh) * 100
            
            if rate >= 90.0:
                if current_unit not in units_status:
                    units_status[current_unit] = {'MAIFI': False, 'SAIDI': False, 'SAIFI': False, 'details': []}
                
                units_status[current_unit][chi_tieu.upper()] = True
                status_text = "Vượt kế hoạch" if rate >= 100.0 else "Cảnh báo sát ngưỡng"
                units_status[current_unit]['details'].append(f"{chi_tieu}: {rate:.1f}% ({status_text} - TH: {th} / KH: {kh})")

        recommendations = []
        for unit, data in units_status.items():
            has_maifi = data.get('MAIFI', False)
            has_said = data.get('SAIDI', False)
            has_saifi = data.get('SAIFI', False)
            
            if has_maifi and (has_said or has_saifi):
                solution = "Kiến nghị thay đổi mã nguyên nhân đối với các Mã LMĐ trên 0 phút để được tối ưu chỉ số MAIFI, SAIDI"
            elif has_maifi and not has_said and not has_saifi:
                solution = "Kiến nghị thay đổi mã nguyên nhân đối với các Mã LMĐ dưới 5 phút để tối ưu hóa chỉ số MAIFI."
            elif (has_said or has_saifi) and not has_maifi:
                solution = "Kiến nghị thay đổi mã nguyên nhân đối với các Mã LMĐ trên 5 phút để tối ưu hóa chỉ số SAIDI, SAIFI."
            else:
                solution = "Rà soát toàn diện và kiến nghị mã lịch miễn trừ theo quy định."

            details_lines = "\n     • ".join(data['details'])
            recommendations.append(f"**{unit}**:\n     • {details_lines}\n     💡 Đề xuất giải pháp: {solution}")

        msg = f"📋 RÀ SOÁT & KIẾN NGHỊ ĐỘ TIN CẬY - THÁNG {target_month}\n"
        msg += "-------------------------------------------------------\n"
        msg += f"⚠️ Tổng số đơn vị cần thực hiện kiến nghị (điều kiện duyệt tự động >= 90%): {len(recommendations)} đơn vị\n\n\n"
        
        if recommendations:
            for i, item in enumerate(recommendations, 1):
                msg += f"{i}. {item}\n\n"
        else:
            msg += f"(Tháng {target_month} hiện tại không có đơn vị nào đạt từ 90% trở lên).\n"
            
        return msg
    except Exception as e:
        return f"Lỗi xử lý kiến nghị: {str(e)}"

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text

    if text == "Báo cáo thực hiện ĐTC PCHT":
        keyboard = [
            [InlineKeyboardButton("📊 Lũy kế năm", callback_data="report_luyke"),
             InlineKeyboardButton("📅 Số liệu tháng", callback_data="report_thang")]
        ]
        reply_markup = InlineKeyboardMarkup(keyboard)
        await update.message.reply_text("Vui lòng chọn loại báo cáo thực hiện ĐTC PCHT:", reply_markup=reply_markup)
        
    elif text == "💡 Biện pháp & Kiến nghị":
        rec_msg = get_recommendation_report()
        await update.message.reply_text(rec_msg, parse_mode="Markdown")
        
    elif text == "Xuất Excel OMS":
        await update.message.reply_text("📁 Đang trích xuất file Excel dữ liệu hệ thống...")
    elif text == "Gửi file Excel":
        await update.message.reply_text("📎 Vui lòng gửi file Excel báo cáo cần xử lý.")
    elif text == "📊 Theo dõi bất thường":
        await update.message.reply_text("⚠️ Đang kiểm tra các điểm vận hành bất thường...")
    elif text == "Trạng thái":
        target_month = datetime.datetime.now().month
        await update.message.reply_text(f"🟢 Hệ thống OMS PCHT đang chạy ngầm, bám sát dữ liệu thực tế Tháng {target_month}.")

async def button_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    if query.data == "report_luyke":
        report_content = generate_report(mode='luy_ke')
        await query.message.reply_text(report_content)
    elif query.data == "report_thang":
        report_content = generate_report(mode='thang')
        await query.message.reply_text(report_content)

if __name__ == '__main__':
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(button_callback))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))
    
    print("Bot đang chạy với phân tách chuẩn xác THC (42), Sự cố (42) và Kế hoạch (28 chỉ tiêu)...")
    app.run_polling()