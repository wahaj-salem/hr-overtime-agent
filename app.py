"""
HR Overtime Automation Agent v2.0
Built with Streamlit + Claude AI
"""

import streamlit as st
import anthropic
import base64
import json
import re
import pandas as pd
from datetime import datetime
from io import BytesIO
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# ============================================================================
# Page Config
# ============================================================================

st.set_page_config(
    page_title="HR Overtime Agent",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================================
# Custom CSS
# ============================================================================

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #1F4E78 0%, #2E86AB 100%);
        padding: 2rem;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin-bottom: 2rem;
    }
    .stat-card {
        background: #f0f2f6;
        padding: 1rem;
        border-radius: 10px;
        text-align: center;
        border-left: 4px solid #1F4E78;
    }
    .success-box {
        background: #d4edda;
        color: #155724;
        padding: 1rem;
        border-radius: 5px;
        border-left: 4px solid #28a745;
    }
    .stButton button {
        background: #1F4E78;
        color: white;
        font-weight: bold;
        border-radius: 5px;
        padding: 0.5rem 2rem;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================================
# Header
# ============================================================================

st.markdown("""
<div class="main-header">
    <h1>🤖 HR Overtime Automation Agent</h1>
    <p style="font-size: 1.2rem; margin: 0;">
        نظام ذكي لاستخراج بيانات الساعات الإضافية والتقييمات باستخدام Claude AI
    </p>
</div>
""", unsafe_allow_html=True)

# ============================================================================
# Sidebar
# ============================================================================

with st.sidebar:
    st.markdown("## ⚙️ الإعدادات")
    
    api_key = st.text_input(
        "🔑 Claude API Key",
        type="password",
        help="من: https://console.anthropic.com"
    )
    
    if not api_key:
        st.warning("⚠️ يرجى إدخال مفتاح API")
    
    st.markdown("---")
    st.markdown("## 🎛️ خيارات")
    
    sort_order = st.radio(
        "ترتيب الموظفين:",
        ["من الأصغر للأكبر (ID)", "حسب الاسم", "بدون ترتيب"]
    )
    
    unclear_handling = st.radio(
        "للخلايا غير الواضحة:",
        ["كتابة 'يتم المراجعة'", "ترك فارغ", "كتابة 0"]
    )
    
    st.markdown("---")
    st.markdown("## 📊 إعدادات التقييم")
    
    extract_rating = st.checkbox(
        "✅ استخراج عمود التقييم",
        value=True
    )
    
    auto_convert_rating = st.checkbox(
        "🔄 تحويل تلقائي (5 → 50)",
        value=True
    )
    
    if extract_rating and auto_convert_rating:
        st.info("""
        💡 **القاعدة:**
        - 5 → 50
        - 8 → 80
        - 50 → 50
        - 100 → 100
        """)
    
    st.markdown("---")
    st.markdown("**v2.0** | © 2026")

# ============================================================================
# Helper Functions
# ============================================================================

def encode_image(image_bytes):
    return base64.standard_b64encode(image_bytes).decode('utf-8')

def get_media_type(filename):
    ext = filename.lower().split('.')[-1]
    types = {
        'png': 'image/png',
        'jpg': 'image/jpeg',
        'jpeg': 'image/jpeg',
        'webp': 'image/webp',
        'gif': 'image/gif'
    }
    return types.get(ext, 'image/jpeg')

def convert_rating(rating_value, auto_convert=True):
    """تحويل التقييم: إذا كان رقم واحد (5, 8) → اضربه في 10"""
    
    if not auto_convert:
        return rating_value
    
    if not rating_value or rating_value in ['', 'يتم المراجعة', 'F', 'M']:
        return rating_value
    
    try:
        num_value = float(str(rating_value).strip())
        
        # رقم من 1 إلى 9 → اضرب في 10
        if 1 <= num_value < 10:
            return str(int(num_value * 10))
        
        # كما هو
        if num_value == int(num_value):
            return str(int(num_value))
        return str(num_value)
        
    except (ValueError, TypeError):
        return rating_value

def extract_data_from_image(client, image_bytes, filename, unclear_text, extract_rating=True):
    """استخراج البيانات من الصورة"""
    
    image_data = encode_image(image_bytes)
    media_type = get_media_type(filename)
    
    rating_instruction = ""
    rating_json_field = ""
    
    if extract_rating:
        rating_instruction = """
4. RATING (تقييم الموظف) - رقم من 1 إلى 100
   - قد يكون مكتوب كرقم واحد (5, 8) أو رقم كامل (50, 80)
   - استخرج القيمة كما هي مكتوبة بالضبط"""
        rating_json_field = ',\n      "rating": "80"'
    
    prompt = f"""أنت خبير في قراءة جداول الساعات الإضافية للموظفين.

اقرأ الصورة بعناية واستخرج البيانات التالية لكل موظف:
1. ID (كود الموظف)
2. NAME (اسم الموظف كامل)
3. الساعات لكل يوم (التاريخ يكون في رؤوس الأعمدة مثل 1/6, 2/6){rating_instruction}

قواعد مهمة:
- إذا كان هناك خط يد غير واضح تماماً، اكتب "{unclear_text}"
- إذا كانت الخلية فارغة، اكتب ""
- إذا كان فيها F أو حرف، اكتبه كما هو
- لا تخمن! إذا لم تكن متأكد 100%، اكتب "{unclear_text}"

أرجع JSON فقط:
{{
  "month": "06",
  "year": "2026",
  "dates": ["1/6", "2/6"],
  "employees": [
    {{
      "id": "1399",
      "name": "hamza nashat",
      "hours": {{"1/6": "4", "2/6": "4"}}{rating_json_field}
    }}
  ]
}}

ابدأ بـ {{ مباشرة."""
    
    try:
        message = client.messages.create(
            model="claude-opus-4-7",
            max_tokens=8000,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": media_type,
                                "data": image_data
                            }
                        },
                        {"type": "text", "text": prompt}
                    ]
                }
            ]
        )
        
        response_text = message.content[0].text
        json_match = re.search(r'\{[\s\S]*\}', response_text)
        if json_match:
            return json.loads(json_match.group())
        return None
        
    except Exception as e:
        st.error(f"خطأ في معالجة {filename}: {str(e)}")
        return None

def merge_all_data(all_data, extract_rating=True, auto_convert=True):
    """دمج البيانات من كل الصور"""
    
    all_employees = {}
    all_dates = set()
    month_info = None
    
    for data in all_data:
        if not data:
            continue
            
        if not month_info:
            month_info = {
                'month': data.get('month', ''),
                'year': data.get('year', '')
            }
        
        for date in data.get('dates', []):
            all_dates.add(date)
        
        for emp in data.get('employees', []):
            emp_id = emp.get('id')
            if not emp_id:
                continue
            
            if emp_id not in all_employees:
                all_employees[emp_id] = {
                    'ID': emp_id,
                    'NAME': emp.get('name', ''),
                    'hours': {},
                    'rating': ''
                }
            
            for date, hours in emp.get('hours', {}).items():
                all_employees[emp_id]['hours'][date] = hours
            
            if extract_rating:
                rating = emp.get('rating', '')
                if rating:
                    converted_rating = convert_rating(rating, auto_convert)
                    all_employees[emp_id]['rating'] = converted_rating
    
    return all_employees, sorted(all_dates, key=lambda x: int(x.split('/')[0])), month_info

def create_dataframe(employees, sorted_dates, sort_order, extract_rating=True):
    """إنشاء DataFrame"""
    
    rows = []
    for emp_id, emp_data in employees.items():
        row = {
            'ID': emp_data['ID'],
            'NAME': emp_data['NAME']
        }
        
        for date in sorted_dates:
            row[date] = emp_data['hours'].get(date, '')
        
        if extract_rating:
            row['التقييم'] = emp_data.get('rating', '')
        
        rows.append(row)
    
    df = pd.DataFrame(rows)
    
    if "الأصغر للأكبر" in sort_order:
        df['ID_sort'] = df['ID'].astype(str).str.extract('(\d+)').astype(float)
        df = df.sort_values('ID_sort').drop('ID_sort', axis=1).reset_index(drop=True)
    elif "حسب الاسم" in sort_order:
        df = df.sort_values('NAME').reset_index(drop=True)
    
    return df

def create_excel_file(df, month_info):
    """إنشاء ملف Excel منسق"""
    
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sheet_name = f"Month_{month_info.get('month', '00')}_{month_info.get('year', '0000')}"
        df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
        
        worksheet = writer.sheets[sheet_name[:31]]
        
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        review_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
        rating_fill = PatternFill(start_color="E8F4F8", end_color="E8F4F8", fill_type="solid")
        
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # تنسيق رؤوس الأعمدة
        rating_col_idx = None
        for col_idx, cell in enumerate(worksheet[1], 1):
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border
            
            if cell.value == 'التقييم':
                rating_col_idx = col_idx
        
        # تنسيق الخلايا
        for row in worksheet.iter_rows(min_row=2):
            for cell in row:
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center')
                
                if cell.value == 'يتم المراجعة':
                    cell.fill = review_fill
                    cell.font = Font(bold=True, color="9C5700")
                
                if rating_col_idx and cell.column == rating_col_idx:
                    if cell.value and cell.value != 'يتم المراجعة':
                        cell.fill = rating_fill
                        cell.font = Font(bold=True, color="1F4E78")
        
        # عرض الأعمدة
        for col in worksheet.columns:
            max_length = 0
            column_letter = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            worksheet.column_dimensions[column_letter].width = min(max_length + 2, 30)
        
        worksheet.freeze_panes = 'C2'
    
    output.seek(0)
    return output

# ============================================================================
# Main UI
# ============================================================================

tab1, tab2, tab3 = st.tabs(["📤 رفع ومعالجة", "📊 النتائج", "ℹ️ التعليمات"])

# Tab 1: Upload
with tab1:
    st.markdown("### 📤 رفع الصور")
    
    uploaded_files = st.file_uploader(
        "اسحب الصور هنا أو اضغط للاختيار",
        type=['png', 'jpg', 'jpeg', 'webp'],
        accept_multiple_files=True
    )
    
    if uploaded_files:
        st.success(f"✅ تم رفع {len(uploaded_files)} صورة")
        
        with st.expander("👀 معاينة الصور"):
            cols = st.columns(min(3, len(uploaded_files)))
            for idx, file in enumerate(uploaded_files):
                with cols[idx % 3]:
                    st.image(file, caption=file.name, use_container_width=True)
        
        if api_key:
            if st.button("🚀 بدء المعالجة", use_container_width=True):
                
                unclear_text_map = {
                    "كتابة 'يتم المراجعة'": "يتم المراجعة",
                    "ترك فارغ": "",
                    "كتابة 0": "0"
                }
                unclear_text = unclear_text_map.get(unclear_handling, "يتم المراجعة")
                
                try:
                    client = anthropic.Anthropic(api_key=api_key)
                    
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    all_data = []
                    
                    for idx, file in enumerate(uploaded_files):
                        status_text.text(f"⏳ معالجة {idx+1}/{len(uploaded_files)}: {file.name}")
                        
                        file_bytes = file.getvalue()
                        data = extract_data_from_image(
                            client, file_bytes, file.name,
                            unclear_text, extract_rating
                        )
                        
                        if data:
                            all_data.append(data)
                        
                        progress_bar.progress((idx + 1) / len(uploaded_files))
                    
                    status_text.text("✅ تمت المعالجة!")
                    
                    employees, sorted_dates, month_info = merge_all_data(
                        all_data, extract_rating, auto_convert_rating
                    )
                    df = create_dataframe(employees, sorted_dates, sort_order, extract_rating)
                    
                    st.session_state['df'] = df
                    st.session_state['month_info'] = month_info
                    st.session_state['employees_count'] = len(employees)
                    st.session_state['dates_count'] = len(sorted_dates)
                    st.session_state['extract_rating'] = extract_rating
                    
                    st.markdown("""
                    <div class="success-box">
                        <h3>✅ تمت المعالجة بنجاح!</h3>
                        <p>اذهب إلى تبويب <strong>📊 النتائج</strong></p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.balloons()
                    
                except Exception as e:
                    st.error(f"❌ خطأ: {str(e)}")
        else:
            st.warning("⚠️ يرجى إدخال مفتاح API")

# Tab 2: Results
with tab2:
    st.markdown("### 📊 النتائج")
    
    if 'df' in st.session_state:
        df = st.session_state['df']
        month_info = st.session_state['month_info']
        has_rating = st.session_state.get('extract_rating', False)
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.markdown(f"""
            <div class="stat-card">
                <h3>👥 الموظفين</h3>
                <h1>{st.session_state.get('employees_count', 0)}</h1>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class="stat-card">
                <h3>📅 الأيام</h3>
                <h1>{st.session_state.get('dates_count', 0)}</h1>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            review_count = (df == 'يتم المراجعة').sum().sum()
            st.markdown(f"""
            <div class="stat-card">
                <h3>⚠️ للمراجعة</h3>
                <h1>{review_count}</h1>
            </div>
            """, unsafe_allow_html=True)
        
        with col4:
            if has_rating and 'التقييم' in df.columns:
                ratings = pd.to_numeric(df['التقييم'], errors='coerce').dropna()
                avg_rating = ratings.mean() if len(ratings) > 0 else 0
                st.markdown(f"""
                <div class="stat-card">
                    <h3>⭐ متوسط التقييم</h3>
                    <h1>{avg_rating:.0f}</h1>
                </div>
                """, unsafe_allow_html=True)
            else:
                month = month_info.get('month', '00')
                year = month_info.get('year', '0000')
                st.markdown(f"""
                <div class="stat-card">
                    <h3>📆 الشهر</h3>
                    <h1>{month}/{year}</h1>
                </div>
                """, unsafe_allow_html=True)
        
        st.markdown("---")
        st.markdown("#### 📋 جدول البيانات")
        
        if has_rating:
            st.info("💡 عمود التقييم: القيم محولة تلقائياً (5 → 50)")
        
        edited_df = st.data_editor(
            df,
            use_container_width=True,
            num_rows="dynamic",
            height=400
        )
        
        st.markdown("---")
        st.markdown("#### 💾 تحميل الملف")
        
        col1, col2 = st.columns(2)
        
        month = month_info.get('month', '00')
        year = month_info.get('year', '0000')
        
        with col1:
            excel_file = create_excel_file(edited_df, month_info)
            filename = f"Overtime_Report_{month}_{year}.xlsx"
            
            st.download_button(
                label="📥 تحميل Excel",
                data=excel_file,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        with col2:
            csv = edited_df.to_csv(index=False, encoding='utf-8-sig')
            csv_filename = f"Overtime_Report_{month}_{year}.csv"
            
            st.download_button(
                label="📥 تحميل CSV",
                data=csv,
                file_name=csv_filename,
                mime="text/csv",
                use_container_width=True
            )
    else:
        st.info("ℹ️ لا توجد بيانات. ارفع الصور أولاً")

# Tab 3: Instructions
with tab3:
    st.markdown("### 📖 كيفية الاستخدام")
    
    st.markdown("""
    #### 🚀 الخطوات:
    
    **1️⃣ احصل على Claude API Key**
    - من: https://console.anthropic.com
    
    **2️⃣ ضع المفتاح في الشريط الجانبي**
    
    **3️⃣ ارفع الصور**
    
    **4️⃣ اضبط الإعدادات:**
    - ترتيب الموظفين
    - معالجة الخلايا غير الواضحة
    - 🆕 تفعيل التقييم
    - 🆕 التحويل التلقائي (5 → 50)
    
    **5️⃣ اضغط "بدء المعالجة"**
    
    **6️⃣ حمل Excel أو CSV**
    
    ---
    
    #### 🆕 ميزة التقييم:
    
    **القاعدة:**
    - `5` → `50`
    - `8` → `80`  
    - `50` → `50`
    - `100` → `100`
    
    يمكن تعطيل التحويل من الإعدادات
    """)

st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #666;'>Built with ❤️ using Streamlit + Claude AI | v2.0</div>",
    unsafe_allow_html=True
)
