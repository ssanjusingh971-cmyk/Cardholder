import streamlit as st
import pandas as pd
import google.generativeai as genai
import os
from PIL import Image
import json
import uuid
import io
import re
import time

# ==========================================
# 🔑 PERMANENT API KEY SETUP
# ==========================================
# Niche quotes ke andar apni Gemini API Key daalein
API_KEY = "" 

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="QSS Global - SmartCard Pro", page_icon="📇", layout="wide", initial_sidebar_state="expanded")

# --- CUSTOM STYLING ---
st.markdown("""
<style>
    .main-title { font-size: 2.3rem; font-weight: 800; color: #1D4ED8; margin-bottom: 0px; }
    .sub-title { font-size: 1.05rem; color: #6B7280; margin-bottom: 20px; font-weight: 500; }
    .metric-box { background-color: #F8FAFC; padding: 18px; border-radius: 10px; text-align: center; border: 1px solid #E2E8F0; }
    .metric-val { font-size: 1.8rem; font-weight: 800; color: #2563EB; }
    .metric-label { font-size: 0.95rem; color: #475569; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

# --- DIRECTORIES & DATABASE SETUP ---
DATA_FILE = "qss_database.csv"
IMAGE_DIR = "card_images"
LOGO_FILE = "Qss.png" # <--- YAHAN NAME UPDATE KIYA HAI
os.makedirs(IMAGE_DIR, exist_ok=True)

COLUMNS = ["Company Name", "Contact Person Name", "Designation", "Mail ID", "Phone Number", "Address", "City", "State", "Website Name", "Front_Image", "Back_Image"]

def load_data():
    if os.path.exists(DATA_FILE):
        return pd.read_csv(DATA_FILE)
    return pd.DataFrame(columns=COLUMNS)

def save_data(df):
    df.to_csv(DATA_FILE, index=False)

if 'db' not in st.session_state:
    st.session_state.db = load_data()

if 'staging_cards' not in st.session_state:
    st.session_state.staging_cards = []

# --- INITIALIZE GEMINI MODEL (gemini-3.6-flash WITH STRICT JSON) ---
model = None
active_key = API_KEY if API_KEY else st.secrets.get("GEMINI_API_KEY", "")
if active_key:
    try:
        genai.configure(api_key=active_key)
        model = genai.GenerativeModel(
            'gemini-3.6-flash',
            generation_config={"response_mime_type": "application/json"}
        )
    except Exception as e:
        st.sidebar.error(f"API Setup Error: {e}")

# --- DUPLICATE NORMALIZATION ---
def extract_10_digits(phone_str):
    nums = re.findall(r'\d+', str(phone_str))
    all_digits = "".join(nums)
    if len(all_digits) >= 10:
        return all_digits[-10:]
    return all_digits

def check_duplicate(new_name, new_phone, new_email):
    if st.session_state.db.empty: return False
    
    clean_new_name = re.sub(r'[^a-zA-Z0-9]', '', str(new_name)).lower()
    clean_new_phone = extract_10_digits(new_phone)
    clean_new_email = str(new_email).lower().strip()
    
    for _, row in st.session_state.db.iterrows():
        clean_db_name = re.sub(r'[^a-zA-Z0-9]', '', str(row['Contact Person Name'])).lower()
        clean_db_phone = extract_10_digits(row['Phone Number'])
        clean_db_email = str(row['Mail ID']).lower().strip()
        
        name_match = (clean_new_name == clean_db_name) and (clean_new_name != "")
        phone_match = (clean_new_phone == clean_db_phone) and (len(clean_new_phone) == 10)
        email_match = (clean_new_email == clean_db_email) and (clean_new_email not in ["", "not available"])
        
        if name_match and (phone_match or email_match):
            return True
    return False

# --- HIGH-SPEED IMAGE OPTIMIZATION (800px TILE) ---
def compress_and_save(img_file):
    img = Image.open(img_file)
    if img.mode == 'RGBA': img = img.convert('RGB')
    img.thumbnail((800, 800))
    path = os.path.join(IMAGE_DIR, f"{uuid.uuid4().hex[:10]}.jpg")
    img.save(path, optimize=True, quality=75)
    return path, img

# --- HELPER: CHUNKING FOR BATCH PACKING ---
def chunk_items(lst, size=5):
    for i in range(0, len(lst), size):
        yield lst[i:i + size]

# --- EXCEL EXPORT (FORMATTED) ---
def get_excel_download():
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    st.session_state.db.to_excel(writer, index=False, sheet_name='Contacts')
    workbook = writer.book
    worksheet = writer.sheets['Contacts']
    header_fmt = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#1D4ED8', 'border': 1, 'valign': 'vcenter'})
    cell_fmt = workbook.add_format({'valign': 'vcenter', 'text_wrap': True})
    
    for col_num, col_name in enumerate(st.session_state.db.columns.values):
        worksheet.write(0, col_num, col_name, header_fmt)
        worksheet.set_column(col_num, col_num, 22, cell_fmt)
    writer.close()
    return output.getvalue()

# --- SIDEBAR ---
with st.sidebar:
    if os.path.exists(LOGO_FILE):
        st.image(LOGO_FILE, use_container_width=True)
    st.markdown("### ⚙️ System Status")
    if model: 
        st.success("Gemini 3.6-Flash Active 🟢")
        st.caption("📦 5-Card Batch Packing Enabled")
    else: 
        st.error("API Key Missing 🔴")
        st.warning("Please paste your API Key in line 16 of Cardholder.py")
    
    st.markdown("---")
    if not st.session_state.db.empty:
        st.download_button(
            label="📊 Download Formatted Excel",
            data=get_excel_download(),
            file_name="QSS_Global_Contacts.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )

# --- HEADER ---
col_l, col_h = st.columns([1, 8])
with col_l:
    if os.path.exists(LOGO_FILE): st.image(LOGO_FILE, width=75)
with col_h:
    st.markdown('<p class="main-title">QSS Global - Card Digitization Platform</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-title">Batch-Packed OCR Pipeline & Master Directory</p>', unsafe_allow_html=True)
st.markdown("---")

tab1, tab2, tab3, tab4 = st.tabs(["📤 Upload Cards", "🔍 Search Directory", "📈 Analytics & Drilldown", "📊 Master Database"])

# =========================================================
# TAB 1: UPLOAD CARDS (BATCHED PACKING PIPELINE)
# =========================================================
with tab1:
    mode = st.radio("Choose Processing Type:", ["📄 Only Front (Single-Sided Cards)", "📇 Front & Back Combo (Interactive Verification)"], horizontal=True)
    
    # -----------------------------
    # OPTION 1: ONLY FRONT MODE (5 CARDS PER 1 API CALL)
    # -----------------------------
    if "Only Front" in mode:
        st.info("💡 **Batch Packing Active:** 5 images are sent in 1 single API call. 15 cards consume only 3 daily requests instead of 15!")
        single_files = st.file_uploader("Select Front Cards", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True, key="single_up")
        
        if not model and single_files:
            st.error("⚠️ API Key is missing! Please paste your key in line 16 of Cardholder.py.")
            
        if st.button("🚀 Process Single Cards (Batch Mode)", type="primary") and single_files and model:
            batches = list(chunk_items(single_files, size=5))
            total_batches = len(batches)
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            results_to_save = []
            errors = []
            
            for b_idx, batch_files in enumerate(batches):
                status_text.text(f"Processing Batch {b_idx + 1} of {total_batches} ({len(batch_files)} cards in 1 call)...")
                
                batch_paths = []
                batch_images = []
                for f in batch_files:
                    path, img = compress_and_save(f)
                    batch_paths.append(path)
                    batch_images.append(img)
                
                batch_prompt = f"""You are analyzing exactly {len(batch_images)} business card images.
The images are provided in sequential order from Index 0 to Index {len(batch_images)-1}.
Extract the details from each card independently.
Combine multiple emails or phone numbers using commas.
Infer City and State accurately from the address or pincode.

CRITICAL: Return a valid JSON ARRAY containing exactly {len(batch_images)} objects, matching the exact order of the provided images:
[
  {{
    "Company Name": "",
    "Contact Person Name": "",
    "Designation": "",
    "Mail ID": "",
    "Phone Number": "",
    "Address": "",
    "City": "",
    "State": "",
    "Website Name": ""
  }}
]"""
                
                try:
                    res = model.generate_content([batch_prompt] + batch_images)
                    parsed_array = json.loads(res.text)
                    if isinstance(parsed_array, dict):
                        parsed_array = [parsed_array]
                        
                    for i, card_data in enumerate(parsed_array):
                        if i < len(batch_paths):
                            card_data["Front_Image"] = batch_paths[i]
                            card_data["Back_Image"] = "Not available"
                            
                            if not check_duplicate(card_data.get("Contact Person Name"), card_data.get("Phone Number"), card_data.get("Mail ID")):
                                results_to_save.append(card_data)
                except Exception as e:
                    errors.append(f"Batch {b_idx + 1} Error: {str(e)}")
                
                progress_bar.progress((b_idx + 1) / total_batches)
                time.sleep(1) # Safe pause between batches
            
            if results_to_save:
                st.session_state.db = pd.concat([st.session_state.db, pd.DataFrame(results_to_save)], ignore_index=True)
                save_data(st.session_state.db)
                st.success(f"⚡ Batch Processing Complete! {len(results_to_save)} new distinct cards added in only {total_batches} API request(s).")
            
            if errors:
                with st.expander("⚠️ Errors occurred during some batches:", expanded=True):
                    for err in errors:
                        st.error(err)
                        
            if not results_to_save and not errors:
                st.warning("All uploaded cards were identified as duplicates and skipped.")
                
            progress_bar.empty()
            status_text.empty()

    # -----------------------------
    # OPTION 2: COMBO MODE (BATCHED STAGING)
    # -----------------------------
    else:
        st.info("💡 **Combo Batch Active:** Front cards are scanned in 5-card batches to save quota, then staged for verification.")
        c_front, c_back = st.columns(2)
        front_uploads = c_front.file_uploader("1. Upload FRONT Sides", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True, key="f_up")
        back_uploads = c_back.file_uploader("2. Upload BACK Sides (Optional)", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True, key="b_up")
        
        if not model and front_uploads:
            st.error("⚠️ API Key is missing! Please paste your key in line 16 of Cardholder.py.")
            
        if st.button("🔍 Scan & Stage for Verification", type="primary") and front_uploads and model:
            with st.spinner("Processing front cards in 5-card batches..."):
                st.session_state.back_images_pool = []
                if back_uploads:
                    for idx, b_file in enumerate(back_uploads):
                        b_path, b_img = compress_and_save(b_file)
                        st.session_state.back_images_pool.append({
                            "id": f"Back #{idx+1}",
                            "path": b_path,
                            "img": b_img,
                            "filename": b_file.name
                        })
                
                st.session_state.staging_cards = []
                batches = list(chunk_items(front_uploads, size=5))
                
                for batch_files in batches:
                    batch_paths = []
                    batch_images = []
                    for f in batch_files:
                        path, img = compress_and_save(f)
                        batch_paths.append(path)
                        batch_images.append(img)
                    
                    combo_prompt = f"""You are analyzing exactly {len(batch_images)} Front sides of business cards.
The images are provided in sequential order from Index 0 to Index {len(batch_images)-1}.
Extract details from each Front side independently.
Combine multiple emails or phones using commas.
Identify City and State explicitly from the address/pincode.

Return a JSON ARRAY containing {len(batch_images)} objects in exact order:
[
  {{
    "Company Name": "",
    "Contact Person Name": "",
    "Designation": "",
    "Mail ID": "",
    "Phone Number": "",
    "Address": "",
    "City": "",
    "State": "",
    "Website Name": ""
  }}
]"""
                    try:
                        res = model.generate_content([combo_prompt] + batch_images)
                        batch_data = json.loads(res.text)
                        if isinstance(batch_data, dict):
                            batch_data = [batch_data]
                            
                        for i, card_info in enumerate(batch_data):
                            if i < len(batch_paths):
                                card_info["Front_Image"] = batch_paths[i]
                                card_info["Selected_Back_Id"] = st.session_state.back_images_pool[0]["id"] if len(st.session_state.back_images_pool) == 1 else "None"
                                st.session_state.staging_cards.append(card_info)
                    except Exception:
                        for p in batch_paths:
                            st.session_state.staging_cards.append({
                                "Company Name": "Unknown", "Contact Person Name": "Unknown", "Designation": "",
                                "Mail ID": "", "Phone Number": "", "Address": "", "City": "", "State": "",
                                "Website Name": "", "Front_Image": p, "Selected_Back_Id": "None"
                            })
                    time.sleep(1)
                st.rerun()

        # RENDER STAGING VERIFICATION
        if st.session_state.staging_cards:
            st.markdown("### 📋 Interactive Batch Verification")
            st.write("Review the cards below before confirming master save:")
            
            if hasattr(st.session_state, 'back_images_pool') and st.session_state.back_images_pool:
                with st.expander("🖼️ View Back Images Reference Gallery", expanded=True):
                    b_cols = st.columns(min(len(st.session_state.back_images_pool), 6))
                    for idx, b_obj in enumerate(st.session_state.back_images_pool):
                        with b_cols[idx % 6]:
                            st.image(b_obj["path"], caption=f"🏷️ {b_obj['id']}", use_container_width=True)
            
            back_options = ["None"] + [b["id"] for b in getattr(st.session_state, 'back_images_pool', [])]
            back_lookup = {b["id"]: b["path"] for b in getattr(st.session_state, 'back_images_pool', [])}
            
            cards_to_save = []
            
            for idx, card in enumerate(st.session_state.staging_cards):
                with st.container(border=True):
                    col_img1, col_img2, col_inputs = st.columns([1.2, 1.2, 3])
                    
                    with col_img1:
                        st.image(card["Front_Image"], caption="Front Side", use_container_width=True)
                    
                    with col_img2:
                        current_back_id = card.get("Selected_Back_Id", "None")
                        cur_idx = back_options.index(current_back_id) if current_back_id in back_options else 0
                        chosen_back = st.selectbox(f"Assign Back Side #{idx+1}:", back_options, index=cur_idx, key=f"back_sel_{idx}")
                        card["Selected_Back_Id"] = chosen_back
                        
                        if chosen_back != "None" and chosen_back in back_lookup:
                            st.image(back_lookup[chosen_back], caption=f"Selected: {chosen_back}", use_container_width=True)
                        else:
                            st.info("No Back Side linked")
                    
                    with col_inputs:
                        row1_c1, row1_c2 = st.columns(2)
                        p_name = row1_c1.text_input("Person Name", card.get("Contact Person Name", ""), key=f"pname_{idx}")
                        c_name = row1_c2.text_input("Company Name", card.get("Company Name", ""), key=f"cname_{idx}")
                        
                        row2_c1, row2_c2 = st.columns(2)
                        desig = row2_c1.text_input("Designation", card.get("Designation", ""), key=f"desig_{idx}")
                        phone = row2_c2.text_input("Phone Number(s)", card.get("Phone Number", ""), key=f"phone_{idx}")
                        
                        row3_c1, row3_c2 = st.columns(2)
                        email = row3_c1.text_input("Mail ID(s)", card.get("Mail ID", ""), key=f"email_{idx}")
                        website = row3_c2.text_input("Website", card.get("Website Name", ""), key=f"web_{idx}")
                        
                        row4_c1, row4_c2 = st.columns(2)
                        city = row4_c1.text_input("City", card.get("City", ""), key=f"city_{idx}")
                        state = row4_c2.text_input("State", card.get("State", ""), key=f"state_{idx}")
                        
                        addr = st.text_input("Address", card.get("Address", ""), key=f"addr_{idx}")
                        
                        is_dup = check_duplicate(p_name, phone, email)
                        if is_dup:
                            st.warning("⚠️ Notice: This Contact Person & Phone/Email already exists in Database.")
                        
                        include = st.checkbox("Include this Card", value=(not is_dup), key=f"inc_{idx}")
                        
                        if include:
                            cards_to_save.append({
                                "Company Name": c_name,
                                "Contact Person Name": p_name,
                                "Designation": desig,
                                "Mail ID": email,
                                "Phone Number": phone,
                                "Address": addr,
                                "City": city,
                                "State": state,
                                "Website Name": website,
                                "Front_Image": card["Front_Image"],
                                "Back_Image": back_lookup.get(chosen_back, "Not available")
                            })

            c_save, c_clear = st.columns([2, 1])
            with c_save:
                if st.button("💾 Confirm & Save Approved Cards to Master DB", type="primary", use_container_width=True):
                    if cards_to_save:
                        st.session_state.db = pd.concat([st.session_state.db, pd.DataFrame(cards_to_save)], ignore_index=True)
                        save_data(st.session_state.db)
                        st.session_state.staging_cards = []
                        st.success(f"Successfully saved {len(cards_to_save)} cards!")
                        st.rerun()
                    else:
                        st.warning("No cards selected for saving.")
            with c_clear:
                if st.button("🗑️ Discard Staging Batch", use_container_width=True):
                    st.session_state.staging_cards = []
                    st.rerun()

# =========================================================
# TAB 2: SEARCH DIRECTORY
# =========================================================
with tab2:
    search_q = st.text_input("🔍 Search Database...", placeholder="Search across Name, Company, City, State, Designation, Phone...")
    if not st.session_state.db.empty:
        if search_q:
            mask = st.session_state.db.apply(lambda row: row.astype(str).str.contains(search_q, case=False).any(), axis=1)
            results = st.session_state.db[mask]
        else:
            results = st.session_state.db
        
        st.write(f"Showing {len(results)} contact(s):")
        for _, row in results.iterrows():
            with st.container(border=True):
                c1, c2, c3 = st.columns([1.2, 1.2, 3])
                with c1:
                    if os.path.exists(str(row.get('Front_Image', ''))):
                        st.image(str(row.get('Front_Image')), caption="Front (Zoom icon ↗️)", use_container_width=True)
                    else:
                        st.caption("No Front Image")
                with c2:
                    if os.path.exists(str(row.get('Back_Image', ''))):
                        st.image(str(row.get('Back_Image')), caption="Back (Zoom icon ↗️)", use_container_width=True)
                    else:
                        st.caption("No Back Image")
                with c3:
                    st.subheader(row.get('Company Name', 'N/A'))
                    st.write(f"**👤 {row.get('Contact Person Name', 'N/A')}** | 💼 *{row.get('Designation', 'N/A')}*")
                    st.write(f"📞 **Phone:** {row.get('Phone Number', 'N/A')}")
                    st.write(f"📧 **Mail:** {row.get('Mail ID', 'N/A')}")
                    st.write(f"📍 **Location:** {row.get('City', 'N/A')}, {row.get('State', 'N/A')}")
                    st.write(f"🏠 **Address:** {row.get('Address', 'N/A')}")
                    if row.get('Website Name') and row.get('Website Name') != "Not available":
                        st.write(f"🌐 **Website:** {row.get('Website Name')}")

# =========================================================
# TAB 3: ANALYTICS & DRILLDOWN
# =========================================================
with tab3:
    if not st.session_state.db.empty:
        df = st.session_state.db
        m1, m2, m3, m4 = st.columns(4)
        
        with m1: st.markdown(f'<div class="metric-box"><div class="metric-val">{df["Company Name"].nunique()}</div><div class="metric-label">Total Companies</div></div>', unsafe_allow_html=True)
        with m2: st.markdown(f'<div class="metric-box"><div class="metric-val">{len(df)}</div><div class="metric-label">Total Contacts</div></div>', unsafe_allow_html=True)
        
        valid_states = [s for s in df['State'].dropna().unique() if s.strip() not in ["", "Not available", "N/A"]]
        valid_cities = [c for c in df['City'].dropna().unique() if c.strip() not in ["", "Not available", "N/A"]]
        
        with m3: st.markdown(f'<div class="metric-box"><div class="metric-val">{len(valid_states)}</div><div class="metric-label">States Covered</div></div>', unsafe_allow_html=True)
        with m4: st.markdown(f'<div class="metric-box"><div class="metric-val">{len(valid_cities)}</div><div class="metric-label">Cities Covered</div></div>', unsafe_allow_html=True)
        
        st.markdown("---")
        st.markdown("### 🗺️ Geographic Contact Drilldown")
        d_col1, d_col2 = st.columns(2)
        
        selected_state = d_col1.selectbox("Filter by State:", ["All States"] + sorted(valid_states))
        
        filtered_df = df if selected_state == "All States" else df[df['State'] == selected_state]
        state_cities = [c for c in filtered_df['City'].dropna().unique() if c.strip() not in ["", "Not available", "N/A"]]
        selected_city = d_col2.selectbox("Filter by City:", ["All Cities"] + sorted(state_cities))
        
        if selected_city != "All Cities":
            filtered_df = filtered_df[filtered_df['City'] == selected_city]
            
        st.dataframe(
            filtered_df[["Company Name", "Contact Person Name", "Designation", "Phone Number", "Mail ID", "City", "State"]],
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("Database is currently empty. Upload cards to see analytics.")

# =========================================================
# TAB 4: MASTER DATABASE
# =========================================================
with tab4:
    st.info("💡 Edit cells directly or select rows via checkbox to delete them. Click 'Save' to apply changes.")
    edited_df = st.data_editor(st.session_state.db, use_container_width=True, num_rows="dynamic", hide_index=False)
    
    if st.button("💾 Save Database Changes", type="primary"):
        st.session_state.db = edited_df
        save_data(st.session_state.db)
        st.success("Database updated successfully!")
