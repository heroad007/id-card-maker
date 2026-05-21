import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import openpyxl
from openpyxl_image_loader import SheetImageLoader
from PIL import Image, ImageDraw, ImageFont
import io

# --- PAGE SETUP ---
st.set_page_config(page_title="ID Card Generator", page_icon="🪪", layout="centered")

# --- VANTA.JS ANIMATED BACKGROUND ---
# This injects the 3D wave animation directly into the root Streamlit background
vanta_script = """
<script>
// Check if Vanta is already loaded to prevent duplicate animations
if (!window.parent.document.getElementById('vanta-waves-script')) {
    // 1. Add Three.js
    const threeScript = window.parent.document.createElement('script');
    threeScript.src = 'https://cdnjs.cloudflare.com/ajax/libs/three.js/r121/three.min.js';
    window.parent.document.head.appendChild(threeScript);

    // 2. Add Vanta.js after Three.js loads
    threeScript.onload = function() {
        const vantaScript = window.parent.document.createElement('script');
        vantaScript.src = 'https://cdn.jsdelivr.net/npm/vanta@latest/dist/vanta.waves.min.js';
        vantaScript.id = 'vanta-waves-script';
        window.parent.document.head.appendChild(vantaScript);
        
        // 3. Initialize Vanta Waves
        vantaScript.onload = function() {
            window.parent.VANTA.WAVES({
                el: window.parent.document.querySelector(".stApp"),
                mouseControls: true,
                touchControls: true,
                gyroControls: false,
                minHeight: 200.00,
                minWidth: 200.00,
                scale: 1.00,
                scaleMobile: 1.00,
                color: 0x7a0000,  /* Deep Maroon color to match your ID card */
                shininess: 30.00,
                waveHeight: 15.00,
                waveSpeed: 0.70,
                zoom: 0.75
            });
        };
    };
}
</script>
"""
components.html(vanta_script, width=0, height=0)

# --- CUSTOM CSS FOR FROSTED GLASS EFFECT ---
st.markdown("""
    <style>
        /* Make Streamlit default background transparent so Vanta shows through */
        .stApp {
            background-color: transparent !important;
        }
        header {
            background-color: transparent !important;
        }
        
        /* Frosted Glass Effect on the main container */
        .block-container {
            background: rgba(255, 255, 255, 0.85); /* Semi-transparent white */
            border-radius: 15px;
            padding: 2rem 3rem !important;
            margin-top: 3rem;
            margin-bottom: 3rem;
            box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.3);
            backdrop-filter: blur(10px);
            -webkit-backdrop-filter: blur(10px);
        }

        /* Title Styling */
        .main-title {
            color: #8B0000;
            text-align: center;
            font-size: 38px;
            font-weight: 900;
            margin-bottom: 5px;
            font-family: 'Arial', sans-serif;
        }
        .sub-title {
            text-align: center;
            color: #444444;
            font-size: 16px;
            margin-bottom: 30px;
        }

        /* Custom Table Styling */
        .custom-table {
            width: 100%;
            border-collapse: collapse;
            margin: 15px 0;
            font-size: 14px;
            font-family: 'Arial', sans-serif;
            background-color: #ffffff;
            border-radius: 8px 8px 0 0;
            overflow: hidden;
            box-shadow: 0 4px 8px rgba(0,0,0,0.05);
        }
        .custom-table thead tr {
            background-color: #8B0000;
            color: #ffffff;
            text-align: left;
        }
        .custom-table th, .custom-table td {
            padding: 10px 15px;
            border: 1px solid #dddddd;
        }
        /* Mobile Scrollable Table */
        .table-responsive {
            overflow-x: auto;
        }
    </style>
""", unsafe_allow_html=True)

# --- HEADER ---
st.markdown('<div class="main-title">🪪 Automated ID Card Generator</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Upload your Excel file to generate a printable A4 PDF of student ID cards.</div>', unsafe_allow_html=True)

# --- INSTRUCTIONS & STYLED TABLE ---
st.info("📌 **INSTRUCTIONS:** Row 1 of your Excel file must contain exactly these headers (no punctuation). Place the student's image inside the cell in Column H.")

table_html = """
<div class="table-responsive">
    <table class="custom-table">
        <thead>
            <tr>
                <th>Adm No</th>
                <th>Name</th>
                <th>Father Name</th>
                <th>Class</th>
                <th>DOB</th>
                <th>Address</th>
                <th>Mobile</th>
                <th>Photo</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>21325</td>
                <td>PRIYANSHU SISODIYA</td>
                <td>SUNIL SISODIYA</td>
                <td>10th</td>
                <td>16.SEP.2009</td>
                <td>DOHAI</td>
                <td>8126528893</td>
                <td style="color:blue;"><i>[Image]</i></td>
            </tr>
        </tbody>
    </table>
</div>
<br>
"""
st.markdown(table_html, unsafe_allow_html=True)

# --- SETTINGS ---
TEMPLATE_FILE = 'template.jpg'
PHOTO_COLUMN = 'H'
PHOTO_X, PHOTO_Y = 248, 208        
PHOTO_WIDTH, PHOTO_HEIGHT = 245, 310   

# Load Fonts safely
try:
    font = ImageFont.truetype("arialbd.ttf", 26)
    font_small = ImageFont.truetype("arialbd.ttf", 22)
except Exception as e:
    st.error("⚠️ Font file 'arialbd.ttf' is missing from GitHub! Please upload it.")
    st.stop()

# --- FILE UPLOADER ---
uploaded_file = st.file_uploader("📂 Upload your Excel File", type=["xlsx"])

if uploaded_file is not None:
    if st.button("🚀 Generate ID Cards", type="primary", use_container_width=True):
        with st.spinner("Processing data and extracting photos... Please wait."):
            try:
                uploaded_file.seek(0) 
                df = pd.read_excel(uploaded_file)
                df.columns = df.columns.str.strip()
                df = df.fillna("")
                
                uploaded_file.seek(0) 
                wb = openpyxl.load_workbook(uploaded_file, data_only=True)
                sheet = wb.active
                image_loader = SheetImageLoader(sheet)
                
            except Exception as e:
                st.error(f"Error reading the Excel file. Make sure it's a valid .xlsx file. Error: {e}")
                st.stop()
            
            generated_cards = []
            
            for index, row in df.iterrows():
                name = str(row.get('Name', '')).strip()
                adm_no = str(row.get('Adm No', '')).split('.')[0].strip()

                if not name or name.lower() == "nan":
                    continue

                img = Image.open(TEMPLATE_FILE)
                draw = ImageDraw.Draw(img)
                
                excel_row = index + 2 
                cell_name = f"{PHOTO_COLUMN}{excel_row}" 
                try:
                    student_photo = image_loader.get(cell_name)
                    if student_photo.mode != 'RGB':
                        student_photo = student_photo.convert('RGB')
                    student_photo = student_photo.resize((PHOTO_WIDTH, PHOTO_HEIGHT))
                    img.paste(student_photo, (PHOTO_X, PHOTO_Y))
                except:
                    pass 

                father = str(row.get('Father Name', ''))
                student_class = str(row.get('Class', ''))
                dob = str(row.get('DOB', ''))
                address = str(row.get('Address', ''))
                mobile = str(row.get('Mobile', '')).split('.')[0] 

                text_color = (0, 0, 0) 
                draw.text((360, 555), name, font=font, fill=text_color)
                draw.text((360, 610), father, font=font, fill=text_color)
                draw.text((360, 665), student_class, font=font, fill=text_color)
                draw.text((360, 720), dob, font=font, fill=text_color)
                draw.text((360, 775), mobile, font=font, fill=text_color)
                draw.text((360, 825), address, font=font, fill=text_color)
                draw.text((70, 920), f"Adm. No.: {adm_no}", font=ImageFont.truetype("arialbd.ttf", 22), fill=text_color)

                generated_cards.append(img)
            
            if len(generated_cards) > 0:
                st.success(f"✅ Successfully processed {len(generated_cards)} students! Generating PDF...")
                
                PAGE_WIDTH, PAGE_HEIGHT = 2480, 3508
                COLS, ROWS = 3, 3
                CARD_WIDTH, GAP_X, GAP_Y = 700, 80, 100
                MARGIN_X = (PAGE_WIDTH - (COLS * CARD_WIDTH) - ((COLS - 1) * GAP_X)) // 2

                pages = []
                current_page = Image.new('RGB', (PAGE_WIDTH, PAGE_HEIGHT), 'white')
                card_count = 0

                for card in generated_cards:
                    ratio = card.height / card.width
                    card_height = int(CARD_WIDTH * ratio)
                    MARGIN_Y = (PAGE_HEIGHT - (ROWS * card_height) - ((ROWS - 1) * GAP_Y)) // 2

                    card = card.resize((CARD_WIDTH, card_height))
                    col = (card_count % (COLS * ROWS)) % COLS
                    row = (card_count % (COLS * ROWS)) // COLS
                    
                    x = MARGIN_X + col * (CARD_WIDTH + GAP_X)
                    y = MARGIN_Y + row * (card_height + GAP_Y)
                    
                    current_page.paste(card, (x, y))
                    card_count += 1
                    
                    if card_count % (COLS * ROWS) == 0:
                        pages.append(current_page)
                        current_page = Image.new('RGB', (PAGE_WIDTH, PAGE_HEIGHT), 'white')

                if card_count % (COLS * ROWS) != 0:
                    pages.append(current_page)

                pdf_buffer = io.BytesIO()
                pages[0].save(pdf_buffer, format='PDF', save_all=True, append_images=pages[1:])
                pdf_buffer.seek(0)
                
                st.balloons()
                st.download_button(
                    label="⬇️ Download Print-Ready PDF",
                    data=pdf_buffer,
                    file_name="Print_Ready_IDs.pdf",
                    mime="application/pdf",
                    type="primary",
                    use_container_width=True
                )
            else:
                st.warning("⚠️ No valid students found in the Excel file.")
