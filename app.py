import streamlit as st
import pandas as pd
import openpyxl
from openpyxl_image_loader import SheetImageLoader
from PIL import Image, ImageDraw, ImageFont
import io

# --- PAGE SETUP ---
st.set_page_config(page_title="ID Card Generator", page_icon="🪪")
st.title("🪪 Automated ID Card Generator")
st.write("Upload your Excel file to generate a printable A4 PDF of student ID cards.")

# --- SETTINGS ---
TEMPLATE_FILE = 'template.jpg'
PHOTO_COLUMN = 'H'
PHOTO_X, PHOTO_Y = 248, 208       
PHOTO_WIDTH, PHOTO_HEIGHT = 245, 310   

# Load the font from the GitHub folder
try:
    font = ImageFont.truetype("arialbd.ttf", 26)
    font_small = ImageFont.truetype("arialbd.ttf", 22)
except Exception as e:
    st.error("Font file 'arialbd.ttf' is missing from GitHub! Please upload it.")
    st.stop()

# --- FILE UPLOADER ---
uploaded_file = st.file_uploader("Upload Students Excel File (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    if st.button("Generate ID Cards", type="primary"):
        with st.spinner("Processing data and extracting photos... Please wait."):
            try:
                # Read data using Pandas
                uploaded_file.seek(0) # Reset file pointer
                df = pd.read_excel(uploaded_file)
                df.columns = df.columns.str.strip()
                df = df.fillna("")
                
                # Read images using Openpyxl
                uploaded_file.seek(0) # Reset file pointer again
                wb = openpyxl.load_workbook(uploaded_file, data_only=True)
                sheet = wb.active
                image_loader = SheetImageLoader(sheet)
                
            except Exception as e:
                st.error(f"Error reading the Excel file. Make sure it's a valid .xlsx file. Error: {e}")
                st.stop()
            
            generated_cards = []
            
            # --- GENERATE CARDS ---
            for index, row in df.iterrows():
                name = str(row.get('Name', '')).strip()
                adm_no = str(row.get('Adm No', '')).strip()

                if not name or name.lower() == "nan":
                    continue

                img = Image.open(TEMPLATE_FILE)
                draw = ImageDraw.Draw(img)
                
                # Extract Photo
                excel_row = index + 2 
                cell_name = f"{PHOTO_COLUMN}{excel_row}" 
                try:
                    student_photo = image_loader.get(cell_name)
                    if student_photo.mode != 'RGB':
                        student_photo = student_photo.convert('RGB')
                    student_photo = student_photo.resize((PHOTO_WIDTH, PHOTO_HEIGHT))
                    img.paste(student_photo, (PHOTO_X, PHOTO_Y))
                except:
                    pass # Skip missing photos

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
                # draw.text((360, 160), f"Adm: {adm_no}", font=font_small, fill=(255,255,255))

                generated_cards.append(img)
            
            # --- CREATE PDF GRID ---
            if len(generated_cards) > 0:
                st.success(f"Successfully processed {len(generated_cards)} students! Generating PDF...")
                
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

                # Convert PDF to memory buffer for download
                pdf_buffer = io.BytesIO()
                pages[0].save(pdf_buffer, format='PDF', save_all=True, append_images=pages[1:])
                pdf_buffer.seek(0)
                
                st.balloons()
                st.download_button(
                    label="⬇️ Download Print-Ready PDF",
                    data=pdf_buffer,
                    file_name="Print_Ready_IDs.pdf",
                    mime="application/pdf",
                    type="primary"
                )
            else:
                st.warning("No valid students found in the Excel file.")
