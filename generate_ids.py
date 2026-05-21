import pandas as pd
import openpyxl
from openpyxl_image_loader import SheetImageLoader
from PIL import Image, ImageDraw, ImageFont
import os
import sys

# --- CONFIGURATION ---
EXCEL_FILE = 'students.xlsx'      
TEMPLATE_FILE = 'template.jpg'  
OUTPUT_FOLDER = 'Completed_IDs'   
PHOTO_COLUMN = 'H'  # <--- Change this if your photos are in a different column!

# --- PHOTO BOX SETTINGS ---
PHOTO_X = 248        # Left/Right position of the photo
PHOTO_Y = 208        # Up/Down position of the photo
PHOTO_WIDTH = 245    # Width of the photo box
PHOTO_HEIGHT = 310   # Height of the photo box


FONT_SIZE = 26
try:
    font = ImageFont.truetype("arialbd.ttf", FONT_SIZE) 
except IOError:
    font = ImageFont.load_default()

if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

# --- READ DATA & IMAGES FROM EXCEL ---
try:
    print("Loading Excel file (this might take a few seconds if it has many photos)...")
    # Pandas reads the text
    df = pd.read_excel(EXCEL_FILE)
    df.columns = df.columns.str.strip()
    df = df.fillna("") 
    
    # Openpyxl reads the images
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    sheet = wb.active
    image_loader = SheetImageLoader(sheet)
    
except Exception as e:
    print(f"❌ ERROR: Could not open {EXCEL_FILE}. Is it open in another program?")
    print(e)
    sys.exit()

print(f"\n✅ Data loaded. Generating ID cards with embedded photos...")

# --- GENERATE ID CARDS ---
generated_count = 0

for index, row in df.iterrows():
    name = str(row.get('Name', '')).strip()
    adm_no = str(row.get('Adm No', '')).strip()

    if not name or name.lower() == "nan":
        continue

    img = Image.open(TEMPLATE_FILE)
    draw = ImageDraw.Draw(img)
    
    # ---------------------------------------------------------
    # EXTRACT PHOTO FROM EXCEL CELL
    # Pandas rows start at 0, but Excel rows start at 1. 
    # Because of the Header row, row index 0 is actually Excel Row 2.
    # ---------------------------------------------------------
    excel_row = index + 2 
    cell_name = f"{PHOTO_COLUMN}{excel_row}" # e.g., "H2"
    
    try:
        # Try to get the image from that specific cell
        student_photo = image_loader.get(cell_name)
        
        # Make sure it's in a format PIL can use, then resize and paste
        if student_photo.mode != 'RGB':
            student_photo = student_photo.convert('RGB')
            
        student_photo = student_photo.resize((PHOTO_WIDTH, PHOTO_HEIGHT))
        img.paste(student_photo, (PHOTO_X, PHOTO_Y))
        
    except ValueError:
        # ValueError means image_loader couldn't find an image in that cell
        print(f"⚠️ Missing Photo in cell {cell_name} for {name}. Leaving box blank.")
    except Exception as e:
         print(f"⚠️ Error loading photo for {name} in cell {cell_name}: {e}")
    # ---------------------------------------------------------

    father = str(row.get('Father Name', ''))
    student_class = str(row.get('Class', ''))
    dob = str(row.get('DOB', ''))
    address = str(row.get('Address', ''))
    mobile = str(row.get('Mobile', '')).split('.')[0] 

    text_color = (0, 0, 0) 

    # --- TEXT COORDINATES ---
    # --- THE CORRECTED COORDINATES ---
    # First number is X (Left/Right). Second number is Y (Up/Down).
    
    # 1. Name
    draw.text((360, 555), name, font=font, fill=text_color)
    
    # 2. Father's Name 
    draw.text((360, 610), father, font=font, fill=text_color)
    
    # 3. Class
    draw.text((360, 665), student_class, font=font, fill=text_color)
    
    # 4. Date of Birth
    draw.text((360, 720), dob, font=font, fill=text_color)
    
    # 5. Contact Number
    draw.text((360, 775), mobile, font=font, fill=text_color)
    
    # 6. Address
    draw.text((360, 825), address, font=font, fill=text_color)

    # # Adm No (Moved slightly left so it doesn't fall off the red edge)
    # draw.text((370, 160), f"Adm: {adm_no}", font=ImageFont.truetype("arialbd.ttf", 22), fill=(255,255,255))

    # --- SAVE ---
    safe_name = "".join([c for c in name if c.isalpha() or c.isdigit() or c==' ']).rstrip()
    output_filename = f"{OUTPUT_FOLDER}/ID_{adm_no}_{safe_name}.jpg"
    img.save(output_filename)
    
    print(f"✔️ Created: {output_filename}")
    generated_count += 1

print(f"\n🎉 Done! Successfully generated {generated_count} ID cards.")