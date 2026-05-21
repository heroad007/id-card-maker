import os
from PIL import Image

# --- CONFIGURATION ---
INPUT_FOLDER = 'Completed_IDs'
OUTPUT_FILE = 'Print_Ready_IDs.pdf'  # Will save as a multi-page PDF!

# Standard A4 Paper Size at 300 DPI (High Quality Print)
PAGE_WIDTH = 2480
PAGE_HEIGHT = 3508

# We will fit 3 columns and 3 rows of ID cards per page (9 cards per sheet)
COLS = 3
ROWS = 3
CARD_WIDTH = 700  # We scale the width to 700 pixels (standard ID width at 300dpi)
GAP_X = 80        # Horizontal space between cards
GAP_Y = 100       # Vertical space between cards

# Calculate margins to perfectly center the grid on the A4 page
MARGIN_X = (PAGE_WIDTH - (COLS * CARD_WIDTH) - ((COLS - 1) * GAP_X)) // 2

# --- READ GENERATED IDS ---
if not os.path.exists(INPUT_FOLDER):
    print(f"❌ Could not find folder '{INPUT_FOLDER}'.")
    exit()

# Get all the JPGs in the folder and sort them
files = [f for f in os.listdir(INPUT_FOLDER) if f.endswith('.jpg')]
files.sort()

if not files:
    print("❌ No ID cards found in the folder!")
    exit()

print(f"Found {len(files)} ID cards. Arranging them on A4 sheets...")

pages = []
current_page = Image.new('RGB', (PAGE_WIDTH, PAGE_HEIGHT), 'white')
card_count = 0

# --- BUILD THE GRID ---
for file in files:
    img_path = os.path.join(INPUT_FOLDER, file)
    card = Image.open(img_path)
    
    # Calculate new height to keep the card's exact proportions (no stretching!)
    ratio = card.height / card.width
    card_height = int(CARD_WIDTH * ratio)
    
    # Calculate Top Margin dynamically based on card height
    MARGIN_Y = (PAGE_HEIGHT - (ROWS * card_height) - ((ROWS - 1) * GAP_Y)) // 2

    # Resize card
    card = card.resize((CARD_WIDTH, card_height))
    
    # Calculate which column and row this card belongs to
    col = (card_count % (COLS * ROWS)) % COLS
    row = (card_count % (COLS * ROWS)) // COLS
    
    # Calculate exact X and Y coordinates to paste the card
    x = MARGIN_X + col * (CARD_WIDTH + GAP_X)
    y = MARGIN_Y + row * (card_height + GAP_Y)
    
    # Paste card onto the A4 page
    current_page.paste(card, (x, y))
    card_count += 1
    
    # If the page is full (9 cards), save it and start a fresh blank page
    if card_count % (COLS * ROWS) == 0:
        pages.append(current_page)
        current_page = Image.new('RGB', (PAGE_WIDTH, PAGE_HEIGHT), 'white')

# Add the final page if it has leftover cards on it
if card_count % (COLS * ROWS) != 0:
    pages.append(current_page)

# --- SAVE AS PDF ---
if pages:
    # This magic line saves all pages into one single PDF file!
    pages[0].save(OUTPUT_FILE, save_all=True, append_images=pages[1:])
    print(f"\n🎉 SUCCESS! Created '{OUTPUT_FILE}' with {len(pages)} pages.")
    print("You can now open this PDF and hit Print!")