import csv
import re
import shutil
from pathlib import Path
from collections import defaultdict

# ============================================================
# SETTINGS
# ============================================================

# Excel/CSV file
DATA_FILE = "Updating-image.xls"

# Column numbers
# Excel Column B = 2
# Excel Column E = 5
HTML_PATH_COLUMN = 2
IMAGE_URL_COLUMN = 5

# Create backups before changing HTML
CREATE_BACKUP = True


# ============================================================
# FIND PROJECT ROOT
# ============================================================

PROJECT_ROOT = Path(__file__).resolve().parent
DATA_PATH = PROJECT_ROOT / DATA_FILE


# ============================================================
# READ CSV
# ============================================================

def read_csv_data():

    products = defaultdict(list)

    print()
    print("Reading:", DATA_PATH)
    print()

    with open(DATA_PATH, "r", encoding="utf-8-sig", newline="") as file:

        reader = csv.reader(file)

        for row_number, row in enumerate(reader, start=1):

            # Skip rows that don't have enough columns
            if len(row) < IMAGE_URL_COLUMN:
                continue

            html_path = row[HTML_PATH_COLUMN - 1].strip()
            image_url = row[IMAGE_URL_COLUMN - 1].strip()

            # Skip blank rows
            if not html_path or not image_url:
                continue

            # Skip header if necessary
            if html_path.lower() in ["html", "html path", "file", "path"]:
                continue

            products[html_path].append(image_url)

    return products


# ============================================================
# UPDATE HTML IMAGES
# ============================================================

def update_html_file(relative_html_path, image_urls):

    html_file = PROJECT_ROOT / relative_html_path

    print("=" * 70)
    print("HTML:", relative_html_path)
    print("Images found in spreadsheet:", len(image_urls))

    # Check if HTML exists
    if not html_file.exists():

        print("❌ HTML FILE NOT FOUND")
        print("Expected:", html_file)

        return False

    # --------------------------------------------------------
    # BACKUP
    # --------------------------------------------------------

    if CREATE_BACKUP:

        backup_file = html_file.with_suffix(
            html_file.suffix + ".backup"
        )

        shutil.copy2(html_file, backup_file)

        print("Backup:", backup_file.name)

    # --------------------------------------------------------
    # READ HTML
    # --------------------------------------------------------

    html = html_file.read_text(encoding="utf-8")

    # --------------------------------------------------------
    # FIND IMG SRC
    # --------------------------------------------------------

    img_pattern = re.compile(
        r'(<img\b[^>]*?\bsrc\s*=\s*)(["\'])(.*?)(\2)',
        re.IGNORECASE
    )

    matches = list(img_pattern.finditer(html))

    print("Images found in HTML:", len(matches))

    # No images
    if not matches:

        print("❌ No <img src=\"...\"> found")

        return False

    # --------------------------------------------------------
    # CHECK IMAGE COUNTS
    # --------------------------------------------------------

    if len(image_urls) > len(matches):

        print()
        print("⚠️ WARNING:")
        print(
            f"Excel has {len(image_urls)} images "
            f"but HTML only has {len(matches)} <img> tags."
        )

        print(
            "Only the available HTML <img> tags will be updated."
        )

    elif len(image_urls) < len(matches):

        print()
        print("⚠️ WARNING:")
        print(
            f"Excel has {len(image_urls)} images "
            f"but HTML has {len(matches)} <img> tags."
        )

        print(
            "The remaining HTML images will keep their existing URLs."
        )

    # --------------------------------------------------------
    # REPLACE IMAGE URLS
    # --------------------------------------------------------

    replacements = min(len(image_urls), len(matches))

    # Replace from bottom to top so character positions remain valid
    for index in range(replacements - 1, -1, -1):

        match = matches[index]

        start = match.start(3)
        end = match.end(3)

        old_url = match.group(3)
        new_url = image_urls[index]

        html = html[:start] + new_url + html[end:]

        print()
        print(f"Image {index + 1}")
        print("OLD:", old_url)
        print("NEW:", new_url)

    # --------------------------------------------------------
    # SAVE
    # --------------------------------------------------------

    html_file.write_text(html, encoding="utf-8")

    print()
    print("✅ Updated:", html_file)

    return True


# ============================================================
# MAIN
# ============================================================

def main():

    print()
    print("==============================================")
    print("   PRODUCT IMAGE URL UPDATER")
    print("==============================================")

    if not DATA_PATH.exists():

        print()
        print("❌ Spreadsheet not found:")
        print(DATA_PATH)
        print()

        return

    products = read_csv_data()

    if not products:

        print()
        print("❌ No valid product data found.")
        print()

        return

    print()
    print("HTML files found:", len(products))
    print()

    success = 0
    failed = 0

    for html_path, image_urls in products.items():

        result = update_html_file(
            html_path,
            image_urls
        )

        if result:
            success += 1
        else:
            failed += 1

    print()
    print("==============================================")
    print("              COMPLETE")
    print("==============================================")
    print()
    print("Successfully updated:", success)
    print("Failed:", failed)
    print()


if __name__ == "__main__":
    main()