import os
import re
import base64
import pandas as pd
import requests
from collections import defaultdict

# ============================================================
# SETTINGS
# ============================================================

# ------------------------------------------------------------
# Local Excel file
# ------------------------------------------------------------

DATA_FILE = "Updating-image.xls"

# ------------------------------------------------------------
# GitHub settings
# ------------------------------------------------------------

GITHUB_OWNER = "rexbatongbakal"
GITHUB_REPO = "PRODUCT-DESCRIPTION"
GITHUB_BRANCH = "main"

# GitHub token is read from the environment
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN")

# ------------------------------------------------------------
# Excel columns
# ------------------------------------------------------------

# Excel Column B = 2
# Excel Column F = 6

HTML_PATH_COLUMN = 2
IMAGE_URL_COLUMN = 6

# ============================================================
# SETUP
# ============================================================

if not GITHUB_TOKEN:
    print()
    print("❌ GITHUB_TOKEN is not set.")
    print()
    print("Run:")
    print('export GITHUB_TOKEN="your_github_token"')
    print()
    exit(1)


# ============================================================
# FIND LOCAL DATA FILE
# ============================================================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_PATH = os.path.join(SCRIPT_DIR, DATA_FILE)


# ============================================================
# GITHUB API
# ============================================================

GITHUB_API = "https://api.github.com"

HEADERS = {
    "Authorization": f"Bearer {GITHUB_TOKEN}",
    "Accept": "application/vnd.github+json",
    "X-GitHub-Api-Version": "2022-11-28"
}


# ============================================================
# READ EXCEL
# ============================================================

def read_excel_data():

    products = defaultdict(list)

    print()
    print("Reading:", DATA_PATH)
    print()

    if not os.path.exists(DATA_PATH):

        print("❌ Spreadsheet not found:")
        print(DATA_PATH)

        return products

    try:

        df = pd.read_excel(
            DATA_PATH,
            header=None,
            engine="xlrd"
        )

    except Exception as e:

        print("❌ Unable to read Excel file:")
        print(e)

        return products

    print("Rows found:", len(df))
    print()

    for row_number, row in df.iterrows():

        # Make sure the row has enough columns
        if len(row) < IMAGE_URL_COLUMN:
            continue

        html_path = str(
            row.iloc[HTML_PATH_COLUMN - 1]
        ).strip()

        image_url = str(
            row.iloc[IMAGE_URL_COLUMN - 1]
        ).strip()

        # Skip blank / NaN values
        if html_path.lower() == "nan":
            continue

        if image_url.lower() == "nan":
            continue

        if not html_path or not image_url:
            continue

        # Skip header
        if html_path.lower() in [
            "html",
            "html path",
            "file",
            "path"
        ]:
            continue

        products[html_path].append(image_url)

    return products


# ============================================================
# GET FILE FROM GITHUB
# ============================================================

def get_github_file(file_path):

    url = (
        f"{GITHUB_API}/repos/"
        f"{GITHUB_OWNER}/"
        f"{GITHUB_REPO}/"
        f"contents/{file_path}"
    )

    params = {
        "ref": GITHUB_BRANCH
    }

    response = requests.get(
        url,
        headers=HEADERS,
        params=params
    )

    if response.status_code == 404:

        print("❌ FILE NOT FOUND ON GITHUB")
        print("Path:", file_path)

        return None

    if not response.ok:

        print("❌ GitHub error:")
        print(response.status_code)
        print(response.text)

        return None

    return response.json()


# ============================================================
# UPDATE FILE ON GITHUB
# ============================================================

def update_github_file(
    file_path,
    new_content,
    sha,
    commit_message
):

    url = (
        f"{GITHUB_API}/repos/"
        f"{GITHUB_OWNER}/"
        f"{GITHUB_REPO}/"
        f"contents/{file_path}"
    )

    encoded_content = base64.b64encode(
        new_content.encode("utf-8")
    ).decode("utf-8")

    payload = {
        "message": commit_message,
        "content": encoded_content,
        "sha": sha,
        "branch": GITHUB_BRANCH
    }

    response = requests.put(
        url,
        headers=HEADERS,
        json=payload
    )

    if response.status_code not in [200, 201]:

        print("❌ Failed to update GitHub file")
        print(response.status_code)
        print(response.text)

        return False

    return True


# ============================================================
# UPDATE HTML FILE
# ============================================================

def update_html_file(
    relative_html_path,
    image_urls
):

    print("=" * 70)
    print("HTML:", relative_html_path)
    print("Images found in spreadsheet:", len(image_urls))
    print()

    # --------------------------------------------------------
    # Get HTML from GitHub
    # --------------------------------------------------------

    github_file = get_github_file(
        relative_html_path
    )

    if not github_file:

        return False

    # --------------------------------------------------------
    # Decode GitHub content
    # --------------------------------------------------------

    try:

        encoded_content = github_file["content"]

        html = base64.b64decode(
            encoded_content
        ).decode("utf-8")

    except Exception as e:

        print("❌ Unable to decode GitHub file:")
        print(e)

        return False

    print(
        "GitHub file found."
    )

    # --------------------------------------------------------
    # FIND IMG SRC
    # --------------------------------------------------------

    img_pattern = re.compile(
        r'(<img\b[^>]*?\bsrc\s*=\s*)'
        r'(["\'])(.*?)(\2)',
        re.IGNORECASE
    )

    matches = list(
        img_pattern.finditer(html)
    )

    print(
        "Images found in HTML:",
        len(matches)
    )

    # --------------------------------------------------------
    # No images
    # --------------------------------------------------------

    if not matches:

        print(
            "❌ No <img src=\"...\"> found"
        )

        return False

    # --------------------------------------------------------
    # CHECK COUNTS
    # --------------------------------------------------------

    if len(image_urls) > len(matches):

        print()
        print("⚠️ WARNING:")
        print(
            f"Excel has {len(image_urls)} images "
            f"but HTML only has {len(matches)} "
            f"<img> tags."
        )

        print(
            "Only available HTML <img> tags "
            "will be updated."
        )

    elif len(image_urls) < len(matches):

        print()
        print("⚠️ WARNING:")
        print(
            f"Excel has {len(image_urls)} images "
            f"but HTML has {len(matches)} "
            f"<img> tags."
        )

        print(
            "Remaining HTML images will "
            "keep their existing URLs."
        )

    # --------------------------------------------------------
    # REPLACE IMAGE URLS
    # --------------------------------------------------------

    replacements = min(
        len(image_urls),
        len(matches)
    )

    changed = False

    # Bottom → top
    for index in range(
        replacements - 1,
        -1,
        -1
    ):

        match = matches[index]

        start = match.start(3)
        end = match.end(3)

        old_url = match.group(3)
        new_url = image_urls[index]

        if old_url == new_url:

            print()
            print(
                f"Image {index + 1}: "
                "Already correct"
            )

            continue

        html = (
            html[:start]
            + new_url
            + html[end:]
        )

        changed = True

        print()
        print(
            f"Image {index + 1}"
        )

        print(
            "OLD:",
            old_url
        )

        print(
            "NEW:",
            new_url
        )

    # --------------------------------------------------------
    # Nothing changed
    # --------------------------------------------------------

    if not changed:

        print()
        print(
            "ℹ️ No changes required."
        )

        return True

    # --------------------------------------------------------
    # Update GitHub
    # --------------------------------------------------------

    commit_message = (
        f"Update images for {relative_html_path}"
    )

    success = update_github_file(
        relative_html_path,
        html,
        github_file["sha"],
        commit_message
    )

    if success:

        print()
        print(
            "✅ Successfully updated on GitHub"
        )

        return True

    return False


# ============================================================
# MAIN
# ============================================================

def main():

    print()
    print("==============================================")
    print("      PRODUCT IMAGE URL UPDATER")
    print("==============================================")
    print()

    print(
        "GitHub repository:",
        f"{GITHUB_OWNER}/{GITHUB_REPO}"
    )

    print(
        "Branch:",
        GITHUB_BRANCH
    )

    print()

    # --------------------------------------------------------
    # Read Excel
    # --------------------------------------------------------

    products = read_excel_data()

    if not products:

        print()
        print(
            "❌ No valid product data found."
        )

        return

    print()
    print(
        "HTML files found:",
        len(products)
    )

    print()

    # --------------------------------------------------------
    # Process files
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # Summary
    # --------------------------------------------------------

    print()
    print("==============================================")
    print("                 COMPLETE")
    print("==============================================")
    print()

    print(
        "Successfully updated:",
        success
    )

    print(
        "Failed:",
        failed
    )

    print()


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":
    main()