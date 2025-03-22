# Card Price Scraper

This script automatically retrieves card descriptions and market prices from **TCGPlayer.com** using a provided Excel file containing links. It updates your Excel sheet with the scraped information and generates a helpful summary, including total card quantity, unique cards, total value, and price distribution.

---

## Step-by-Step Installation Guide

This guide will help you set up and run the script even if you've never used Python or any other programming language.

---

### Step 1: Install Python

1. **Download Python**
   - Visit the official Python website: [python.org](https://www.python.org/downloads/).
   - Download **Python 3.11** for Windows.

2. **Install Python**
   - Run the downloaded file.
   - **Important:** During installation, make sure to check the box that says:
     - ✅ **"Add Python to PATH"**.

3. **Verify Python Installation**
   - Open **Command Prompt** (search "cmd" in the Windows Start menu).
   - Type the following command and hit Enter:

     ```sh
     python --version
     ```
   - You should see a version like: `Python 3.11.x`.

---

### Step 2: Download and Set Up the Project

1. **Download Project Files**
   - Place the provided `app.py`, `requirements.txt`, and your Excel template (`original_data.xlsx`) in a clearly labeled project folder, e.g.:
     ```sh
     C:\User\*YOURUSERNAME*\Documents\CardPriceScraper\
     ```

2. **Make a Copy of the Excel File**
   - Right-click your original Excel file, in the data folder, (`original_data.xlsx`) and select **Copy**, then right-click again in the same folder and select **Paste**.
   - Rename the copied file, in the data folder, (e.g., `card_prices.xlsx`). This ensures your original file stays safe.

---

### Step 3: Set Up a Virtual Environment

A virtual environment keeps your project isolated and prevents conflicts between different Python programs.

1. **Open Command Prompt in Your Project Folder**
   - Navigate to your project folder in File Explorer.
   - Hold `Shift`, right-click the folder, and select:

     ```sh
     Open PowerShell window here
     ```

2. **Create and Activate the Virtual Environment**

   **Create:**

   ```sh
   python -m venv venv
   ```

   **Activate:**

   ```sh
   venv\Scripts\activate
   ```

   - You should now see `(venv)` at the beginning of your command line, indicating the environment is active.

3. **Install Dependencies**

   Run this command in the activated environment:

   ```sh
   pip install -r requirements.txt
   ```

---

### Step 4: Set Up ChromeDriver

The script uses **Selenium**, which needs ChromeDriver to automate Chrome browser.

1. **Check Your Chrome Browser Version**
   - Open Chrome browser, type `chrome://settings/help` in the address bar.
   - Note your Chrome version (e.g., **122.0.6261.95**).

2. **Download ChromeDriver**
   - Go to [ChromeDriver](https://chromedriver.chromium.org/downloads).
   - Choose and download the version that exactly matches your Chrome version.

3. **Install ChromeDriver**
   - Extract the downloaded `chromedriver.exe` file.
   - Move `chromedriver.exe` to your project folder (where `app.py` is located).

---

### Step 5: Prepare Your Excel File

- Open the copied Excel file, in the data folder, (`card_prices.xlsx`).
- Navigate to the **"Cards"** worksheet.
- Paste the **TCGPlayer.com** URLs for your cards in the **"Link"** column.
- Enter the quantity of each card in the **"Qty."** column.
- Save and close the Excel file.
- Ensure that you have the entire file path to the `card_prices.xlsx` set as your `excel_file` in app.py, on line 11.
  ```python
  excel_file = r"C:\User\*YOURUSERNAME*\Documents\CardPriceScraper\data\card_prices.xlsx"
  ```

---

## Running the Script

1. **Ensure the Virtual Environment is Active**
   - If not activated already, open Command Prompt in your project folder and type:

   ```sh
   venv\Scripts\activate
   ```

2. **Run the Script**

   ```sh
   python app.py
   ```

3. **Wait for Completion**
   - The script will:
     - Open the Chrome browser in the background.
     - Scrape descriptions and prices.
     - Update your Excel file.
     - Generate a summary.

   Once finished, you'll see:

   ```
   Excel file updated successfully.
   ```

4. **Review Results**
   - Open your Excel file (`card_prices.xlsx`) to see updated card descriptions, prices, and the summary.

---

## Troubleshooting

- **ChromeDriver errors:**
  - Double-check that your ChromeDriver version matches your Chrome version.
  - Ensure no Chrome instances are running when you start the script.

- **Excel file issues:**
  - Make sure your Excel file is closed before running the script.

- **Debugging:**
  - If issues arise, enable debugging mode by changing the `debug` variable in `app.py`:

  ```python
  debug = True
  ```

- This will show more detailed logs that can help identify issues.

---

## License

This project is designed for personal use and educational purposes. Feel free to modify and learn from it.

