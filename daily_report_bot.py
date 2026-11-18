"""
daily_report_bot.py

Gen AI Architect Program - Assignment 1
PyAutoGUI Automation

Indha program:
1. Chrome open pannum.
2. Public weather website-la irundhu temperature/weather information copy pannum.
3. Microsoft Excel open pannum.
4. New row-la current date/time, fetched weather data, short comment enter pannum.
5. Excel file-ai current date YYYY-MM-DD format-la save pannum.
6. Final Excel sheet-oda screenshot save pannum.

IMPORTANT:
- Program run aagumbodhu mouse/keyboard manually use panna koodadhu.
- Excel installed irukkanum.
- Chrome installed irukkanum.
"""

import os
import re
import sys
import time
import platform
import subprocess
from datetime import datetime
from pathlib import Path

import pyautogui
import pyperclip


# ============================================================
# BLOCK 1 - BASIC CONFIGURATION
# ============================================================
# Indha block-la automation-ku thevaiyana basic settings define panrom.
# City-ai change panna WEATHER_CITY value mattum change pannunga.
# Output files current project folder-kulla "output" directory-la save aagum.

WEATHER_CITY = "Chennai"
WEATHER_URL = f"https://wttr.in/{WEATHER_CITY}?format=3"

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"

TODAY = datetime.now().strftime("%Y-%m-%d")
REPORT_FILE = OUTPUT_DIR / f"daily_report_{TODAY}.xlsx"
SCREENSHOT_FILE = OUTPUT_DIR / f"daily_report_{TODAY}.png"

# PyAutoGUI fail-safe enable pannrom.
# Mouse-ai screen-oda top-left corner-ku kondu pona automation stop aagum.
pyautogui.FAILSAFE = True

# Automation smooth-a irukka small pause use panrom.
pyautogui.PAUSE = 0.3


# ============================================================
# BLOCK 2 - HELPER FUNCTION: WAIT
# ============================================================
# Apps/web pages open aaga konjam time edukkalam.
# Adhanala fixed sleep use pannrom.
# Slow machine-la timeout value increase pannalaam.

def wait(seconds=2):
    time.sleep(seconds)


# ============================================================
# BLOCK 3 - HELPER FUNCTION: OPEN CHROME
# ============================================================
# Indha function operating system based-a Chrome open pannum.
# Mac-la "open -a Google Chrome" command use panrom.
# Windows-la "start chrome" use panrom.
# Linux-la "google-chrome" try panrom.

def open_chrome():
    system = platform.system()

    print("[INFO] Chrome opening...")

    if system == "Darwin":
        subprocess.Popen(["open", "-a", "Google Chrome"])
    elif system == "Windows":
        subprocess.Popen(["cmd", "/c", "start", "", "chrome"])
    elif system == "Linux":
        subprocess.Popen(["google-chrome"])
    else:
        raise RuntimeError(f"Unsupported operating system: {system}")

    wait(4)


# ============================================================
# BLOCK 4 - HELPER FUNCTION: OPEN EXCEL
# ============================================================
# Indha function Microsoft Excel application-ai open pannum.
# Mac-la open command use pannrom.
# Windows-la start excel use pannrom.
# Linux-la Excel native-a available illadhadhala error message kudukrom.
#
# Assignment-la Microsoft Excel (or Mac-la Numbers) use panna sollirukku.
# Indha implementation Microsoft Excel-ai target pannudhu.

def open_excel():
    system = platform.system()

    print("[INFO] Microsoft Excel opening...")

    if system == "Darwin":
        subprocess.Popen(["open", "-a", "Microsoft Excel"])
    elif system == "Windows":
        subprocess.Popen(["cmd", "/c", "start", "", "excel"])
    elif system == "Linux":
        raise RuntimeError(
            "Linux detected. Assignment specifies Microsoft Excel or Numbers on Mac."
        )
    else:
        raise RuntimeError(f"Unsupported operating system: {system}")

    wait(6)


# ============================================================
# BLOCK 5 - FETCH WEATHER DATA USING BROWSER + PYAutoGUI
# ============================================================
# Idhu assignment-oda important part.
# Browser-la public website open pannrom.
# Page-la irukkura information-ai mouse/keyboard automation moolama
# select-all + copy pannrom.
#
# Website: wttr.in
# Example output:
# Chennai: 🌤️ +29°C
#
# Clipboard-la vandha text-ai Python read pannitu useful weather line
# extract pannrom.

def fetch_weather_from_browser():
    print("[INFO] Opening public weather website...")

    # Chrome active window-a irukkanum.
    pyautogui.hotkey("ctrl", "l")
    wait(0.5)

    # Public website URL type pannrom.
    pyautogui.write(WEATHER_URL, interval=0.01)
    pyautogui.press("enter")

    # Website load aagura vara wait pannrom.
    wait(5)

    print("[INFO] Copying information from browser page...")

    # Browser page content-ai select pannrom.
    pyautogui.hotkey("ctrl", "a")
    wait(0.5)

    # Selected page content clipboard-ku copy pannrom.
    pyautogui.hotkey("ctrl", "c")
    wait(1)

    # Clipboard content-ai Python-la retrieve pannrom.
    page_text = pyperclip.paste().strip()

    if not page_text:
        raise RuntimeError("Could not copy data from the public website.")

    print(f"[DEBUG] Browser copied data: {page_text[:200]}")

    # Expected weather format-la useful line search pannrom.
    # Example: Chennai: 🌤️ +29°C
    weather_match = re.search(
        rf"{re.escape(WEATHER_CITY)}:.*?(?:\+\d+°C|\-\d+°C).*",
        page_text,
        re.IGNORECASE
    )

    if weather_match:
        weather_data = weather_match.group(0).strip()
    else:
        # Website output slightly different-a irundhaal,
        # first non-empty line-ai fallback-a use pannrom.
        weather_data = page_text.splitlines()[0].strip()

    if not weather_data:
        raise RuntimeError("Weather information could not be extracted.")

    print(f"[INFO] Fetched data: {weather_data}")

    return weather_data


# ============================================================
# BLOCK 6 - CREATE SHORT COMMENT
# ============================================================
# Assignment requirement-la fetched data-ku own short comment
# add panna sollirukku.
#
# Temperature numeric value extract pannitu simple comment create panrom.
# Example:
# 30°C -> "Warm weather - stay hydrated."
# 22°C -> "Pleasant weather for outdoor activities."
# 18°C -> "Cool weather - carry a light jacket."

def create_comment(weather_data):
    # Celsius temperature-ai text-la irundhu extract panrom.
    temperature_match = re.search(r"([+-]?\d+(?:\.\d+)?)\s*°C", weather_data)

    if not temperature_match:
        return "Weather information captured successfully."

    temperature = float(temperature_match.group(1))

    if temperature >= 32:
        return "Hot weather - stay hydrated."
    elif temperature >= 27:
        return "Warm weather - suitable for normal activities."
    elif temperature >= 20:
        return "Pleasant weather for outdoor activities."
    else:
        return "Cool weather - carry a light jacket."


# ============================================================
# BLOCK 7 - PREPARE EXCEL DATA
# ============================================================
# Current date & time runtime-la generate panrom.
# User manually date/time type panna maatom.
#
# Three required values:
# 1. Today's date & time
# 2. Fetched data
# 3. Short comment

def prepare_report_data(weather_data):
    current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    comment = create_comment(weather_data)

    return current_datetime, weather_data, comment


# ============================================================
# BLOCK 8 - ENTER DATA INTO EXCEL USING PYAUTOGUI
# ============================================================
# Excel open pannitu blank workbook create panrom.
# First row-la column headings enter pannrom.
# Second row-la required data enter pannrom.
#
# Clipboard paste use pannuvadhala special characters / emoji irundhaalum
# typing vida reliable-a irukkum.

def enter_data_into_excel(current_datetime, weather_data, comment):
    print("[INFO] Creating Excel workbook...")

    # New blank workbook open panna Ctrl+N use pannrom.
    pyautogui.hotkey("ctrl", "n")
    wait(3)

    # Header row.
    headers = ["Date & Time", "Fetched Data", "Comment"]

    # Tab-separated values create pannrom.
    header_text = "\t".join(headers)
    data_text = "\t".join([current_datetime, weather_data, comment])

    # Header row-ai A1 cell-la paste pannrom.
    pyperclip.copy(header_text)
    pyautogui.hotkey("ctrl", "v")
    wait(1)

    # Next row A2-ku move pannrom.
    pyautogui.press("down")

    # Actual report data-ai A2 cell-la paste pannrom.
    pyperclip.copy(data_text)
    pyautogui.hotkey("ctrl", "v")
    wait(2)

    # Columns readable-a irukka all columns select pannitu
    # AutoFit Columns command use pannrom.
    pyautogui.hotkey("ctrl", "a")
    wait(0.5)

    # Excel ribbon shortcut:
    # Alt+H -> O -> I generally AutoFit Column Width.
    # Different Excel versions-la shortcut vary aagalam.
    pyautogui.press("alt")
    pyautogui.press("h")
    pyautogui.press("o")
    pyautogui.press("i")

    wait(2)


# ============================================================
# BLOCK 9 - SAVE EXCEL FILE
# ============================================================
# Assignment requirement:
# Filename-la current date YYYY-MM-DD format mandatory.
#
# Example:
# daily_report_2026-11-18.xlsx
#
# Ctrl+Shift+S use pannitu Save As dialog open pannrom.
# Full output path type pannitu save pannrom.

def save_excel_file():
    print(f"[INFO] Saving Excel file: {REPORT_FILE}")

    pyautogui.hotkey("ctrl", "shift", "s")
    wait(3)

    # Save As dialog-la file path enter panna:
    # Mac/Windows Excel behavior version-ku version vary aagalam.
    # File name field active illana Ctrl+L use pannrom.
    pyautogui.hotkey("ctrl", "l")
    wait(0.5)

    pyautogui.write(str(REPORT_FILE), interval=0.01)
    pyautogui.press("enter")

    wait(4)

    # Existing file overwrite confirmation vandhaal "Replace/Yes" select pannrom.
    # Usually Windows/Excel-la Alt+Y or Enter work aagum.
    pyautogui.press("enter")

    wait(3)

    print("[INFO] Excel file save operation completed.")


# ============================================================
# BLOCK 10 - TAKE SCREENSHOT OF FINAL EXCEL SHEET
# ============================================================
# Assignment requirement-la final Excel sheet-oda screenshot save panna
# sollirukku.
#
# Excel window foreground-la irukkumbodhu PyAutoGUI screenshot edukkum.
# Screenshot PNG format-la output folder-la save aagum.

def take_final_screenshot():
    print(f"[INFO] Taking screenshot: {SCREENSHOT_FILE}")

    # Small wait so Excel final state visible-a irukkum.
    wait(2)

    screenshot = pyautogui.screenshot()
    screenshot.save(SCREENSHOT_FILE)

    print("[INFO] Screenshot saved successfully.")


# ============================================================
# BLOCK 11 - MAIN AUTOMATION FLOW
# ============================================================
# Idhu complete bot execution flow.
# Requirements-la kudutha sequence same order-la implement pannirukkom:
#
# 1. Chrome open
# 2. Public website visit
# 3. Important data copy
# 4. Excel open
# 5. Date/time + fetched data + comment enter
# 6. Excel file save with today's date
# 7. Final sheet screenshot save

def main():
    print("=" * 60)
    print("Daily Report Bot - Starting")
    print("=" * 60)

    # Output directory create pannrom.
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    try:
        # Step 1 & 2: Browser open + public website data fetch.
        open_chrome()
        weather_data = fetch_weather_from_browser()

        # Step 3: Excel open.
        open_excel()

        # Step 4: Runtime date/time + comment prepare.
        current_datetime, fetched_data, comment = prepare_report_data(
            weather_data
        )

        # Step 5: Data Excel-la enter pannrom.
        enter_data_into_excel(
            current_datetime,
            fetched_data,
            comment
        )

        # Step 6: Excel file save pannrom.
        save_excel_file()

        # Step 7: Final Excel screenshot save pannrom.
        take_final_screenshot()

        print("=" * 60)
        print("Daily Report Bot - Completed Successfully")
        print(f"Excel     : {REPORT_FILE}")
        print(f"Screenshot: {SCREENSHOT_FILE}")
        print("=" * 60)

    except Exception as error:
        print("=" * 60)
        print("Daily Report Bot - FAILED")
        print(f"Error: {error}")
        print("=" * 60)

        # Error vandhaal useful debugging information print pannrom.
        print("Please check:")
        print("1. Chrome is installed and available.")
        print("2. Microsoft Excel is installed.")
        print("3. Internet connection is available.")
        print("4. Do not move mouse/keyboard while bot is running.")
        print("5. If Excel UI differs, adjust the save/autofit shortcuts.")

        sys.exit(1)


# ============================================================
# BLOCK 12 - PROGRAM ENTRY POINT
# ============================================================
# Indha standard Python entry point use pannuvadhala,
# file direct-a run pannumbodhu mattum main() execute aagum.

if __name__ == "__main__":
    main()
