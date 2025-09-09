# Author: Duane Rinehart, MBA (drinehar@sdccd.edu)
# 9-SEP-2025
# downloads embedded Google docs presentation slides as PNG files

# uses headless browser to emulate key presses to navigate slides, then grabs screenshots

# INSTRUCTIONS:
# create and activate a Python virtual environment
# pip install {modules}
# playwright install

# Just add doc_id below and run the script

import asyncio
from playwright.async_api import async_playwright
import os
import hashlib
from PIL import Image


ddoc_id = '2PACX-1vRRVYnc-UTR3F160I-SDazUCkTHwhnALqnIpokRmS2FTkSK8dAISByIaP00a0Bekf6Jr9Ih3ovsRbOy'

embed_url = f"https://docs.google.com/presentation/d/e/{ddoc_id}/pub?start=false&loop=false&delayms=3000"
dest_folder = "slides"

async def save_slides():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        await page.goto(embed_url)

        os.makedirs(dest_folder, exist_ok=True)

        prev_hash = None
        slide_num = 1
        slide_paths = []  # keep track of saved files

        while True:
            await page.wait_for_timeout(1000)

            # Take screenshot into memory
            buf = await page.screenshot(full_page=True)

            # Hash screenshot to detect duplicates
            h = hashlib.md5(buf).hexdigest()
            if h == prev_hash:
                print("No new slide detected → stopping.")
                break

            # Save to file
            path = f"{dest_folder}/slide{slide_num:02d}.png"
            with open(path, "wb") as f:
                f.write(buf)
            print("Saved", path)

            slide_paths.append(path)
            prev_hash = h
            slide_num += 1

            # Advance to next slide
            await page.keyboard.press("ArrowRight")

        await browser.close()

        # === Convert PNGs to a single PDF ===
        if slide_paths:
            images = [Image.open(p).convert("RGB") for p in slide_paths]
            pdf_path = f"{dest_folder}/slides.pdf"
            images[0].save(pdf_path, save_all=True, append_images=images[1:])
            print("PDF saved as", pdf_path)

asyncio.run(save_slides())