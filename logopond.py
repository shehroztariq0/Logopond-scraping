import os
import time
import re
import requests
import traceback
from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup

# Setup Chrome WebDriver
options = Options()
options.add_argument("--headless")  # Run in headless mode
options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options  # ✅ Use this
)


driver.maximize_window()

# Start URL
start_url = "https://logopond.com/gallery/list/?gallery=featured&filter="
driver.get(start_url)
time.sleep(3)

# Output folder and Excel workbook
os.makedirs("images", exist_ok=True)
wb = Workbook()
ws = wb.active
# Save data to Excel sheet and save file immediately after appending
ws.append(["title", "original_img_name", "description", "tags_str"])

# Save the workbook after each entry to avoid data loss
try:
    wb.save("titles.xlsx")
except Exception as e:
    print(f"❌ Error saving row to Excel: {e}")

# Sanitize filenames to remove invalid chars
def shorten_filename(filename):
    return re.sub(r'[\\/*?:"<>|]', "", filename)

def scrape_logos():
    logo_items = driver.find_elements(By.CLASS_NAME, "logo_item")
    print(f"🖼️ Found {len(logo_items)} logos on the page.")

    for item in logo_items:
        try:
            title_elem = item.find_element(By.CLASS_NAME, "logo_title")
            title = title_elem.get_attribute("title").strip()
            detail_url = title_elem.get_attribute("href")

            img_elem = item.find_element(By.CLASS_NAME, "theimg")
            img_src = img_elem.get_attribute("src")
            if not img_src.startswith("http"):
                img_src = "https://logopond.com" + img_src

            original_img_name = shorten_filename(os.path.basename(img_src))
            img_path = os.path.join("images", original_img_name)

            # Download image if not already downloaded
            if not os.path.exists(img_path):
                response = requests.get(img_src, headers={"User-Agent": "Mozilla/5.0"})
                if response.status_code == 200:
                    with open(img_path, "wb") as f:
                        f.write(response.content)
                    print(f"✅ Downloaded image: {original_img_name}")
                else:
                    print(f"❌ Failed to download image ({response.status_code}): {original_img_name}")
            else:
                print(f"⚠️ Image already exists: {original_img_name}")

            # Open detail page in new tab
            driver.execute_script("window.open(arguments[0]);", detail_url)
            driver.switch_to.window(driver.window_handles[1])

            description = "No description found"
            tags_str = ""

            try:
                # Wait for .hook element on detail page
                hook_elem = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "hook"))
                )
                soup = BeautifulSoup(hook_elem.get_attribute("innerHTML"), "html.parser")

                # Extract description
                desc_tag = soup.find("strong", string=re.compile("Description", re.I))
                if desc_tag:
                    description_parts = []
                    for sib in desc_tag.next_siblings:
                        if getattr(sib, 'name', None) in ["strong", "br"]:
                            break
                        if isinstance(sib, str):
                            description_parts.append(sib.strip())
                        else:
                            description_parts.append(sib.get_text(" ", strip=True))
                    description = " ".join(description_parts).strip()

                # Extract tags
                tag_section = soup.find("strong", string=re.compile("Tags", re.I))
                tags = []
                if tag_section:
                    for sib in tag_section.next_siblings:
                        if getattr(sib, "name", None) == "a":
                            tags.append(sib.get_text(strip=True))
                        elif getattr(sib, "name", None) in ["strong", "br"]:
                            break
                tags_str = ", ".join(tags)

            except TimeoutException:
                print("⚠️ Timeout waiting for detail page content, skipping description and tags.")
            except Exception as e:
                print(f"⚠️ Error parsing detail page: {e}")
                traceback.print_exc()

            finally:
                # Close detail tab and switch back
                driver.close()
                driver.switch_to.window(driver.window_handles[0])

            # Save data to Excel sheet
            ws.append([title, original_img_name, description, tags_str])
            wb.save(os.path.join(os.getcwd(), "titles.xlsx"))  # Save incrementally
            print(f"💾 Saved entry for: {title}")

        except Exception as e:
            print(f"❗ Error scraping a logo: {e}")
            traceback.print_exc()

def click_more_until_end():
    while True:
        try:
            more_btn = driver.find_element(By.CSS_SELECTOR,
                "a.button.large-2.medium-2.tween-3.small-4.large-centered.medium-centered.tween-centered.small-centered")
            if "disabled" in more_btn.get_attribute("class"):
                print("🚫 'More' button is disabled.")
                break
            driver.execute_script("arguments[0].scrollIntoView(true);", more_btn)
            time.sleep(1)
            more_btn.click()
            time.sleep(4)
        except NoSuchElementException:
            print("ℹ️ No 'More' button found.")
            break
        except ElementClickInterceptedException:
            print("↪️ Click intercepted, retrying...")
            time.sleep(2)
        except Exception as e:
            print(f"❗ Error clicking 'More': {e}")
            break

def go_to_next_page():
    try:
        wait = WebDriverWait(driver, 10)
        next_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
            "a.button.large-4.medium-4.tween-4.small-4.lleft.mleft.tleft.sleft")))
        
        print("➡️ Navigating to next page...")
        driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
        time.sleep(1)
        
        try:
            next_btn.click()
        except ElementClickInterceptedException:
            print("⚠️ Click intercepted, using JavaScript click.")
            driver.execute_script("arguments[0].click();", next_btn)

        time.sleep(4)
        return True

    except NoSuchElementException:
        print("✅ Reached last page.")
        return False
    except TimeoutException:
        print("⏳ Timeout waiting for the next button.")
        return False
    except Exception as e:
        print(f"❗ Error on next page: {e}")
        return False


# Main scraping loop
while True:
    click_more_until_end()
    scrape_logos()
    if not go_to_next_page():
        break

# Save results
output_file = os.path.join(os.getcwd(), "titles.xlsx")
try:
    wb.save(output_file)
    print(f"🎉 Done. Images saved in 'images/', data saved in '{output_file}'")
except Exception as e:
    print(f"❌ Error saving Excel file: {e}")
