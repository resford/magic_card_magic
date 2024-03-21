import os
import requests
from config import openai_api_key, google_creds_file, spreadsheet_id, folder_id
from PIL import Image
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from openai import OpenAI
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException, TimeoutException, ElementClickInterceptedException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time

class MTGCardArtCreator:
    def __init__(self, openai_api_key, google_creds_file, spreadsheet_id, folder_id):
        self.client = OpenAI(api_key=openai_api_key)
        self.spreadsheet_id = spreadsheet_id
        self.folder_id = folder_id
        self.creds = service_account.Credentials.from_service_account_file(
            google_creds_file,
            scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        )
        self.sheet_service = build('sheets', 'v4', credentials=self.creds)
        self.drive_service = build('drive', 'v3', credentials=self.creds)

    def adjust_image_size(self, image_path):
        with Image.open(image_path) as img:
            size = (900, 900)
            img.thumbnail(size, Image.Resampling.LANCZOS)
            img.save(image_path, optimize=True)

    def generate_and_upload_images(self):
        range_name = 'Cards'  # Adjust as needed
        result = self.sheet_service.spreadsheets().values().get(spreadsheetId=self.spreadsheet_id, range=range_name).execute()
        values = result.get('values', [])
        created_cards_info = []

        for row in values[315:375]:  # Assuming the first row contains headers
            if len(row) >= 10 and row[2] is not None and row[2] != "":  # Ensure row has enough data and prompt is not null
                filename, prompt = row[1], row[2]
                try:
                    response = self.client.images.generate(model="dall-e-3", prompt=prompt, size="1024x1024", quality="standard", n=1)
                    image_url = response.data[0].url
                    image_response = requests.get(image_url)
                    if image_response.status_code == 200:
                        image_path = f'./images/{filename}.jpg'
                        os.makedirs(os.path.dirname(image_path), exist_ok=True)
                        with open(image_path, 'wb') as file:
                            file.write(image_response.content)

                        if os.path.getsize(image_path) > 2 * 1024 * 1024:
                            self.adjust_image_size(image_path)

                        file_metadata = {'name': f'{filename}.jpg', 'mimeType': 'image/jpeg', 'parents': [self.folder_id]}
                        media = MediaFileUpload(image_path, mimetype='image/jpeg')
                        uploaded_file = self.drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                        print(f"Uploaded {filename}.jpg with ID: {uploaded_file.get('id')}")

                    card_info = {
                        'image_path': f'./images/{filename}.jpg',
                        'card_title': row[1],
                        'other_text_fields': {
                            'card_type': row[3],
                            'subtype': row[5],
                            'mana_value': row[10],
                            'power': row[11],
                            'toughness': row[12],
                            'static_abilities': row[13],
                            'triggered_abilities_1': row[14],
                            'triggered_abilities_2': row[15],
                            'triggered_abilities_3': row[16],
                            'triggered_abilities_4': row[17],
                            'rarity': row[18],
                            'flavor': row[20]
                        }
                    }
                    created_cards_info.append(card_info)
                except Exception as e:
                    print(f"Error processing '{filename}': {e}")

        return created_cards_info

class MTGCardCreator:
    def __init__(self, image_path, card_title, other_text_fields=None):
        self.image_path = image_path
        self.card_title = card_title
        self.other_text_fields = other_text_fields if other_text_fields else {}
        self.driver = self.init_driver()
        self.is_logged_in = False  # Flag to track login status

    def init_driver(self):
        options = Options()
        options.add_experimental_option('excludeSwitches', ['enable-logging'])  # Suppress logging
        return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    def login(self):
        if self.is_logged_in:  # Check if already logged in
            return  # Skip login if already logged in

        username = "mckonkie_sloughswater"
        password = "grand2024"

        self.driver.get("https://mtgcardsmith.com/login")
        time.sleep(2)
        username_input = self.driver.find_element(By.ID, "username")
        username_input.send_keys(username)

        password_input = self.driver.find_element(By.ID, "password")
        password_input.send_keys(password)

        login_button = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='submit'][value='Login']"))
        )
        login_button.click()
        self.is_logged_in = True  # Update flag after successful login

    def navigate_to_page(self, url="https://mtgcardsmith.com/mtg-card-maker/"):
        self.driver.get(url)
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file']"))
            )
        except TimeoutException:
            self.driver.execute_script("window.stop();")

    def upload_image_and_confirm(self):
        absolute_image_path = os.path.abspath(self.image_path)
        file_input = self.driver.find_element(By.CSS_SELECTOR, "input[type='file']")
        file_input.send_keys(absolute_image_path)
        confirm_button = WebDriverWait(self.driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".slim-editor-btn.slim-btn-confirm"))
        )
        confirm_button.click()
        time.sleep(2)

    def finalize_card_creation(self):
        next_button = WebDriverWait(self.driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='submit'][value='Next']"))
        )
        next_button.click()

    def enter_card_title_and_other_fields(self):
        self.wait_and_send_keys("//input[@name='name'][@placeholder='Card Title (Required)']", self.card_title)

        if 'mana_value' in self.other_text_fields:
            self.wait_and_send_keys("//input[@name='custom_mana'][@placeholder='Select your casting cost...']", self.other_text_fields['mana_value'])

        if 'card_type' in self.other_text_fields:
            print(f"Trying to select card_type")
            self.select_custom_dropdown_option('s2id_autogen1', self.other_text_fields['card_type'])
            print(f"Done selecting creature_type")

        ability_texts = []

        if 'static_abilities' in self.other_text_fields:
            static_ability_text = self.other_text_fields['static_abilities']
            if static_ability_text:
                # Split triggered ability text into chunks of 45 characters
                chunks = [static_ability_text[i:i+45] for i in range(0, len(static_ability_text), 45)]
                formatted_chunks = []

                # Iterate over each chunk
                for chunk in chunks:
                    # Find the index of the first space after character 45
                    split_index = chunk.find(' ')

                    # If a space is found after character 45, split the text at that index
                    if split_index != -1:
                        formatted_chunks.append(chunk[:split_index])
                        formatted_chunks.append(chunk[split_index+1:])
                    else:
                        # If no space is found after character 45, split at character 45
                        formatted_chunks.append(chunk)

                # Join the chunks with line breaks
                formatted_static_ability_text = '\n'.join(formatted_chunks)
                ability_texts.append(formatted_static_ability_text)

        if 'triggered_abilities_1' in self.other_text_fields:
            triggered_ability_1_text = self.other_text_fields['triggered_abilities_1']
            if triggered_ability_1_text:
                # Split triggered ability text into chunks of 45 characters
                chunks = [triggered_ability_1_text[i:i+45] for i in range(0, len(triggered_ability_1_text), 45)]
                formatted_chunks = []

                # Iterate over each chunk
                for chunk in chunks:
                    # Find the index of the first space after character 45
                    split_index = chunk.find(' ')

                    # If a space is found after character 45, split the text at that index
                    if split_index != -1:
                        formatted_chunks.append(chunk[:split_index])
                        formatted_chunks.append(chunk[split_index+1:])
                    else:
                        # If no space is found after character 45, split at character 45
                        formatted_chunks.append(chunk)

                # Join the chunks with line breaks
                formatted_triggered_ability_1_text = '\n'.join(formatted_chunks)
                ability_texts.append(formatted_triggered_ability_1_text)

        if 'triggered_abilities_2' in self.other_text_fields:
            triggered_ability_text = self.other_text_fields['triggered_abilities_2']
            if triggered_ability_text:
                # Split triggered ability text into chunks of 45 characters
                chunks = [triggered_ability_text[i:i+45] for i in range(0, len(triggered_ability_text), 45)]
                formatted_chunks = []

                # Iterate over each chunk
                for chunk in chunks:
                    # Find the index of the first space after character 45
                    split_index = chunk.find(' ')

                    # If a space is found after character 45, split the text at that index
                    if split_index != -1:
                        formatted_chunks.append(chunk[:split_index])
                        formatted_chunks.append(chunk[split_index+1:])
                    else:
                        # If no space is found after character 45, split at character 45
                        formatted_chunks.append(chunk)

                # Join the chunks with line breaks
                formatted_triggered_ability_text = '\n'.join(formatted_chunks)
                ability_texts.append(formatted_triggered_ability_text)

        if 'triggered_abilities_3' in self.other_text_fields:
            triggered_ability_text = self.other_text_fields['triggered_abilities_3']
            if triggered_ability_text:
                # Split triggered ability text into chunks of 45 characters
                chunks = [triggered_ability_text[i:i+45] for i in range(0, len(triggered_ability_text), 45)]
                formatted_chunks = []

                # Iterate over each chunk
                for chunk in chunks:
                    # Find the index of the first space after character 45
                    split_index = chunk.find(' ')

                    # If a space is found after character 45, split the text at that index
                    if split_index != -1:
                        formatted_chunks.append(chunk[:split_index])
                        formatted_chunks.append(chunk[split_index+1:])
                    else:
                        # If no space is found after character 45, split at character 45
                        formatted_chunks.append(chunk)

                # Join the chunks with line breaks
                formatted_triggered_ability_text = '\n'.join(formatted_chunks)
                ability_texts.append(formatted_triggered_ability_text)

        if 'triggered_abilities_4' in self.other_text_fields:
            triggered_ability_text = self.other_text_fields['triggered_abilities_4']
            if triggered_ability_text:
                # Split triggered ability text into chunks of 45 characters
                chunks = [triggered_ability_text[i:i+45] for i in range(0, len(triggered_ability_text), 45)]
                formatted_chunks = []

                # Iterate over each chunk
                for chunk in chunks:
                    # Find the index of the first space after character 45
                    split_index = chunk.find(' ')

                    # If a space is found after character 45, split the text at that index
                    if split_index != -1:
                        formatted_chunks.append(chunk[:split_index])
                        formatted_chunks.append(chunk[split_index+1:])
                    else:
                        # If no space is found after character 45, split at character 45
                        formatted_chunks.append(chunk)

                # Join the chunks with line breaks
                formatted_triggered_ability_text = '\n'.join(formatted_chunks)
                ability_texts.append(formatted_triggered_ability_text)

        if 'oracle' in self.other_text_fields:
            oracle_text = self.other_text_fields['oracle']
            if oracle_text:
                # Split oracle text into chunks of 40 characters and join with line breaks
                oracle_chunks = [oracle_text[i:i+40] for i in range(0, len(oracle_text), 40)]
                formatted_oracle_text = '\n'.join(oracle_chunks)
                ability_texts.append(formatted_oracle_text)

        if 'flavor' in self.other_text_fields:
            flavor_text = self.other_text_fields['flavor']
            if flavor_text:
                # Add an initial line break
                ability_texts.append('\n')

                # Split flavor text into chunks of 40 characters and join with line breaks
                flavor_chunks = [flavor_text[i:i+40] for i in range(0, len(flavor_text), 40)]
                formatted_flavor_text = '\n'.join(flavor_chunks)

                # Append the formatted flavor text to the ability_texts list
                ability_texts.append(formatted_flavor_text)


        # Concatenate all ability texts together
        combined_text = '\n'.join(ability_texts)

        # Paste the combined text into the text box
        self.wait_and_send_keys("//textarea[@name='description']", combined_text)

        if 'rarity' in self.other_text_fields:
            print(f"Trying to select rarity")
            self.select_dropdown_option_by_value("rarity", self.other_text_fields['rarity'])
            print(f"Done selecting rarity")

    def select_dropdown_option_by_value(self, dropdown_id, value):
        try:
            print(f"Trying to select dropdown")
            dropdown_element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.ID, dropdown_id))
            )
            select = Select(dropdown_element)
            select.select_by_value(value)
        except TimeoutException:
            print(f"Timeout waiting for dropdown with ID {dropdown_id}")
        except NoSuchElementException:
            print(f"Dropdown with ID {dropdown_id} not found")
        except Exception as e:
            print(f"An error occurred while selecting {value} from dropdown {dropdown_id}: {e}")


    def wait_and_send_keys(self, selector, text, by=By.XPATH):
        element = WebDriverWait(self.driver, 20).until(
            EC.element_to_be_clickable((by, selector))
        )
        element.clear()
        element.send_keys(text)

    def select_custom_dropdown_option(self, dropdown_id, option_text):
        # Find and click the <span> element to open the dropdown fully
        span_element = self.driver.find_element(By.ID, "select2-chosen-2")
        span_element.click()

        # After clicking the span, click the Select2 box if necessary
        select2_box = self.driver.find_element(By.ID, dropdown_id)
        self.driver.execute_script("arguments[0].click();", select2_box)

        # Wait for the specific search input to become visible
        search_input = WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((By.ID, "s2id_autogen2_search"))
        )

        # Type the option text into the specific search box
        # Type the option text into the specific search box
        search_input.send_keys(option_text)

        # Press Enter key
        search_input.send_keys(Keys.ENTER)

        # Add a short sleep to ensure the selection is processed
        time.sleep(1)
        print(f"selected creature type")


    def robust_click(self, selector, by):
        element = WebDriverWait(self.driver, 20).until(
            EC.element_to_be_clickable((by, selector))
        )
        self.driver.execute_script("arguments[0].scrollIntoView(true);", element)
        time.sleep(1)
        try:
            element.click()
        except ElementClickInterceptedException:
            self.driver.execute_script("arguments[0].click();", element)

    def preview_card(self):
        # Click the "Preview Card" button
        preview_button = WebDriverWait(self.driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "input[value='Preview Card']"))
        )
        preview_button.click()

        # Wait for the "Publish" button on the new screen
        publish_button = WebDriverWait(self.driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='/src/actions/save']"))
        )
        publish_button.click()

    def run(self):
        self.navigate_to_page()
        self.upload_image_and_confirm()
        self.finalize_card_creation()
        self.enter_card_title_and_other_fields()
        self.preview_card()


def main():
    openai_api_key = 'YOUR_OPENAI_API_KEY'
    google_creds_file = 'PATH_TO_YOUR_GOOGLE_CREDS_FILE.json'
    spreadsheet_id = 'YOUR_SPREADSHEET_ID'
    folder_id = 'YOUR_FOLDER_ID'

    # Initialize the MTGCardArtCreator with your API keys and IDs
    art_creator = MTGCardArtCreator(openai_api_key, google_creds_file, spreadsheet_id, folder_id)

    # Generate and upload images based on the spreadsheet data
    created_cards_info = art_creator.generate_and_upload_images()

    # Check if there are any cards to process
    if not created_cards_info:
        print("No cards to create.")
        return

    # Initialize the MTGCardCreator with the first card's details
    card_creator = MTGCardCreator(created_cards_info[0]['image_path'], created_cards_info[0]['card_title'], created_cards_info[0]['other_text_fields'])
    card_creator.login()  # Perform login here, before creating any cards

    # Now, iterate over all cards including the first one, because we're not creating a new instance each time
    for card_info in created_cards_info:
        # Update the instance with new card info instead of creating a new instance
        card_creator.image_path = card_info['image_path']
        card_creator.card_title = card_info['card_title']
        card_creator.other_text_fields = card_info['other_text_fields']

        # Execute the steps to create the card using the existing instance
        card_creator.run()

if __name__ == "__main__":
    main()
