import os
import requests
from config import openai_api_key, google_creds_file, spreadsheet_id, folder_id, username, password
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
    """
    A class to generate and upload Magic: The Gathering card art images using OpenAI's image generation API, and manage them on Google Drive and Google Sheets.

    Attributes:
    openai_api_key (str): The API key for OpenAI.
    google_creds_file (str): Path to the Google service account credentials file.
    spreadsheet_id (str): The ID of the Google Sheets spreadsheet containing card information.
    folder_id (str): The ID of the Google Drive folder where images will be uploaded.
    client (OpenAI): The OpenAI client for API access.
    creds (Credentials): Google service account credentials.
    sheet_service (Resource): The Google Sheets API service instance.
    drive_service (Resource): The Google Drive API service instance.
    """
    def __init__(self, openai_api_key, google_creds_file, spreadsheet_id, folder_id):
        """
        Initializes the MTGCardArtCreator with the necessary credentials and IDs.

        Parameters:
        openai_api_key (str): The API key for OpenAI.
        google_creds_file (str): Path to the Google service account credentials file.
        spreadsheet_id (str): The ID of the Google Sheets spreadsheet containing card information.
        folder_id (str): The ID of the Google Drive folder where images will be uploaded.
        """
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
        """
        Adjusts the size of an image to a maximum of 900x900 pixels and saves it.

        Parameters:
        image_path (str): The path to the image file to be adjusted.
        """
        with Image.open(image_path) as img:
            size = (900, 900)
            img.thumbnail(size, Image.Resampling.LANCZOS)
            img.save(image_path, optimize=True)

    def generate_and_upload_images(self):
        """
        Generates and uploads images based on the prompts from the Google Sheets spreadsheet.

        Iterates through a specific range in the 'Cards' sheet, generates images using OpenAI's DALL·E, and uploads them to Google Drive.

        Returns:
        list: A list of dictionaries containing information about the created cards and their image paths.
        """
        range_name = 'Cards'  # Adjust as needed
        result = self.sheet_service.spreadsheets().values().get(spreadsheetId=self.spreadsheet_id, range=range_name).execute()
        values = result.get('values', [])
        created_cards_info = []

        for row in values[1:]:  # Assuming the first row contains headers
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
    """
    A class to automate the process of creating MTG cards on the MTGCardsmith website using Selenium.

    Attributes:
    image_path (str): The path to the card's image file.
    card_title (str): The title of the MTG card.
    other_text_fields (dict): Other text fields related to the card, such as type, abilities, etc.
    driver (WebDriver): The Selenium WebDriver instance.
    is_logged_in (bool): Flag indicating whether the user is logged in or not.
    """
    def __init__(self, image_path, card_title, other_text_fields=None):
        """
        Initializes the MTGCardCreator with an image path, card title, and optional text fields.

        Parameters:
        image_path (str): The path to the card's image file.
        card_title (str): The title of the MTG card.
        other_text_fields (dict, optional): Other text fields related to the card. Defaults to None.
        """
        self.image_path = image_path
        self.card_title = card_title
        self.other_text_fields = other_text_fields if other_text_fields else {}
        self.driver = self.init_driver()
        self.is_logged_in = False  # Flag to track login status

    def init_driver(self):
        """
        Initializes and returns a Selenium WebDriver instance with Chrome options.

        Returns:
        WebDriver: The initialized Selenium WebDriver instance.
        """
        options = Options()
        options.add_experimental_option('excludeSwitches', ['enable-logging'])  # Suppress logging
        return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    def login(self):
        """
        Logs into the MTGCardsmith website using predefined credentials.
        """
        if self.is_logged_in:  # Check if already logged in
            return  # Skip login if already logged in

        username = username
        password = password

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
        """
        Navigates to a specified URL using the Selenium WebDriver.

        Parameters:
        url (str): The URL to navigate to. Defaults to the MTG card maker page.
        """
        self.driver.get(url)
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file']"))
            )
        except TimeoutException:
            self.driver.execute_script("window.stop();")

    def upload_image_and_confirm(self):
        """
        Uploads an image for the card and confirms the upload on the MTGCardsmith website.
        """
        absolute_image_path = os.path.abspath(self.image_path)
        file_input = self.driver.find_element(By.CSS_SELECTOR, "input[type='file']")
        file_input.send_keys(absolute_image_path)
        confirm_button = WebDriverWait(self.driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".slim-editor-btn.slim-btn-confirm"))
        )
        confirm_button.click()
        time.sleep(2)

    def finalize_card_creation(self):
        """
        Finalizes the card creation process by clicking the 'Next' button on the MTGCardsmith website.
        """
        next_button = WebDriverWait(self.driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='submit'][value='Next']"))
        )
        next_button.click()

    def enter_card_title_and_other_fields(self):
        """
        Enters the card title and other related text fields into the MTGCardsmith card creator form.
        """
        self.wait_and_send_keys("//input[@name='name'][@placeholder='Card Title (Required)']", self.card_title)

        if 'mana_value' in self.other_text_fields:
            self.wait_and_send_keys("//input[@name='custom_mana'][@placeholder='Select your casting cost...']", self.other_text_fields['mana_value'])

        if 'card_type' in self.other_text_fields:
            print(f"Trying to select card_type")
            self.select_custom_dropdown_option('s2id_autogen1', self.other_text_fields['card_type'])
            print(f"Done selecting creature_type")

        # New approach for handling abilities
        ability_keys = ['static_abilities', 'triggered_abilities_1', 'triggered_abilities_2', 'triggered_abilities_3', 'triggered_abilities_4']
        for key in ability_keys:
            self.process_and_add_ability_text(key)


    def process_and_add_ability_text(self, ability_key):
        """
        Processes and adds the text for a given ability field to the 'ability_texts' list.

        Parameters:
        ability_key (str): The key corresponding to the ability in 'other_text_fields'.
        """
        ability_text = self.other_text_fields.get(ability_key)
        if ability_text:
            # Split ability text into chunks of 45 characters, respecting word boundaries where possible.
            formatted_chunks = self.split_ability_text(ability_text, 45)
            # Join the chunks with line breaks
            formatted_ability_text = '\n'.join(formatted_chunks)
            self.ability_texts.append(formatted_ability_text)

    @staticmethod
    def split_ability_text(text, chunk_size):
        """
        Splits text into chunks of a specified size, attempting to respect word boundaries.

        Parameters:
        text (str): The text to split.
        chunk_size (int): The maximum size of each chunk.

        Returns:
        list: A list of text chunks.
        """
        chunks = [text[i:i+chunk_size] for i in range(0, len(text), chunk_size)]
        formatted_chunks = []

        for chunk in chunks:
            split_index = chunk.rfind(' ')
            if split_index != -1 and len(chunk) > chunk_size:
                formatted_chunks.append(chunk[:split_index])
                formatted_chunks.append(chunk[split_index+1:])
            else:
                formatted_chunks.append(chunk)

        return formatted_chunks

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
        """
        Executes the complete card creation process by navigating to the page, uploading an image, and entering card details.
        """
        self.navigate_to_page()
        self.upload_image_and_confirm()
        self.finalize_card_creation()
        self.enter_card_title_and_other_fields()
        self.preview_card()


def main():
    openai_api_key = openai_api_key
    google_creds_file = google_creds_file
    spreadsheet_id = spreadsheet_id
    folder_id = folder_id

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
