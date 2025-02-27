import os
import re
import time
import logging
import smtplib
import pymongo
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv
from urllib.parse import urljoin
from urllib.parse import urlencode
from collections import defaultdict
from email.mime.text import MIMEText
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from email.mime.multipart import MIMEMultipart
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver

# ----------------------------------
# :: ENV Variable Loader
# ----------------------------------

"""
The load_dotenv() function loads environment variables from a .env file into the environment, allowing you to access them using os.getenv() in your code.
"""

load_dotenv()


# ----------------------------------
# :: Logging Variable
# ----------------------------------

"""
This code sets the logging level of the "undetected_chromedriver" logger to ERROR, suppressing less severe log messages.
"""

logging.getLogger("undetected_chromedriver").setLevel(logging.ERROR)

# ----------------------------------
# :: Xpath Paths
# ----------------------------------

"""
This code retrieves environment variable values and assigns them to respective variables such as DIR, PHONE, WEBSITE, ADDRESS, etc.
"""

dir = os.getenv("DIR")
page = os.getenv("PAGE")
area = os.getenv("AREA")
phone = os.getenv("PHONE")
trade = os.getenv("TRADE")
button = os.getenv("BUTTON")
region = os.getenv("REGION")
website = os.getenv("WEBSITE")
address = os.getenv("ADDRESS")
base_url = os.getenv("BASE_URL")
contact_us = os.getenv("CONTACT_US")
company_name = os.getenv("COMPANY_NAME")
chrome_driver = os.getenv("CHROME_DRIVER")
mongo_connection = os.getenv("MONGO_CONNECTION")
locations = os.getenv("LOCATION", "").split(",")
contractors = os.getenv("CONTRACTORS", "").split(",")
contractors_excel = os.getenv("CONTRACTORS_EXCEL", "").split(",")
excel_sheet_folder_path = os.getenv("EXCEL_SHEET_FOLDER_PATH")
excel_sheet_folder_path = os.getenv("EXCEL_SHEET_FOLDER_PATH")

# ----------------------------------
# :: Blue Book Class
# ----------------------------------

"""
This class, BlueBook, performs web scraping to gather data from a website, processes it, stores it in a MongoDB database, and saves the results to an Excel file.
"""


class BlueBook:

    # ----------------------------------
    # :: __init__ function
    # ----------------------------------

    """
    This code initializes a class, setting up a web driver, MongoDB client, and collection, handling potential errors during the process.
    """

    def __init__(self):
        self.driver = None
        self.client = None
        self.location = None
        self.driver_quit = False
        self.contractor = None
        try:
            self.driver = self.google_chrome_function()
            database = self.mongodb_connection_function(
                db_name="bluebook", collection_name="company_details"
            )
            self.client = database["client"]
            self.collection = database["collection"]
        except Exception as e:
            logging.error(f"Failed to initialize BlueBook: {e}")
            raise

    # ----------------------------------
    # :: Start Request Function
    # ----------------------------------

    """
        This function scrapes hrefs from multiple pages, checks for existing entries in a MongoDB collection, and inserts only the new hrefs.
    """

    def start_requests(self):
        try:
            for location in locations:
                self.location = location.strip()
                for contractor in contractors:
                    self.driver.get("https://www.thebluebook.com/")
                    self.contractor = contractor.strip()
                    wait = WebDriverWait(self.driver, 30)
                    search_term_input = wait.until(
                        EC.visibility_of_element_located(
                            (
                                By.XPATH,
                                "/html/body/div[1]/main/section/div/div/div[1]/form/div/div[2]/div/input",
                            )
                        )
                    )
                    search_term_input.clear()
                    search_term_input.send_keys(contractor)

                    region_input = wait.until(
                        EC.presence_of_element_located(
                            (
                                By.XPATH,
                                "/html/body/div[1]/main/section/div/div/div[1]/form/div/div[3]/div[1]/div/input",
                            )
                        )
                    )
                    region_input.clear()
                    region_input.send_keys(location)

                    search_buttons = self.driver.find_elements(
                        By.XPATH,
                        "/html/body/div[1]/main/section/div/div/div[1]/form/div/div[3]/div[1]/div/button",
                    )
                    if search_buttons:
                        current_url = self.driver.current_url
                        search_buttons[0].click()
                        wait.until(lambda driver: driver.current_url != current_url)
                    else:
                        logging.error("Search button not found!")

                    url = self.driver.current_url
                    if url is not None and url.strip() is not None:
                        self.driver.get(url)
                        page_number = self.xpath_varification_function(
                            driver=self.driver, xpath=page, timeout=20
                        )
                        page_text = page_number[0].text if page_number else None
                        numbers = max(map(int, re.findall(r"\d+", page_text)))
                        self.collection.create_index("href", unique=True)
                        href_data_list = []
                        for i in range(1, int(numbers) + 1):
                            page_url = f"{url}&page={i}"
                            self.driver.get(page_url)
                            anchor_elements = self.xpath_varification_function(
                                self.driver, contact_us
                            )
                            page_hrefs = [
                                {"href": anchor.get_attribute("href")}
                                for anchor in anchor_elements
                                if anchor.get_attribute("href")
                            ]

                            href_data_list.extend(page_hrefs)

                            logging.info(
                                f"Processed page {i} and extracted {len(page_hrefs)} hrefs."
                            )

                        if href_data_list:
                            operations = []
                            send_email = "No"

                            for data in href_data_list:
                                operation = pymongo.UpdateOne(
                                    {"href": data["href"]},
                                    {
                                        "$setOnInsert": {
                                            **data,
                                            "send_email": send_email,
                                            "location": self.location,
                                            "trade": self.contractor,
                                        }
                                    },
                                    upsert=True,
                                )
                                operations.append(operation)

                            if operations:
                                result = self.collection.bulk_write(operations)
                                logging.info(
                                    f"Bulk write operation completed: {result.inserted_count} inserted, {result.modified_count} updated."
                                )
                            else:
                                logging.info("No new hrefs to insert.")
                        else:
                            logging.warning("No hrefs found to process.")
        except Exception as e:
            logging.error(f"An error occurred in start_requests: {e}")
        finally:
            self.get_element_page()

    # ----------------------------------
    # :: Phone Page Function
    # ----------------------------------

    """
    This function iterates through stored URLs, scrapes specific element data (phone, website, address, etc.)
    from each page, and updates the MongoDB collection with the extracted information.
    """

    def get_element_page(self):
        hrefs_cursor = self.collection.find(
            {
                "href": {"$ne": None},
                "send_email": "No",
                "company_name": {"$exists": False},
            }
        )
        element_map = {
            phone: "phone",
            website: "website",
            dir: "dir",
            address: "address",
            company_name: "company_name",
        }
        for document in hrefs_cursor:
            href = document.get("href")
            doc_id = document.get("_id")
            if not href:
                continue
            full_url = urljoin(base_url, href)
            for element, element_name in element_map.items():
                try:
                    self.driver.get(full_url)
                    time.sleep(1)
                    elements = self.xpath_varification_function(self.driver, element)

                    if elements:
                        for component in elements:
                            if element in [phone, website]:
                                phone_text = component.get_attribute("href").replace(
                                    "tel:", ""
                                )
                                self.mongodb_update_function(
                                    self.collection, element_name, phone_text, doc_id
                                )
                            else:
                                text = component.text.strip()
                                self.mongodb_update_function(
                                    self.collection, element_name, text, doc_id
                                )

                    else:
                        logging.warning(f"No {element_name} found on {full_url}")
                except Exception as e:
                    logging.error(f"Error while processing {full_url}")

    # ----------------------------------
    # :: Google Chrome Function
    # ----------------------------------

    """
    This function initializes a Chrome WebDriver with custom options and handles any errors during setup.
    """

    def google_chrome_function(self):
        try:
            chrome_driver_path = os.getenv("CHROME_DRIVER")
            chrome_options = Options()
            chrome_options.add_argument("--v=1")
            chrome_options.add_argument("--headless")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--enable-logging")
            chrome_options.add_argument("--disable-dev-shm-usage")
            service = Service(executable_path=chrome_driver_path)
            driver = webdriver.Chrome(service=service, options=chrome_options)
            return driver

        except EnvironmentError as env_err:
            logging.error(str(env_err))
            raise

        except WebDriverException as wd_err:
            logging.error(f"Failed to initialize WebDriver: {str(wd_err)}")
            raise

        except Exception as e:
            logging.error(
                f"An unexpected error occurred while initializing Chrome WebDriver: {str(e)}"
            )
            raise

    # ----------------------------------
    # :: Google Chrome Function
    # ----------------------------------

    """
    This function connects to a MongoDB database and collection, logging success or error messages based on the outcome.
    """

    def mongodb_connection_function(self, db_name, collection_name):
        try:
            client = pymongo.MongoClient(mongo_connection)
            db = client[db_name]
            collection = db[collection_name]
            return {"client": client, "collection": collection}
        except Exception as e:
            logging.error(
                f"An unexpected error occurred while connecting to MongoDB: {str(e)}"
            )
            raise

    # ----------------------------------
    # :: Google Chrome Function
    # ----------------------------------

    """
    This function verifies the presence of an element(s) on a webpage using a specified XPath, with a timeout option,
    and handles errors if the element is not found or other issues occur.
    """

    def xpath_varification_function(self, driver, xpath, timeout=20):

        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        elements = driver.find_elements(By.XPATH, xpath)
        if elements:
            return elements

    # ----------------------------------
    # :: MongoDb Update Function
    # ----------------------------------

    """
    The mongodb_update_function updates a document in a MongoDB collection or inserts it if it doesn't exist, logging the outcome.
    """

    def mongodb_update_function(self, collection, element_name, element, doc_id):
        try:
            result = collection.update_one(
                {"_id": doc_id}, {"$set": {element_name: element}}, upsert=True
            )

            if result.modified_count > 0:
                logging.info(f"Document with _id {doc_id} was updated.")
            elif result.upserted_id:
                logging.info(f"Document with _id {doc_id} was inserted.")
            else:
                logging.info(f"No change made to document with _id {doc_id}.")
        except Exception as e:
            logging.error(f"Unexpected error: {str(e)}")
            raise

    # ---------------------------------------
    # :: Email Send Function
    # ---------------------------------------

    """
    The email_send function composes and sends an email with an attachment from a specified MongoDB path using SMTP,
    leveraging credentials stored in environment variables.
    """

    def email_send(self, file_path, count):
        HOST = os.getenv("HOST")
        PORT = os.getenv("PORT")
        FROM_EMAIL = os.getenv("FROM_EMAIL")
        TO_EMAIL = os.getenv("TO_EMAIL", "").split(",")
        PASSWORD = os.getenv("PASSWORD")
        if not all([HOST, PORT, FROM_EMAIL, PASSWORD, TO_EMAIL]):
            logging.error("Missing required environment variables.")
            return False
        subject = f"Date: {datetime.now()} Time: {datetime.now().strftime('%H:%M')} file records"
        body = f"There are {count} complete records"
        msg = MIMEMultipart()
        msg["From"] = FROM_EMAIL
        msg["To"] = ", ".join(TO_EMAIL)
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))
        if os.path.exists(file_path):
            try:
                with open(file_path, "rb") as attachment:
                    part = MIMEText(attachment.read(), "base64", "utf-8")
                    part.add_header(
                        "Content-Disposition",
                        f"attachment; filename={os.path.basename(file_path)}",
                    )
                    msg.attach(part)
            except Exception as e:
                logging.error(f"Error attaching file: {e}")
                return False

        try:
            with smtplib.SMTP(HOST, int(PORT)) as smtp:
                smtp.starttls()
                smtp.login(FROM_EMAIL, PASSWORD)
                smtp.sendmail(FROM_EMAIL, TO_EMAIL, msg.as_string())
                logging.info("Email sent successfully!")
                return True
        except smtplib.SMTPException as e:
            logging.error(f"SMTP error occurred: {e}")
        except Exception as e:
            logging.error(f"An unexpected error occurred: {e}")
        return False

    # ---------------------------------------
    # :: Excel File Save Function
    # ---------------------------------------

    """
    The Excel File Save Function retrieves company details from a MongoDB database,
    creates a DataFrame, and saves it as an Excel file in a specified folder, with a timestamped filename.

    """

    def excel_file_save_function(self):
        try:
            for location in locations:
                for contractor in contractors:
                    logging.info(f"Processing records for: {contractor}")
                    results = self.collection.find(
                        {
                            "trade": contractor,
                            "location": location,
                            "send_email": "No",
                            "company_name": {"$exists": True},
                        }
                    )
                    documents = []
                    for doc in results:
                        doc = defaultdict(lambda: "not found", doc)
                        documents.append(doc)

                    required_columns = [
                        "company_name",
                        "phone",
                        "dir",
                        "website",
                        "address",
                        "location",
                        "trade",
                    ]
                    if documents:
                        logging.info(
                            f"Found {len(documents)} valid documents for {contractor}."
                        )
                        df = pd.DataFrame(documents, columns=required_columns)
                        if not os.path.exists(excel_sheet_folder_path):
                            os.makedirs(excel_sheet_folder_path)
                        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                        file_name = os.path.join(
                            excel_sheet_folder_path, f"{contractor}_{timestamp}.xlsx"
                        )

                        try:
                            logging.info(
                                f"Saving file as {file_name} for {contractor}..."
                            )
                            df.to_excel(file_name, index=False, header=True)
                        except Exception as e:
                            logging.error(
                                f"Error saving Excel file for {contractor}: {e}"
                            )
                            continue
                        email_status = self.email_send(
                            file_path=file_name, count=len(documents)
                        )
                        if email_status:
                            logging.info(f"Email sent successfully for {contractor}.")
                            self.collection.update_many(
                                {"send_email": "No", "trade": contractor},
                                {"$set": {"send_email": "Yes"}},
                            )
                        else:
                            logging.error(f"Email failed for {contractor}.")
                    else:
                        logging.warning(f"No valid documents found for {contractor}.")

                return "All cities processed."
        except Exception as e:
            logging.error(f"An unexpected error occurred: {e}")
            return "An error occurred while processing the file."

    # ----------------------------------
    # :: ___del__ function
    # ----------------------------------

    """
    This destructor closes the webdriver, MongoDB client, and attempts to save data to an Excel file upon object deletion.
    """

    def __del__(self):
        try:
            logging.info("Calling excel file save function...")
            self.excel_file_save_function()
            if self.driver:
                self.driver.quit()
            if self.client:
                self.client.close()
        except Exception as e:
            logging.error(f"Error while saving Excel file: {e}")


# ----------------------------------
# :: Main Function
# ----------------------------------

"""
This main function instantiates the BlueBook spider, starts the scraping process, and handles potential exceptions.
"""


def main():
    try:
        spider = BlueBook()
        spider.get_element_page()
    except Exception as e:
        logging.error(f"An error occurred in the main function: {e}")


# ----------------------------------
# :: Run code
# ----------------------------------

"""
This code block ensures the main() function is executed only when the script is run directly (not imported as a module).
"""

if __name__ == "__main__":
    main()