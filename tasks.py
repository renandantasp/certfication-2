from robocorp.tasks import task
from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
from RPA.Tables import Tables, Table, Column
from RPA.PDF import PDF
import os
import zipfile
from config import ORDERS_PATH, ORDERS_LINK, COMPANY_PATH, ZIP_PATH, RECEIPTS_DIR

def download_orders_csv() -> None:
    """
    Access the url of the orders.csv and
    download to the output file
    """
    
    http = HTTP()

    os.makedirs(os.path.dirname(ORDERS_PATH), exist_ok=True)

    if os.path.exists(ORDERS_PATH):
        os.remove(ORDERS_PATH)

    http.download(ORDERS_LINK, ORDERS_PATH)

    if os.path.exists(ORDERS_PATH):
        print(f"File downloaded successfully: {ORDERS_PATH}")
    else:
        print("Download failed")

def get_orders() -> Table:
    """
    Read the downloaded csv file
    to a Table type
    """
    download_orders_csv()
    tables = Tables()
    orders = tables.read_table_from_csv(ORDERS_PATH,True)
    return orders

def close_annoying_modal(browser: Selenium) -> None:
    browser.click_button("OK")

def reacess_form(browser: Selenium) -> None:
    browser.click_button("xpath://button[@id='order-another']")

def fill_the_form(browser: Selenium, order: dict[Column, any]) -> str:
    browser.select_from_list_by_value("xpath://select[@id='head']", order['Head'])
    browser.click_element(f"xpath://input[@id='id-body-{order['Body']}']")
    browser.input_text("xpath://input[@placeholder='Enter the part number for the legs']", order['Legs'])
    browser.input_text("xpath://input[@placeholder='Shipping address']", order['Address'])
    browser.click_button("xpath://button[@id='order']")

    order_successful = False
    while not order_successful:
        if browser.is_element_visible("xpath://button[@id='order-another']"):
            order_successful = True
        else:
            browser.wait_and_click_button("xpath://button[@id='order']")

def store_receipt_as_pdf(browser: Selenium, order_number: str):
    pdf = PDF()
    receipt = browser.get_element_attribute("xpath://div[@id='receipt']", 'outerHTML')

    os.makedirs(os.path.dirname(RECEIPTS_DIR), exist_ok=True)
    
    pdf_file_path = f'{RECEIPTS_DIR}/{order_number}.pdf'
    pdf.html_to_pdf(receipt, pdf_file_path)

    image_path = f'{RECEIPTS_DIR}/{order_number}.png'
    browser.screenshot("xpath://div[@id='robot-preview-image']", image_path)

    pdf.add_files_to_pdf(files=[pdf_file_path, image_path], target_document=pdf_file_path)
    os.remove(image_path)

    return pdf_file_path

def archive_receipts():
    with zipfile.ZipFile(ZIP_PATH, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dir, files in os.walk(RECEIPTS_DIR):
            for file in files:
                file_path = os.path.join(root, file)
                zipf.write(file_path, os.path.relpath(file_path, RECEIPTS_DIR))

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    orders = get_orders()

    browser = Selenium(timeout=10, page_load_timeout=10)
    browser.headless = True
    
    browser.open_available_browser(COMPANY_PATH)

    for order in orders:
        close_annoying_modal(browser)
        fill_the_form(browser, order)
        store_receipt_as_pdf(browser, order['Order number'])
        reacess_form(browser)
    archive_receipts()
        