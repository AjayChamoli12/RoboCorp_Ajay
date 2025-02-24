from robocorp import browser, vault
from robocorp.tasks import task
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from RPA.Archive import Archive
from RPA.Tables import Table
import pandas as pd
import json as js


@task
def minimal_task():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
      slowmo=1000,
   )
    download_csv_file()
    Open_website()
    fill_form_with_CSV_data()
    archive_receipts()
       

def download_csv_file():
    http=HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv",overwrite=True)


def fill_form_with_CSV_data():
   """Read data from CSV and fill in the Order form"""
   data=pd.read_csv("orders.csv")
   print (type(data)) 
   for index, row in data.iterrows():
      fill_and_Place_Order(row)


def Open_website():
    """open_robot_order_website()"""
    page=browser.page()
    browser.goto("https://robotsparebinindustries.com/#/robot-order")


def fill_and_Place_Order(Order_Details):
    """Fills in the Order form and submits it"""
    page=browser.page()
    page.click("button:text('OK')")
    page.select_option("#head",str(Order_Details["Head"]))
    page.locator("#id-body-"+ str(Order_Details["Body"] )).click()
    page.locator("xpath=//label[contains(.,'3. Legs:')]/../input").fill(str(Order_Details["Legs"]))
    page.fill("#address",str(Order_Details["Address"]))
    order_anotherbutton=False
    
    """order_anotherbutton=page.locator("#order-another").is_visible()"""
    
    while order_anotherbutton==False:
        page.click("button:text('Order')")
        order_anotherbutton=page.locator("#order-another").is_visible()
        if order_anotherbutton == True:
            pdf_path = store_receipt_as_pdf(int(Order_Details["Order number"]))
            screenshot_path = screenshot_robot(int(Order_Details["Order number"]))
            embed_screenshot_to_receipt(screenshot_path, pdf_path)
            order_another_bot()
            break

def order_another_bot():
    """Clicks on order another button to order another bot"""
    page = browser.page()
    page.click("#order-another")
    

def screenshot_robot(order_number):
    """Takes screenshot of the ordered bot image"""
    page = browser.page()
    screenshot_path = "output/screenshots/{0}.png".format(order_number)
    page.locator("#robot-preview-image").screenshot(path=screenshot_path)
    return screenshot_path

def embed_screenshot_to_receipt(screenshot_path, pdf_path):
    """Embeds the screenshot to the bot receipt"""
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(image_path=screenshot_path, 
                                   source_path=pdf_path, 
                                   output_path=pdf_path)
    
def archive_receipts():
    """Archives all the receipt pdfs into a single zip archive"""
    lib = Archive()
    lib.archive_folder_with_zip("./output/receipts", "./output/receipts.zip")

def store_receipt_as_pdf(order_number):
    """This stores the robot order receipt as pdf"""
    page = browser.page()
    order_receipt_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    pdf_path = "output/receipts/{0}.pdf".format(order_number)
    pdf.html_to_pdf(order_receipt_html, pdf_path)
    return pdf_path
