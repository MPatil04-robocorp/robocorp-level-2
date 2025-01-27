from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from RPA.PDF import PDF
import time,os
from RPA.Archive import Archive

@task
def order_robots_from_RobotSpareBin():
    
    browser.configure(
        slowmo = 1000,
    )

    open_the_intranet_website()
    orders = get_orders()
    # print(orders)
    close_annoying_model()
    for order in orders:
        fill_the_form(order)
    archive_receipts()
    


def open_the_intranet_website():
    browser.goto('https://robotsparebinindustries.com/#/robot-order')

def get_orders():
    '''getting orders'''
    http = HTTP()
    http.download(url = 'https://robotsparebinindustries.com/orders.csv',overwrite=True)

    
    library = Tables()
    orders  =  library.read_table_from_csv("orders.csv",columns=["Order number","Head","Body","Legs","Address"])
    # print(orders)
    # for order in orders:
    #     print(order)
    return orders

def close_annoying_model():
    page = browser.page()
    page.click('//div[@class="alert-buttons"]/button[contains(text(),"OK")]')
    
def fill_the_form(order):
    page = browser.page()
    time.sleep(1)
    if not os.path.exists(f"output/{str(order['Order number'])}.png"):
        page.select_option("#head", str(order["Head"]))
        page.click('//div[@class="stacked"]//label/input[@value="{}"]'.format(str(order['Body'])))
        page.fill('//input[@placeholder="Enter the part number for the legs"]',str(order['Legs']))
        page.fill('#address',str(order['Address']))
        page.click("#preview")
        try:
            time.sleep(2)
            page.click('#order')
            time.sleep(2)
            store_receipt_as_pdf(order['Order number'])
            screenshot_robot(order['Order number'])
            embed_screenshot_to_receipt(f"output/{str(order['Order number'])}.png",f"output/{str(order['Order number'])}.pdf")
            page.click('#order-another')
            close_annoying_model()
            
        except:
            print("Fails to order a robot",order["Order number"])
            page.reload()
            close_annoying_model()
    else:
        embed_screenshot_to_receipt(f"output/{str(order['Order number'])}.png",f"output/{str(order['Order number'])}.pdf")


def store_receipt_as_pdf(order_number):
    ''''''
    page = browser.page()
    receipt = page.locator("#order-completion").inner_html()
    pdf = PDF()
    pdf.html_to_pdf(receipt, f"output/{str(order_number)}.pdf")

def screenshot_robot(order_number):
    page= browser.page()
    page.screenshot(path=f"output/{str(order_number)}.png")


def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(image_path=screenshot, 
                                   source_path=pdf_file, 
                                   output_path=pdf_file) 
def archive_receipts():
    """Archives all the receipt pdfs into a single zip archive"""
    lib = Archive()
    lib.archive_folder_with_zip("./output/receipts", "./output/receipts.zip")
