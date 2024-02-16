from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.PDF import PDF
from RPA.Tables import Tables
from RPA.Archive import Archive



@task
def order_robots_from_RobotSpareBin():
    """Insert the sales data for the week and export it as a PDF"""
    browser.configure(
        slowmo=100,
    )
    
    download_excel_file()
    data_excel = get_orders()
    open_robot_order_website()
    set_excel_data_to_form(data_excel)
    archive_receipts()


def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    

def close_annoying_modal():
    page = browser.page()  
    page.click("text=OK")


def fill_and_submit_sales_form(row):
    """Fills in the sales data and click the 'Submit' button"""
    page = browser.page()  
    
    page.select_option("#head", str(row.get('Head')))
    page.click(f"#id-body-{str(row.get('Body'))}")
    page.fill("//input[@class='form-control' and @type='number' and @min='1' and @max='6' and @placeholder='Enter the part number for the legs' and @required='']", row.get('Legs'))
    page.fill("#address", str(row.get('Address')))
    page.click("#preview")

    while True:
        page.click("#order")   
        if not page.is_visible("//div[@class='alert alert-danger']"):
            break


def order_new_robot():
    """Order new robot"""
    page = browser.page() 
    page.click("#order-another")


def extract_order_data():
    """Extracr order data from robot"""
    page = browser.page() 

    order_number = page.locator("//p[@class='badge badge-success']").inner_text()
    pdf = store_receipt_as_pdf(order_number)
    png = screenshot_robot(order_number)
    embed_screenshot_to_receipt(png, pdf)
    

def download_excel_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True, target_file='static/Robot_data.csv')


def get_orders():
    """Read data from excel and fill in the order form"""
    excel = Tables()
    data = excel.read_table_from_csv("static/Robot_data.csv")
    return data


def set_excel_data_to_form(data):
    for x in data:
        close_annoying_modal()
        fill_and_submit_sales_form(x)
        extract_order_data()
        order_new_robot()


def screenshot_robot(order_number):
    """Take a screenshot of the page"""
    page = browser.page()
    png_name = f"output/orders/img/{order_number}.png"
    page.screenshot(path=png_name)
    return png_name


def store_receipt_as_pdf(order_number):
    """Export the data to a pdf file"""
    page = browser.page()
    sales_results_html = page.locator("#receipt").inner_html()

    pdf = PDF()
    pdf_name = f"output/orders/pdf/{order_number}.pdf"
    pdf.html_to_pdf(sales_results_html, pdf_name)
    return pdf_name


def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()
    pdf.add_files_to_pdf(files=[screenshot] , target_document=pdf_file)


def archive_receipts():
    files = Archive()
    files.archive_folder_with_zip('output/orders/pdf/', 'output/orders/order.zip')