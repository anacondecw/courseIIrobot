from robocorp.tasks import task
from robocorp import browser
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
from RPA.FileSystem import FileSystem


from RPA.HTTP import HTTP

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    open_robot_order_website()
    orders = get_orders()

    for row in orders:
        fill_and_submit_sales_form(row)
        
    archive_receipts()



def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order") 
    page = browser.page()



def get_orders():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    library = Tables()
    orders = library.read_table_from_csv(
    "orders.csv", columns=["Order number","Head","Body","Legs","Address"])
    return orders



def fill_and_submit_sales_form(sales_rep):
    """Fills in the sales data and click the 'Submit' button"""
    page = browser.page()
    page.click('//*[@id="root"]/div/div[2]/div/div/div/div/div/button[1]')


    page.select_option("#head", str(sales_rep["Head"]))
    page.click("#id-body-"+str(str(sales_rep["Body"])))
    page.fill(".form-control", sales_rep["Legs"])
    page.fill("#address", sales_rep["Address"])

    page.click("#order")

    # Imprimir el resultado de la verificación
    try: 
        page.click('//*[@class="alert alert-danger"]')
        page.click("#order")
 
        path_pdf=store_receipt_as_pdf(sales_rep["Order number"])
        path_screen=screenshot_robot(sales_rep["Order number"])
        embed_screenshot_to_receipt(path_screen,path_pdf,sales_rep["Order number"])
        
        page.click("#order-another")


    except:
            
        path_pdf=store_receipt_as_pdf(sales_rep["Order number"])
        path_screen=screenshot_robot(sales_rep["Order number"])
        embed_screenshot_to_receipt(path_screen,path_pdf,sales_rep["Order number"])
        page.click("#order-another")


def store_receipt_as_pdf(order_number):
    page = browser.page()
    sales_results_html = page.locator("#receipt").inner_html()

    pdf = PDF()
    path="output/"+str(order_number)+".pdf"
    pdf.html_to_pdf(sales_results_html, path)
    return path

def screenshot_robot(order_number):    
    """Take a screenshot of the page"""
    page = browser.page()
    path="output/"+str(order_number)+".png"
    return path

    page.screenshot(path=path)

def embed_screenshot_to_receipt(screenshot, pdf_file,order):
    pdf=PDF()
    pdf.add_files_to_pdf(
    files=[pdf_file,screenshot],
    target_document="output/merged/merged-pdf"+str(order)+".pdf"
    )

def archive_receipts():
    # Inicializar el objeto de FileSystem
    fs = Archive()

    # Directorio donde se encuentran los archivos PDF de los recibos
    receipts_directory = "output/merged/"
    
    # Archivo ZIP que será creado
    zip_filename = "output/receipts_archive.zip"

    # Archivar todos los archivos PDF en un archivo ZIP
    fs.archive_folder_with_zip(receipts_directory, zip_filename)


