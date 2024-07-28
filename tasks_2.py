from robocorp.tasks import task
from robocorp import browser
from RPA.Tables import Tables
from RPA.HTTP import HTTP
from RPA.PDF import PDF
from RPA.FileSystem import FileSystem
from RPA.Archive import Archive


@task
def level2_task():
    """ Downloads the data, then completes the orders and create recipts
    """
    browser.configure(slowmo=100)
    browser.page().set_viewport_size({"width": 1920, "height": 1080})
    cleanup()
    connect_to_order_page()
    fill_orders()
    archive_recipts()
    cleanup()


def cleanup():
    fs = FileSystem()
    if fs.does_directory_exist("output/orders"):
        fs.remove_directory("output/orders", True)

    if fs.does_directory_exist("output/screenshots"):
        fs.remove_directory("output/screenshots", True)


def connect_to_order_page():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    page = browser.page()
    page.click("button:text('OK')")


def read_csv(path):
    http = HTTP()
    http.download(
        url=path, overwrite=True)
    csv = Tables()
    orders = csv.read_table_from_csv(
        'orders.csv', columns=['Order number', 'Head', 'Body', 'Legs', 'Address'])
    return orders


def take_screenshot(index):
    page = browser.page()
    screenshot_path = "output/screenshots/order_preview_{0}.png".format(index)
    page.locator("#robot-preview-image").screenshot(path=screenshot_path)
    return screenshot_path


def generate_pdfs(input, index, screenshot_path):
    pdf = PDF()
    pdfPath = "output/orders/order_receipts_{0}.pdf".format(index)
    pdf.html_to_pdf(input,
                    pdfPath)
    pdf.add_watermark_image_to_pdf(
        image_path=screenshot_path,
        source_path=pdfPath,
        output_path="output/orders/order_receipts_final_{0}.pdf".format(index))


def fill_orders():
    orders = read_csv("https://robotsparebinindustries.com/orders.csv")
    page = browser.page()
    model_info = {
        "1": "Roll-a-thor",
        "2": "Peanut crusher",
        "3": "D.A.V.E",
        "4": "Andy Roid",
        "5": "Spanner mate",
        "6": "Drillbit 2000"
    }

    for order in orders:
        page.select_option(
            "select#head", value=model_info[order["Head"]]+" head")
        page.click("//input[@value='{0}']".format(order["Body"]))
        page.fill(
            "input[placeholder='Enter the part number for the legs']", order["Legs"])
        page.fill("input[placeholder='Shipping address']", order["Address"])

        page.click("text=Preview")
        sc_path = take_screenshot(order["Order number"])

        page.click("button#order")
        while (page.is_visible(".alert-danger")):
            page.click("button#order")

        recipt_results = page.locator("#receipt").inner_html()
        generate_pdfs(recipt_results, order["Order number"], sc_path)

        page.click("button#order-another")
        page.click("button:text('OK')")


def archive_recipts():
    archive = Archive()
    archive.archive_folder_with_zip(
        "./output/orders", "output/order_recipts.zip", include="order_receipts_final_*.pdf")
