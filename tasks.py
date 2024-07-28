from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF


@task
def minimal_task():
    """Insert the sales data for the week and export it as a PDF"""
    browser.configure(slowmo=10,)
    open_the_intranet_website()
    log_in()
    download_excel("https://robotsparebinindustries.com/SalesData.xlsx")
    fill_and_submit_sales_form()
    export_as_pdf()
    collect_results()


def open_the_intranet_website():
    """Open browser and connects to the URL"""
    browser.goto("https://robotsparebinindustries.com/")


def log_in():
    page = browser.page()
    page.fill("#username", "maria")
    page.fill("#password", "thoushallnotpass")
    page.click("button:text('Log in')")


def fill_and_submit_sales_form():
    excel = Files()
    page = browser.page()

    excel.open_workbook('SalesData.xlsx')
    worksheet = excel.read_worksheet_as_table("data", True, True)
    excel.close_workbook()

    for row in worksheet:
        page.fill("#firstname", row["First Name"])
        page.fill("#lastname", row["Last Name"])
        page.fill("#salesresult", str(row["Sales"]))
        page.select_option("#salestarget", str(row["Sales Target"]))
        page.click("text=Submit")


def download_excel(path):
    http = HTTP()
    http.download(
        url=path, overwrite=True)


def collect_results():
    page = browser.page()
    page.screenshot(path="output/sales_sumary.png")
    page.click("text=Log out")


def export_as_pdf():
    page = browser.page()
    sales_results = page.locator("#sales-results").inner_html()

    pdf = PDF()
    pdf.html_to_pdf(sales_results, "output/sales_results.pdf")
