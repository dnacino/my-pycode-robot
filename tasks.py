from robocorp.tasks import task
from robocorp import browser, http, log
from RPA.PDF import PDF
from RPA.Tables import Tables
from RPA.FileSystem import FileSystem
from RPA.Archive import Archive


@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=100,
        headless=False
    )

    download_order_file()
    open_robot_order_page()
    order_table = read_order_file_tocsv()
    for ordered_unit in order_table:
        fillup_order_in_the_form(ordered_unit)
        dreceipt, screenshot_fn = preview_order_take_screenshot(ordered_unit)
        save_order_details(ordered_unit, dreceipt, screenshot_fn)
        browser.page().click("text=Order another robot")
    archive_the_order_files()


def download_order_file():
    '''downoad csv file'''
    csv_url = 'https://robotsparebinindustries.com/orders.csv'
    http.download(csv_url, overwrite=True)


def open_robot_order_page():
    '''open robot order site'''
    robot_site = "https://robotsparebinindustries.com/#/robot-order"
    browser.goto(robot_site)


def read_order_file_tocsv():
    '''read csv content to the table'''
    tb = Tables()
    return tb.read_table_from_csv(path="orders.csv", delimiters=",", header=True)


def fillup_order_in_the_form(dUnit):
    page = browser.page()
    page.click("text=OK")
    page.click("id=head")
    page.select_option("id=head", dUnit["Head"])
    id_body_str = "id=id-body-" + str(dUnit["Body"])
    page.locator(id_body_str).scroll_into_view_if_needed
    page.check(id_body_str)
    page.get_by_placeholder(
        "Enter the part number for the legs").fill(value=dUnit["Legs"])
    page.fill("id=address", value=dUnit["Address"])


def preview_order_take_screenshot(dUnit):
    page = browser.page()
    page.locator("text=Preview").scroll_into_view_if_needed
    page.click("text=Preview")
    order_path_str = "output/order_" + dUnit["Order number"] + ".png"
    page.wait_for_timeout(1000)
    page.locator("id=robot-preview-image").screenshot(path=order_path_str)
    while True:
        at_fault = False
        page.click("id=order")
        log.console_message("Click order while level\n", kind="stdout")
        try:
            dreceipt = page.locator("id=receipt").inner_html()
        except:
            at_fault = True
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            log.console_message(
                "About to click order try level\n", kind="stdout")
        if at_fault == False:
            break
    log.console_message(dreceipt, kind="stdout")
    return dreceipt, order_path_str


def save_order_details(dUnit, dreceipt, order_path_str):
    iopdf = PDF()
    receipt_path_str = "output/receipt_order_" + dUnit["Order number"] + ".pdf"

    iopdf.html_to_pdf(dreceipt, receipt_path_str)
    iopdf.open_pdf(receipt_path_str)
    iopdf.add_watermark_image_to_pdf(order_path_str, receipt_path_str)
    try:
        iopdf.close_pdf(receipt_path_str)
    except:
        pass
    fs = FileSystem()
    fs.remove_file(order_path_str, missing_ok=True)


def archive_the_order_files():
    fs = FileSystem()
    fs.remove_file("output/placed_orders.zip")
    zippy = Archive()
    zippy.archive_folder_with_zip(
        folder="output", archive_name="output/placed_orders.zip", include="*.pdf")
    pdf_files = fs.find_files('output/*.pdf')
    for fn_File in pdf_files:
        pdf_fn = fn_File[0]
        fs.remove_file(pdf_fn)
