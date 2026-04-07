from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF

@task
def robot_spare_bin_python():
    """
    Main task entry point.

    Executes end-to-end automation workflow:
    - Opens intranet site
    - Authenticates user
    - Downloads Excel sales data
    - Populates and submits form for each record
    - Captures results
    - Exports results as PDF
    - Logs out
    """
    # Configure browser behavior (slow motion aids debugging and stability)
    browser.configure(
        slowmo=100,
    )

    # Step 1: Navigate to application
    open_the_intranet_website()

    # Step 2: Perform authentication
    log_in()

    # Step 3: Retrieve input data
    download_excel_file()

    # Step 4: Process and submit sales records
    # fill_and_submit_sales_form()  # Legacy hardcoded method (kept for reference)
    fill_form_with_excel_data()

    # Step 5: Capture results snapshot
    collect_results()  

    # Step 6: Export results to PDF
    export_as_pdf()

    # Step 7: End session
    log_out()


def open_the_intranet_website():
    """
    Navigates to the RobotSpareBin intranet portal.
    """
    browser.goto("https://robotsparebinindustries.com/")


def log_in():
    """
    Authenticates user into the system.

    NOTE:
    - Credentials are hardcoded; in production, use a secure vault.
    """
    page = browser.page()

    # Populate login credentials
    page.fill("#username", "maria")
    page.fill("#password", "thoushallnotpass")

    # Submit login form
    page.click("button:text('Log in')")


# def fill_and_submit_sales_form():
#     """
#     Legacy implementation using hardcoded values.
#     Retained for reference/testing purposes.
#     """
#     page = browser.page()
#
#     page.fill("#firstname", "John")
#     page.fill("#lastname", "Smith")
#     page.fill("#salesresult", "123")
#     page.select_option("#salestarget", "10000")
#     page.click("text=Submit")


def fill_and_submit_sales_form(sales_rep):
    """
    Populates and submits the sales form using dynamic data.

    Args:
        sales_rep (dict): A single record from Excel containing:
            - First Name
            - Last Name
            - Sales Target
            - Sales
    """
    page = browser.page()

    # Fill employee first name
    page.fill("#firstname", sales_rep["First Name"])

    # Fill employee last name
    page.fill("#lastname", sales_rep["Last Name"])

    # Select sales target from dropdown (converted to string for compatibility)
    page.select_option("#salestarget", str(sales_rep["Sales Target"]))

    # Input actual sales value
    page.fill("#salesresult", str(sales_rep["Sales"]))

    # Submit the form
    page.click("text=Submit")


def download_excel_file():
    """
    Downloads the Excel file containing weekly sales data.

    Overwrites existing file if already present.
    """
    http = HTTP()

    http.download(
        url="https://robotsparebinindustries.com/SalesData.xlsx",
        overwrite=True
    )


def fill_form_with_excel_data():
    """
    Reads Excel data and processes each row sequentially.

    Workflow:
    - Opens workbook
    - Reads worksheet as structured table
    - Iterates through each row and submits form
    - Closes workbook after reading
    """
    excel = Files()

    # Open Excel workbook
    excel.open_workbook("SalesData.xlsx")

    # Read worksheet into table structure
    worksheet = excel.read_worksheet_as_table("data", header=True)

    # Close workbook to free resources
    excel.close_workbook()

    # Iterate through rows and submit form for each entry
    for row in worksheet:
        fill_and_submit_sales_form(row)


def collect_results():
    """
    Captures a screenshot of the final sales summary page.

    Useful for audit/logging purposes.
    """
    page = browser.page()

    page.screenshot(path="output/sales_summary.png")


def export_as_pdf():
    """
    Exports the sales results section into a PDF document.

    Extracts HTML content and converts it into a formatted PDF.
    """
    page = browser.page()

    # Extract HTML content of results section
    sales_results_html = page.locator("#sales-results").inner_html()

    pdf = PDF()

    # Convert HTML content to PDF file
    pdf.html_to_pdf(
        sales_results_html,
        "output/sales_results.pdf"
    )


def log_out():
    """
    Logs out of the application to properly terminate the session.
    """
    page = browser.page()

    page.click("text=Log out")