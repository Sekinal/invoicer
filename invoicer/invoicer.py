import pathlib
from rich.console import Console
from rich.progress import Progress
from rich.logging import RichHandler
import logging
from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import datetime, timedelta

# Set up logging
log_dir = pathlib.Path('logs')
log_dir.mkdir(exist_ok=True)
log_file = log_dir / 'app.log'
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', filename=log_file, filemode='w')
logger = logging.getLogger(__name__)
handler = RichHandler(show_time=False)
logger.addHandler(handler)

console = Console()

def read_credentials(file_path):
    """
    Reads email and password from a specified file.

    Args:
        file_path (str): Path to the credentials file.

    Returns:
        tuple: A tuple containing email and password.
    """
    email = None
    password = None
    with pathlib.Path(file_path).open('r') as f:
        for line in f:
            line = line.strip()
            if line.startswith('Email: '):
                email = line.split(': ')[1]
            elif line.startswith('Password: '):
                password = line.split(': ')[1]
    if email and password:
        return email, password
    else:
        if not pathlib.Path(file_path).is_file():
            raise ValueError(f"The file {file_path} does not exist.")
        elif email is None:
            raise ValueError(f"The file {file_path} is missing the 'Email: ' line. If the line exists, make sure you left a space after the :.")
        elif password is None:
            raise ValueError(f"The file {file_path} is missing the 'Password: ' line. If the line exists, make sure you left a space after the :.")

def read_data(file_path):
    """
    Reads data from a CSV file into a pandas DataFrame.

    Args:
        file_path (str): Path to the CSV file.

    Returns:
        pandas.DataFrame: DataFrame containing the data from the CSV file.
    """
    try:
        df = pd.read_csv(file_path)
        columns_to_extract = ['Owner Name', 'Invoiceable Hours', 'Invoice Rate', 'Green Waste Number', 'Invoice note']
        for column in columns_to_extract:
            if column not in df.columns:
                raise ValueError(f"The column '{column}' does not exist in the CSV file.")
        df = df[columns_to_extract]
        df = df.where(pd.notnull(df), '')
        return df
    except FileNotFoundError:
        raise ValueError(f"The file {file_path} does not exist.")
    except pd.errors.EmptyDataError:
        raise ValueError(f"The file {file_path} is empty.")
    except pd.errors.ParserError as e:
        raise ValueError(f"Error parsing the file {file_path}: {e}")
    except Exception as e:
        raise ValueError(f"An unexpected error occurred while reading the file {file_path}: {e}")

def initialize_invoicer(email, password, p):
    """
    Initializes the invoicer by logging into the application using Playwright.

    Args:
        email (str): User's email.
        password (str): User's password.
        p (Playwright): Playwright instance.

    Returns:
        tuple: Browser and Page instances.
    """
    try:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        page.goto("https://app.afirmo.com/")
        page.wait_for_timeout(5000)      
        page.locator("#username").fill(email)
        page.locator("#password").fill(password)
        page.get_by_role("button", name="Sign in").click()
        return browser, page
    except Exception as e:
        logger.error(f"Error initializing invoicer: {e}")
        raise

def create_invoice(owner_name, invoiceable_hrs, hourly_rate, green_waste_no, invoice_note, page):
    """
    Creates an invoice for the given owner details.

    Args:
        owner_name (str): Name of the invoice owner.
        invoiceable_hrs (float): Number of invoiceable hours.
        hourly_rate (float): Hourly rate for the invoice.
        green_waste_no (int): Number of green waste units.
        invoice_note (str): Note for the invoice.
        page (Page): Playwright Page instance.
    """
    page.goto("https://app.afirmo.com/sales/invoices/add")
    page.wait_for_timeout(1000)
    page.get_by_placeholder("Search contacts").fill(owner_name)
    page.wait_for_timeout(1000)
    page.get_by_role("button", name=owner_name).click()
    
    date_obj = datetime.strptime(page.locator("#date").input_value(), '%d/%m/%Y')
    two_weeks_later = date_obj + timedelta(days=14)
    new_date_str = two_weeks_later.strftime('%d/%m/%Y')
    
    page.locator("#expiryDate").fill('')
    page.locator("#expiryDate").fill(str(new_date_str))
    page.get_by_placeholder("Select product").click()
    page.get_by_role("button", name="Gardening- hourly").click()
    page.locator("#items\[0\]\.qty").fill('')
    page.locator("#items\[0\]\.qty").fill(str(invoiceable_hrs))
    page.locator("#items\[0\]\.unitAmount").fill('')
    page.locator("#items\[0\]\.unitAmount").fill(str(hourly_rate))   
    
    if green_waste_no and green_waste_no != '' and green_waste_no != 0:
        page.get_by_role("button", name="new row").click()
        page.get_by_placeholder("Select product").click()
        page.get_by_role("button", name="Green waste removal").click()
        page.locator("#items\[1\]\.qty").fill('')
        page.locator("#items\[1\]\.qty").fill(str(green_waste_no))
    
    if invoice_note:
        page.get_by_role("button", name="new Notes").click()
        page.locator("input[name=\"inotes\\[0\\]\\.text\"]").fill(str(invoice_note))
    
    page.wait_for_timeout(1000)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(5000)
    
def main():
    console.print("Invoicer App", style="blue")

    credentials_file_path = input("Enter the path to the credentials file: ")
    email, password = read_credentials(credentials_file_path)

    data_file_path = input("Enter the path to the data CSV file: ")
    df = read_data(data_file_path)
    print(df)
    failed_owners = []

    with Progress() as progress:
        task = progress.add_task("[red]Creating invoices...", total=len(df))
        
        with sync_playwright() as p:
            browser = None
            page = None
            
            for _, row in df.iterrows():
                owner_name = row['Owner Name']
                invoiceable_hrs = row['Invoiceable Hours']
                hourly_rate = row['Invoice Rate']
                green_waste_no = row['Green Waste Number']
                invoice_note = row['Invoice note']

                try:
                    if browser is None or page is None:
                        browser, page = initialize_invoicer(email, password, p)
                    
                    create_invoice(owner_name, invoiceable_hrs, hourly_rate, green_waste_no, invoice_note, page)
                except Exception as e:
                    logger.error(f"Error creating invoice for {owner_name}: {e}")
                    console.print(f"Error creating invoice for {owner_name}", style="red")
                    failed_owners.append(owner_name)
                    
                    if browser:
                        browser.close()
                    browser = None
                    page = None

                progress.update(task, advance=1)

            if browser:
                browser.close()

        if failed_owners:
            console.print("Failed to create invoices for:", style="red")
            for owner in failed_owners:
                console.print(f"  - {owner}", style="red")

if __name__ == "__main__":
    main()