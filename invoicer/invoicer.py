import pathlib
from rich.console import Console
from rich.progress import Progress
from rich.logging import RichHandler
import logging
from playwright.sync_api import sync_playwright
import pandas as pd

# Set up logging
log_dir = pathlib.Path('logs')
if not log_dir.exists():
    log_dir.mkdir()
log_file = log_dir / 'app.log'
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', filename=log_file, filemode='w')
logger = logging.getLogger(__name__)
handler = RichHandler(show_time=False)
logger.addHandler(handler)

console = Console()

def read_credentials(file_path):
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
    try:
        # Read the CSV file into a pandas DataFrame
        df = pd.read_csv(file_path)
        
        # Specify the columns to extract
        columns_to_extract = ['Owner Name', 'Invoiceable Hours', 'Invoice Rate', 'Green Waste Number', 'Invoice note']
        
        # Check if all columns exist in the DataFrame
        for column in columns_to_extract:
            if column not in df.columns:
                raise ValueError(f"The column '{column}' does not exist in the CSV file.")
        
        # Reorder the columns to put Owner Name first
        df = df[columns_to_extract]
        
        # Replace NaN values with None
        df = df.where(pd.notnull(df), '')
        
        # Return the extracted DataFrame
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
    try:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        page.goto("https://app.afirmo.com/")
        page.wait_for_timeout(5000)      
        # Interact with login form
        page.locator("#username").fill(email)
        page.locator("#password").fill(password)
        
        page.get_by_role("button", name="Sign in").click()

        return browser, page
    except Exception as e:
        logger.error(f"Error initializing invoicer: {e}")
        raise

def create_invoice(owner_name, invoiceable_hrs, hourly_rate, green_waste_no, invoice_note, page):
    page.goto("https://app.afirmo.com/sales/invoices/add")
    page.wait_for_timeout(1000)
    page.get_by_placeholder("Search contacts").fill(owner_name)
    page.get_by_role("button", name=owner_name).click()
    page.get_by_placeholder("Select product").click()
    page.get_by_role("button", name="Gardening- hourly").click()
    page.locator("#items\[0\]\.qty").fill('')
    page.locator("#items\[0\]\.qty").fill(str(invoiceable_hrs))
    page.locator("#items\[0\]\.unitAmount").fill('')
    page.locator("#items\[0\]\.unitAmount").fill(str(hourly_rate))   
    
    if green_waste_no is not None and green_waste_no != '' and green_waste_no != 0:
        page.get_by_role("button", name="new row").click()
        page.get_by_placeholder("Select product").click()
        page.get_by_role("button", name="Green waste removal").click()
        page.locator("#items\[1\]\.qty").fill('')
        page.locator("#items\[1\]\.qty").fill(str(green_waste_no))
    
    if invoice_note is not None and invoice_note != '':
        page.get_by_role("button", name="new Notes").click()
        page.locator("input[name=\"inotes\\[0\\]\\.text\"]").fill(str(invoice_note))
    
    page.wait_for_timeout(1000)
    page.query_selector('xpath=/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div/div/div/form/div[1]/div[2]/button').click()
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
                    
                    # Close the current browser session and set to None to reinitialize on next iteration
                    if browser:
                        browser.close()
                    browser = None
                    page = None

                progress.update(task, advance=1)

            # Close the browser if it's still open
            if browser:
                browser.close()

        if failed_owners:
            console.print("Failed to create invoices for:", style="red")
            for owner in failed_owners:
                console.print(f"  - {owner}", style="red")

if __name__ == "__main__":
    main()