import json
import tkinter as tk
from datetime import datetime
from tkinter.filedialog import *
import openpyxl
import shutil
import os
import requests
import cred


def read_currencies_from_file():
    file_path = "currencies.json"
    data = {}

    try:
        with open(file_path, 'r') as file:
            data = json.load(file)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"Error reading JSON file: {e}")

    return data


def write_currencies_to_file(data):
    with open("currencies.json", "w") as file:
        json.dump(data, file, indent=4)


def get_file():
    filename = askopenfilename()
    if filename:
        return filename


currency_adjustments = read_currencies_from_file()


def write_excel(market_de_de, market_de_eu, market_de_ch, seller_de_de, seller_de_eu, total, name):
    # Dialog zur Auswahl der Excel-Datei öffnen
    # file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    # Specify the path to the Downloads folder
    downloads_folder = os.path.expanduser("~/Downloads")

    # Check if the Downloads folder exists
    if not os.path.exists(downloads_folder):
        print("Downloads folder not found. Make sure it exists.")
        return

    # Specify the full path for the copy
    copy_file_path = os.path.join(downloads_folder, f"{name}.xlsx")

    # Split the input string by underscore
    parts = name.split('_')
    firm_name = ""
    date = ""
    if len(parts) == 2:
        firm_name = parts[0]
        date = parts[1]
    else:
        print("Invalid input format. Expected 'Firmename_Month Year'.")

    try:
        # Check if the copy file already exists
        if os.path.exists(copy_file_path):
            os.remove(copy_file_path)  # Delete the existing copy

        # Create a copy of the Excel file
        source_file = os.path.join(downloads_folder, "Vorlage_Umsätze Amazon Umsatzsteuer_leer.xlsx")
        shutil.copyfile(source_file, copy_file_path)

        # Open the Excel file
        copy_workbook = openpyxl.load_workbook(copy_file_path)
        copy_sheet = copy_workbook.active

        # Insert data into the copy
        copy_sheet['E8'] = market_de_de
        copy_sheet['E9'] = market_de_eu
        copy_sheet['E10'] = market_de_ch
        copy_sheet['G16'] = seller_de_de
        copy_sheet['E24'] = seller_de_eu
        copy_sheet['C69'] = total
        copy_sheet['A1'] = firm_name
        copy_sheet['D2'] = "Steuermeldung " + date

        # Save the copy
        copy_workbook.save(copy_file_path)

        print(f"Copy created and data written to: {copy_file_path}")
    except Exception as e:
        print(f"Error copying and modifying the Excel file: {e}")


def main():
    filename = get_file()

    if filename:
        with open(filename, 'r', encoding="utf-8") as file:
            lines = file.read().splitlines()
            if lines:
                headers = lines[0].split('\t')  # Split the first line into headers using tabs
                data = []
                for line in lines[1:]:
                    values = line.split('\t')
                    row_dict = dict(zip(headers, values))
                    data.append(row_dict)

                # Initialize a counter to start from 1
                count = 1

                # Display BUYER_VAT_NUMBER values with numbers
                for row in data:
                    if (row.get('TAX_COLLECTION_RESPONSIBILITY') == 'SELLER' and
                            row.get('SALE_ARRIVAL_COUNTRY') != 'DE' and
                            row.get('SALE_DEPART_COUNTRY') == 'DE'):
                        buyer_vat_number = row.get('BUYER_VAT_NUMBER')
                        print(f"{count}:\t{buyer_vat_number}")
                        count += 1

                # Let the user choose which numbers to change
                user_input = input(
                    "Die Nummern mit falschen ID Komma getrennt angeben, oder Enter zum weiterfahren: ")
                numbers_to_change = [int(num.strip()) for num in user_input.split(",") if num.strip().isdigit()]

                count = 1  # Reset the counter

                for row in data:
                    if (row.get('TAX_COLLECTION_RESPONSIBILITY') == 'SELLER' and
                            row.get('SALE_ARRIVAL_COUNTRY') != 'DE' and
                            row.get('SALE_DEPART_COUNTRY') == 'DE'):
                        if count in numbers_to_change:
                            row['TAX_COLLECTION_RESPONSIBILITY'] = 'MARKETPLACE'
                        count += 1


                market_de_de = [row for row in data if
                                row.get('TAX_COLLECTION_RESPONSIBILITY') == 'MARKETPLACE' and
                                row.get('SALE_ARRIVAL_COUNTRY') == 'DE' and
                                row.get('SALE_DEPART_COUNTRY') == 'DE']

                market_de_eu = [row for row in data if
                                row.get('TAX_COLLECTION_RESPONSIBILITY') == 'MARKETPLACE' and
                                row.get('SALE_ARRIVAL_COUNTRY') not in ('DE', 'CH') and
                                row.get('SALE_DEPART_COUNTRY') == 'DE']

                market_de_ch = [row for row in data if
                                row.get('TAX_COLLECTION_RESPONSIBILITY') in ('MARKETPLACE', 'SELLER') and
                                row.get('SALE_ARRIVAL_COUNTRY') == 'CH' and
                                row.get('SALE_DEPART_COUNTRY') == 'DE']

                seller_de_de = [row for row in data if
                                row.get('TAX_COLLECTION_RESPONSIBILITY') == 'SELLER' and
                                row.get('SALE_ARRIVAL_COUNTRY') == 'DE' and
                                row.get('SALE_DEPART_COUNTRY') == 'DE']

                seller_de_eu = [row for row in data if
                                row.get('TAX_COLLECTION_RESPONSIBILITY') == 'SELLER' and
                                row.get('SALE_ARRIVAL_COUNTRY') not in ('DE', 'CH') and
                                row.get('SALE_DEPART_COUNTRY') == 'DE']

                total = [row for row in data if
                         row.get('SALE_DEPART_COUNTRY') == 'DE']

                market_de_de = get_total_from_list(market_de_de)
                market_de_eu = get_total_from_list(market_de_eu)
                market_de_ch = get_total_from_list(market_de_ch)
                seller_de_de = get_total_from_list(seller_de_de)
                seller_de_eu = get_total_from_list(seller_de_eu)
                total = get_total_from_list(total)

                print(f"Market DE -> DE: {market_de_de}")
                print(f"Market DE -> EU: {market_de_eu}")
                print(f"Market DE -> CH: {market_de_ch}")
                print(f"Seller DE -> DE: {seller_de_de}")
                print(f"Seller DE -> EU: {seller_de_eu}")

                total_sum = market_de_de + market_de_eu + market_de_ch + seller_de_de + seller_de_eu

                print(f"Total: {total}")
                print(f"Total: {total_sum}")

                write_currencies_to_file(currency_adjustments)
                write_excel(market_de_de, market_de_eu, market_de_ch, seller_de_de, seller_de_eu, total, os.path.basename(filename)[:len(os.path.basename(filename)) - 4])


def get_total_from_list(data):
    total_sum = 0
    for row in data:
        value = row.get("TOTAL_ACTIVITY_VALUE_AMT_VAT_INCL")
        if value is not None and value.strip() != "":
            currency_code = row.get('TRANSACTION_CURRENCY_CODE')
            date_long = row.get('TRANSACTION_COMPLETE_DATE')
            date_object = datetime.strptime(date_long, "%d-%m-%Y")
            date = date_object.strftime("%Y-%m-%d")
            total_value = float(value)

            # Check if the currency code is not EUR or CHF
            if currency_code not in 'EUR' and row.get("TRANSACTION_COMPLETE_DATE") != "":
                if date not in currency_adjustments:
                    # Ask the user for an adjustment value
                    # user_input = input(f"Enter Kurs für {currency_code} am {date} : ")
                    rate = get_currency_rate(date, currency_code)
                    print(f"Rate for  {currency_code} at {date} is {rate}")
                    total_value /= rate
                else:
                    print(f"Rate for  {currency_code} at {date} is {currency_adjustments.get(date, {}).get('rates', {}).get(currency_code, None)}")
                    total_value /= currency_adjustments.get(date, {}).get('rates', {}).get(currency_code, None)

            total_sum += total_value

    return total_sum if total_sum is not None else 0


def get_currency_rate(date, currency_code):
    # Set API Endpoint and API key
    access_key = cred.api_key

    # Construct the API URL
    url = f'http://api.exchangeratesapi.io/v1/{date}?access_key={access_key}'

    try:
        # Send a GET request to the API endpoint
        response = requests.get(url)
        response.raise_for_status()  # Raise an exception for HTTP errors

        # Parse the JSON response
        exchange_rates = response.json()

        currency_adjustments[date] = {'rates': exchange_rates['rates']}
        return exchange_rates['rates'][currency_code]
    except requests.exceptions.RequestException as e:
        print(f"Error making the request: {e}")


if __name__ == "__main__":
    root = tk.Tk()  # Create a tkinter window
    root.withdraw()  # Hide the main window

    main()  # Run the main function
    # get_currency_rate("2023","01","CHF")
    root.destroy()  # Close the tkinter window
