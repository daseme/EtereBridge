# user_interface.py
import sys
import pandas as pd


def prompt_for_sales_person(config):
    sales_people = config.sales_people
    print("\n1. Sales Person:")
    for idx, person in enumerate(sales_people, 1):
        print(f"   [{idx}] {person}")
    while True:
        try:
            choice = int(input("\nSelect sales person (enter number): "))
            if 1 <= choice <= len(sales_people):
                return sales_people[choice-1]
            print(f"Enter a number between 1 and {len(sales_people)}.")
        except ValueError:
            print("Enter a valid number.")

def prompt_for_billing_type():
    print("\n2. Billing Type:")
    print("   [C] Calendar")
    print("   [B] Broadcast")
    while True:
        billing_input = input("\nSelect billing type (C/B): ").strip().upper()
        if billing_input in ['C', 'B']:
            return "Calendar" if billing_input == 'C' else "Broadcast"
        print("Enter 'C' or 'B'.")

def prompt_for_revenue_type():
    print("\n3. Revenue Type:")
    print("   [B] Branded Content")
    print("   [D] Direct Response Sales")
    print("   [I] Internal Ad Sales")
    print("   [P] Paid Programming")
    revenue_types = {'B': "Branded Content", 'D': "Direct Response Sales", 
                     'I': "Internal Ad Sales", 'P': "Paid Programming"}
    while True:
        revenue_input = input("\nSelect revenue type (B/D/I/P): ").strip().upper()
        if revenue_input in revenue_types:
            return revenue_types[revenue_input]
        print("Enter B, D, I, or P.")

def prompt_for_order_type():
    print("\n4. Order Type:")
    print("   [A] Agency")
    print("   [N] Non-Agency")
    print("   [T] Trade")
    agency_fee = None
    while True:
        order_input = input("\nSelect order type (A/N/T): ").strip().upper()
        if order_input in ['A', 'N', 'T']:
            order_types = {'A': "Agency", 'N': "Non-Agency", 'T': "Trade"}
            if order_input == 'A':
                print("\n5. Agency Fee Type:")
                print("   [S] Standard (15%)")
                print("   [C] Custom")
                while True:
                    fee_type = input("\nSelect fee type (S/C): ").strip().upper()
                    if fee_type == 'S':
                        agency_fee = 0.15
                        break
                    elif fee_type == 'C':
                        while True:
                            try:
                                custom_fee = float(input("\nEnter custom fee percentage (without %): "))
                                if 0 <= custom_fee <= 100:
                                    agency_fee = custom_fee / 100
                                    break
                                print("Enter a percentage between 0 and 100.")
                            except ValueError:
                                print("Enter a valid number.")
                        break
                    print("Enter 'S' or 'C'.")
            return order_types[order_input], agency_fee
        print("Enter A, N, or T.")

def prompt_for_estimate():
    """Prompt user for estimate number (optional). Returns estimate as string (empty if not provided)."""
    return input("\nWhat is the estimate number? (Optional, press Enter to skip): ").strip()

def prompt_for_type(config):
    print("\nType Selection:")
    for idx, type_opt in enumerate(config.type_options, 1):
        print(f"   [{idx}] {type_opt}")
    while True:
        try:
            choice = int(input("\nSelect type (enter number): ").strip())
            if 1 <= choice <= len(config.type_options):
                return config.type_options[choice - 1]
            print("Enter a number between 1 and", len(config.type_options))
        except ValueError:
            print("Enter a valid number.")

def prompt_for_affidavit():
    while True:
        affidavit = input("\nIs this an affidavit? (Y/N): ").strip().upper()
        if affidavit in ['Y', 'N']:
            return affidavit
        print("Enter 'Y' or 'N'.")

def prompt_for_contract():
    """
    Prompt the user to enter a contract number.
    """
    while True:
        contract = input("\nPlease enter the contract number: ").strip()
        if contract:
            return contract
        print("Contract number cannot be empty. Please enter a valid contract number.")


def collect_user_inputs(config):
    """Collect and return all user inputs as a dictionary."""
    inputs = {}
    inputs["sales_person"] = prompt_for_sales_person(config)
    inputs["billing_type"] = prompt_for_billing_type()
    inputs["revenue_type"] = prompt_for_revenue_type()
    order_type, agency_fee = prompt_for_order_type()
    inputs["agency_flag"] = order_type
    inputs["agency_fee"] = agency_fee
    inputs["estimate"] = prompt_for_estimate()
    inputs["contract"] = prompt_for_contract()  # New prompt for estimate number
    # Type selection is now automatic—no prompt needed.
    inputs["affidavit"] = prompt_for_affidavit()
    return inputs


def verify_languages(df: pd.DataFrame, language_info):
    """
    Display detected languages and sample entries; return the language Series.
    """
    detected_counts, row_languages = language_info
    print("\n" + "-"*80)
    print("Language Detection Results".center(80))
    print("-"*80)
    for lang_code, count in detected_counts.items():
        print(f"   • {lang_code}: {count} entries")
    print("\nSample entries:")
    for lang_code in detected_counts:
        rows = df[row_languages == lang_code]
        if not rows.empty:
            print(f"\n{lang_code}:")
            samples = rows['rowdescription'].head(2)
            for desc in samples:
                print(f"   • {desc}")
    print("\nDoes this look correct? (Y/N)")
    if input().strip().lower() == 'n':
        print("\nAvailable language codes: E, M, T, Hm, SA, V, C, K, J")
        while True:
            try:
                row_num = int(input("\nEnter row number to change language, or press Enter to continue: ").strip() or -1)
                if row_num == -1:
                    break
                if 0 <= row_num < len(df):
                    print(f"Current: {df.iloc[row_num]['rowdescription']}")
                    new_lang = input("Enter new language code: ").strip().upper()
                    if new_lang in ['E','M','T','Hm','SA','V','C','K','J']:
                        row_languages.iloc[row_num] = new_lang
                else:
                    print("Row number out of range.")
            except ValueError:
                print("Invalid input.")
    return row_languages
