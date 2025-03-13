# user_interface.py
import sys
import os
import pandas as pd
from typing import List, Optional


def prompt_for_sales_person(config):
    sales_people = config.sales_people
    print("\n1. Sales Person:")
    for idx, person in enumerate(sales_people, 1):
        print(f"   [{idx}] {person}")
    while True:
        try:
            choice = int(input("\nSelect sales person (enter number): "))
            if 1 <= choice <= len(sales_people):
                return sales_people[choice - 1]
            print(f"Enter a number between 1 and {len(sales_people)}.")
        except ValueError:
            print("Enter a valid number.")


def prompt_for_billing_type():
    print("\n2. Billing Type:")
    print("   [C] Calendar")
    print("   [B] Broadcast")
    while True:
        billing_input = input("\nSelect billing type (C/B): ").strip().upper()
        if billing_input in ["C", "B"]:
            return "Calendar" if billing_input == "C" else "Broadcast"
        print("Enter 'C' or 'B'.")


def prompt_for_revenue_type():
    print("\n3. Revenue Type:")
    print("   [B] Branded Content")
    print("   [D] Direct Response Sales")
    print("   [I] Internal Ad Sales")
    print("   [P] Paid Programming")
    print("   [T] Trade")
    revenue_types = {
        "B": "Branded Content",
        "D": "Direct Response Sales",
        "I": "Internal Ad Sales",
        "P": "Paid Programming",
        "T": "Trade",
    }
    while True:
        revenue_input = input("\nSelect revenue type (B/D/I/P/T): ").strip().upper()
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
        if order_input in ["A", "N", "T"]:
            order_types = {"A": "Agency", "N": "Non-Agency", "T": "Trade"}
            if order_input == "A":
                print("\n5. Agency Fee Type:")
                print("   [S] Standard (15%)")
                print("   [C] Custom")
                while True:
                    fee_type = input("\nSelect fee type (S/C): ").strip().upper()
                    if fee_type == "S":
                        agency_fee = 0.15
                        break
                    elif fee_type == "C":
                        while True:
                            try:
                                custom_fee = float(
                                    input("\nEnter custom fee percentage (without %): ")
                                )
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
    return input(
        "\nWhat is the estimate number? (Optional, press Enter to skip): "
    ).strip()


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
        if affidavit in ["Y", "N"]:
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
    Display all unique descriptions with their detected languages and allow pattern-based corrections.
    Returns the corrected language Series.
    """
    detected_counts, row_languages = language_info
    print("\n" + "-" * 80)
    print("Language Detection Results".center(80))
    print("-" * 80)
    
    # Show summary counts - SORTED by language code
    for lang_code, count in sorted(detected_counts.items()):
        print(f"   • {lang_code}: {count} entries")
    
    # Group by unique descriptions and show their detected languages
    unique_descriptions = {}
    for idx, desc in df["rowdescription"].items():
        if desc not in unique_descriptions:
            unique_descriptions[desc] = {
                "language": row_languages.loc[idx],
                "count": 0,
                "indices": []
            }
        unique_descriptions[desc]["count"] += 1
        unique_descriptions[desc]["indices"].append(idx)
    
    print(f"\nFound {len(unique_descriptions)} unique line descriptions:")
    print("-" * 80)
    
    # Display all unique descriptions with their detected language, sorted by language code
    sorted_descriptions = sorted(unique_descriptions.items(), key=lambda x: x[1]['language'])
    for i, (desc, info) in enumerate(sorted_descriptions, 1):
        print(f"{i:2d}. [{info['language']}] {desc} ({info['count']} occurrences)")
    
    print("\nDoes this look correct? (Y/N)")
    if input().strip().lower() == "n":
        print("\nAvailable language codes: E, M, T, Hm, SA, V, C, K, J")
        print("\nYou can correct languages in several ways:")
        print("1. Fix specific line descriptions")
        print("2. Pattern-based correction")
        
        choice = input("\nSelect correction method (1/2): ").strip()
        
        if choice == "1":
            # Correct by specific line description
            while True:
                try:
                    line_num = input("\nEnter line number to change language (or press Enter to finish): ").strip()
                    if not line_num:
                        break
                    
                    line_num = int(line_num)
                    if 1 <= line_num <= len(unique_descriptions):
                        # Get the description and current language
                        desc = list(unique_descriptions.keys())[line_num - 1]
                        current_lang = unique_descriptions[desc]["language"]
                        print(f"Current: [{current_lang}] {desc}")
                        
                        # Get new language
                        new_lang = input("Enter new language code: ").strip().upper()
                        if new_lang in ["E", "M", "T", "Hm", "SA", "V", "C", "K", "J"]:
                            # Update all matching rows
                            indices = unique_descriptions[desc]["indices"]
                            for idx in indices:
                                row_languages.loc[idx] = new_lang
                            
                            # Update our tracking dictionary
                            unique_descriptions[desc]["language"] = new_lang
                            count = unique_descriptions[desc]["count"]
                            print(f"Updated {count} occurrences of '{desc}' to {new_lang}")
                        else:
                            print("Invalid language code")
                    else:
                        print(f"Please enter a number between 1 and {len(unique_descriptions)}")
                except ValueError:
                    print("Invalid input. Please enter a number.")
        
        elif choice == "2":
            # Pattern-based correction
            while True:
                pattern = input("\nEnter text pattern to match (or press Enter to finish): ").strip()
                if not pattern:
                    break
                
                # Find matching descriptions
                matches = [desc for desc in unique_descriptions.keys() 
                          if pattern.lower() in desc.lower()]
                
                if not matches:
                    print(f"No descriptions contain '{pattern}'")
                    continue
                
                # Show matching descriptions
                print(f"\nFound {len(matches)} matching descriptions:")
                for i, desc in enumerate(matches, 1):
                    lang = unique_descriptions[desc]["language"]
                    count = unique_descriptions[desc]["count"]
                    print(f"{i:2d}. [{lang}] {desc} ({count} occurrences)")
                
                # Get target language
                new_lang = input("\nSet these to which language code? ").strip().upper()
                if new_lang in ["E", "M", "T", "Hm", "SA", "V", "C", "K", "J"]:
                    # Apply changes to all matching descriptions
                    total_updated = 0
                    for desc in matches:
                        indices = unique_descriptions[desc]["indices"]
                        for idx in indices:
                            row_languages.loc[idx] = new_lang
                        
                        # Update tracking dictionary
                        count = unique_descriptions[desc]["count"]
                        unique_descriptions[desc]["language"] = new_lang
                        total_updated += count
                    
                    print(f"Updated {total_updated} total occurrences across {len(matches)} descriptions")
                else:
                    print("Invalid language code, skipping this pattern")
    
    # Show summary of changes - SORTED by language code
    print("\nUpdated language distribution:")
    updated_counts = row_languages.value_counts().to_dict()
    for lang_code, count in sorted(updated_counts.items()):
        print(f"   • {lang_code}: {count} entries")
    
    return row_languages

def print_header(log_file):
    header = f"""
    ╔════════════════════════════════════════════════════════════════════════════╗
    ║                        Excel File Processing Tool                           ║
    ╚════════════════════════════════════════════════════════════════════════════╝

    Version: 2.0
    Log File: {log_file}
    """
    print(header)


def select_processing_mode() -> str:
    """Ask the user whether to process all files or select one at a time."""
    print("\n" + "-" * 80)
    print("Processing Mode Selection".center(80))
    print("-" * 80)
    print("\nChoose how you want to process your files:")
    print("  [A] Process all files automatically")
    print("  [S] Select and process files one at a time")

    while True:
        choice = input("\nYour choice (A/S): ").strip().upper()
        if choice in ["A", "S"]:
            return choice
        print(
            "❌ Invalid choice. Please enter 'A' for all files or 'S' to select files."
        )


def display_batch_summary(
    successful: List["ProcessingResult"],  # forward reference as a string
    failed: List["ProcessingResult"],
    log_file: str,
):

    print("\n" + "=" * 80)
    print("Batch Processing Summary".center(80))
    print("=" * 80)

    total = len(successful) + len(failed)
    success_rate = (len(successful) / total * 100) if total > 0 else 0

    print(f"\nTotal files processed: {total}")
    print(f"Successfully processed: {len(successful)} ({success_rate:.1f}%)")
    print(f"Failed to process: {len(failed)}")

    if failed:
        print("\nFailed Files:")
        for result in failed:
            print(f"❌ {result.filename}")
            print(f"   Error: {result.error_message}")

    if any(r.warnings for r in successful):
        print("\nWarnings:")
        for result in successful:
            if result.warnings:
                print(f"⚠️ {result.filename}:")
                for warning in result.warnings:
                    print(f"   - {warning}")
    if successful:
        print("\nProcessed Files:")
        for result in successful:
            print(f"✅ {result.filename} -> {result.output_file}")

    print(f"\nDetailed logs available at: {log_file}")


def choose_input_file(files: List[str], input_dir: str) -> Optional[str]:
    """Prompt the user to select a file from the input directory."""
    print("\n" + "-" * 80)
    print("File Selection".center(80))
    print("-" * 80)
    print("\nAvailable files for processing:")

    # Create two columns if there are many files
    mid_point = (len(files) + 1) // 2
    for i, filename in enumerate(files, 1):
        line = f"  [{i:2d}] {filename}"
        if i <= mid_point and i + mid_point <= len(files):
            second_file = files[i + mid_point - 1]
            second_item = f"  [{i + mid_point:2d}] {second_file}"
            print(f"{line:<40} {second_item}")
        else:
            print(line)

    while True:
        try:
            choice = input(
                "\nEnter the number of the file you want to process (or 'q' to quit): "
            ).strip()
            if choice.lower() == "q":
                print("\nExiting program...")
                sys.exit(0)
            choice = int(choice)
            if 1 <= choice <= len(files):
                selected_file = files[choice - 1]
                print(f"\n✅ Selected: {selected_file}")
                return os.path.join(input_dir, selected_file)
            else:
                print(f"❌ Please enter a number between 1 and {len(files)}")
        except ValueError:
            print("❌ Please enter a valid number or 'q' to quit")


def prompt_batch_settings(config) -> dict:
    """
    Collects batch-specific settings from the user and returns a dictionary
    of shared inputs for the batch. If the batch is for WorldLink orders, returns
    the default settings.
    """
    print("\n" + "-" * 80)
    print("Batch Processing Setup".center(80))
    print("-" * 80)

    settings = {}
    is_worldlink = (
        input("\nIs this a batch of WorldLink orders? (Y/N): ").strip().lower() == "y"
    )
    settings["is_worldlink"] = is_worldlink

    if is_worldlink:
        print("\nUsing WorldLink default settings...")
        # Here we return an empty dict or a flag; the caller can then use get_worldlink_defaults.
        settings["use_defaults"] = True
    else:
        shared_inputs = (
            input("\nDo all files in this batch share the same user inputs? (Y/N): ")
            .strip()
            .lower()
        )
        if shared_inputs == "y":
            print("\nCollecting shared user inputs for the batch...")
            # Collect shared inputs using the existing collect_user_inputs function.
            settings["inputs"] = collect_user_inputs(config)
        else:
            settings["inputs"] = None
    return settings
