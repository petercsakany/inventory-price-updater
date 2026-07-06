import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd


def select_file(title_text, file_types):
    """Opens a native file selection window."""
    root = tk.Tk()
    root.withdraw()  # Hide the main tiny tkinter window
    root.attributes("-topmost", True)  # Bring the dialog to the front
    file_path = filedialog.askopenfilename(title=title_text, filetypes=file_types)
    root.destroy()
    return file_path


def parse_description(description):
    """Parses the Week 22 description column based on retail string structures.

    Root: First 7 digits
    SV: Last 1 or 2 digits of the 3-digit code after the root
    Product: Remaining text excluding the last 6 trailing digits
    """
    try:
        desc_str = str(description).strip()

        # Extract Root (First 7 digits)
        root = desc_str[:7]

        # Extract SV (Look at the 3-digit part right after the root at index 8:11)
        three_digit_part = desc_str[8:11]
        # Convert to integer to strip leading zeros naturally (e.g., "000" -> "0", "002" -> "2")
        sv = str(int(three_digit_part))

        # Extract Product (Everything following the 3-digit section, minus the last 6 digits)
        product = desc_str[12:-6].strip()

        return root, sv, product
    except Exception:
        # Fallback if a line doesn't match standard formatting criteria
        return "UNKNOWN", "0", desc_str


def main():
    # ----------------------------------------------------
    # 1. FILE SELECTION DIALOGS
    # ----------------------------------------------------
    excel_types = [("Excel Files", "*.xlsx *.xls")]

    print("Opening file window for Price Monitor...")
    monitor_path = select_file("Select the Price Monitor File (e.g. Week 22)", excel_types)
    if not monitor_path:
        print("Operation cancelled: No price monitor file selected.")
        return

    print("Opening file window for Master Inventory...")
    inventory_path = select_file("Select the Master Inventory File", excel_types)
    if not inventory_path:
        print("Operation cancelled: No inventory file selected.")
        return

    try:
        # ----------------------------------------------------
        # 2. LOAD DATA
        # ----------------------------------------------------
        print("Loading sheets into memory...")
        # Treat barcodes strictly as text strings to prevent dropped leading zeros
        df_monitor = pd.read_excel(monitor_path, dtype={"Barcode": str})
        df_inventory = pd.read_excel(inventory_path, dtype={"Barcode": str})

        # Ensure correct formatting and handle potential missing values
        df_monitor["Barcode"] = df_monitor["Barcode"].str.strip()
        df_inventory["Barcode"] = df_inventory["Barcode"].str.strip()

        # ----------------------------------------------------
        # 3. PROCESSING & COMPARISON LOGIC
        # ----------------------------------------------------
        updated_prices_count = 0
        added_products = []

        # Create an inventory lookup map for lightning-fast matching: {barcode: row_index}
        inventory_lookup = {
            row["Barcode"]: idx for idx, row in df_inventory.iterrows() if pd.notna(row["Barcode"])
        }

        # Loop through every row in the update sheet
        for _, row in df_monitor.iterrows():
            barcode = row.get("Barcode")
            new_price = row.get("New Price")
            description = row.get("Description")

            # Basic verification checks
            if pd.isna(barcode) or barcode == "" or pd.isna(new_price):
                continue

            try:
                new_price = float(new_price)
            except ValueError:
                continue

            # Case A: Product exists in inventory -> Check price threshold
            if barcode in inventory_lookup:
                inv_row_idx = inventory_lookup[barcode]
                current_price = df_inventory.at[inv_row_idx, "Price"]

                try:
                    current_price = float(current_price) if pd.notna(current_price) else 0.0
                except ValueError:
                    current_price = 0.0

                # ONLY update if the new monitor price is strictly higher
                if new_price > current_price:
                    df_inventory.at[inv_row_idx, "Price"] = new_price
                    updated_prices_count += 1

            # Case B: Product cannot be found -> Add parsing fields and append
            else:
                root_val, sv_val, product_name = parse_description(description)

                new_item = {
                    "Root": root_val,
                    "SV": sv_val,
                    "Product": product_name,
                    "Barcode": barcode,
                    "Price": new_price,
                }

                # Append to our local dataframe layout tracking
                df_inventory = pd.concat([df_inventory, pd.DataFrame([new_item])], ignore_index=True)
                # Keep tracking metadata for the completion dialog box
                added_products.append(f"• {product_name} (Barcode: {barcode}, Price: €{new_price:.2f})")
                # Update loop mapping dict tracking so duplicates in the monitor don't add twice
                inventory_lookup[barcode] = len(df_inventory) - 1

        # ----------------------------------------------------
        # 4. SAVE EXCEL FILE OUTPUT
        # ----------------------------------------------------
        # Create output file path alongside your current master file
        dir_name = os.path.dirname(inventory_path)
        output_path = os.path.join(dir_name, "updated_inventory_output.xlsx")

        print(f"Writing outputs to file destination: {output_path}")
        df_inventory.to_excel(output_path, index=False)

        # ----------------------------------------------------
        # 5. CONSTRUCT COMPLETION DIALOG REPORT
        # ----------------------------------------------------
        report_title = "Inventory Update Completed!"

        report_message = f"Process finished successfully!\n\n"
        report_message += f"✔ Prices updated (higher values found): {updated_prices_count}\n"
        report_message += f"✔ New products appended to file: {len(added_products)}\n\n"

        if added_products:
            report_message += "Newly Added Items Breakdown:\n"
            # Show up to the first 15 entries so the dialog box fits the screen nicely
            report_message += "\n".join(added_products[:15])
            if len(added_products) > 15:
                report_message += f"\n...and {len(added_products) - 15} more items."
        else:
            report_message += "No new items needed to be added to the catalog index."

        # Show final popup message to user
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showinfo(report_title, report_message)
        root.destroy()

    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Error Occurred", f"An unexpected error blocked execution:\n{str(e)}")
        root.destroy()


if __name__ == "__main__":
    main()