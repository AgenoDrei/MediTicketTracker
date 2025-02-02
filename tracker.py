import sys
import os
import mechanicalsoup
import pandas as pd
import datetime
import multiprocessing
from multiprocessing import Pool, freeze_support



class TrackerProcessor():
    def __init__(self, excel_file, code_column_name):
        self.excel_file = excel_file
        self.code_column_name = code_column_name
        self.data_sheet_name = "Daten"
        self.codes_sheet_name = "Freie Codes"
        self.base_url = "https://ticket.medimeisterschaften.com/"
        self.exp_url = "https://ticket.medimeisterschaften.com/?voucher_invalid"
        self.code_col = 'Code'
        self.num_threads = 8

        data_df, codes_df = self._load_data()
        self.data_df = data_df
        self.codes_df = codes_df
        

    def _load_data(self):
        """Load Excel sheets into Pandas DataFrames."""
        data_df = pd.read_excel(self.excel_file, sheet_name=self.data_sheet_name)
        codes_df = pd.DataFrame(columns=["ID", "Vorname", "Nachname", "Email", "Code", "Status"])  # Clear the codes sheet
        
        assert not data_df.empty, "Data sheet not found"
        return data_df, codes_df

    def process_vouchers_parallel(self, gui_callback):
        """Process voucher codes in parallel using multiprocessing."""
        self.pool = Pool(processes=self.num_threads)
            
        args = [(row, self.base_url, self.exp_url, self.code_col) for _, row in self.data_df.iterrows()]
        
        # results = pool.map(process_single_row, args)
        for arg in args: 
            self.pool.apply_async(process_single_row, args=arg, callback=gui_callback)
        


    def save_data(self, data_df, codes_df):
        """Save the updated data frames to the Excel file."""
        with pd.ExcelWriter(self.excel_file, engine="xlsxwriter") as writer:
            data_df.to_excel(writer, sheet_name=self.data_sheet_name, index=False)
            codes_df.to_excel(writer, sheet_name=self.codes_sheet_name, index=False)


def main():
    # Configurations
    workbook_file = "Ticketcodes.xlsx"
    workbook_file = os.path.join("./", workbook_file)
    print(f"Ticket Tracker - Medimeisterschaften {datetime.datetime.now().year}")
    tracker = TrackerProcessor(workbook_file, 'Code')

    try:
        data_df, codes_df = tracker.process_vouchers_parallel()
        tracker.save_data(data_df, codes_df)
    except Exception as e:
        print(f"An error occurred: {e}")
        sys.exit(1)


def run(excel_file, code_column_name, gui_callback):
    print(f"Ticket Tracker - Medimeisterschaften {datetime.datetime.now().year}")
    #data_df, codes_df = tracker.process_vouchers_parallel()
    #tracker.save_data(data_df, codes_df
    #return len(data_df)
    tracker = TrackerProcessor(excel_file, code_column_name)
    tracker.process_vouchers_parallel(gui_callback=gui_callback)


def process_single_row(row, base_url, exp_url, code_col):
    """Process a single row with a new browser instance."""
    browser = mechanicalsoup.StatefulBrowser()
    browser.open(base_url)
    browser.select_form('form[action="https://ticket.medimeisterschaften.com/redeem"]')

    if row.get("Status") == "eingelöst":
        return row, None  # Already redeemed, skip processing

    ticket_code = row.get(code_col)
    if pd.isna(ticket_code):
        return row, None  # No code, skip processing

    browser["voucher"] = ticket_code
    response = browser.submit_selected()

    if exp_url in response.url:
        row["Status"] = "eingelöst"
        print(f"{ticket_code} eingelöst")
        return row, None  # Successfully redeemed
    else:
        row["Status"] = "nicht eingelöst"
        print(f"{ticket_code} nicht eingelöst")
        return row, {
            "ID": row.get("ID"),
            "Vorname": row.get("Vorname"),
            "Nachname": row.get("Nachname"),
            "Email": row.get("Email"),
            "Code": ticket_code,
            "Status": "nicht eingelöst",
        }
    

if __name__ == "__main__":
    #freeze_support()
    main()