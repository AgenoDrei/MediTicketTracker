import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from PIL import Image, ImageTk
from tracker import run


class TicketTrackerApp:
    def __init__(self, root):
        """Initialize the UI components."""
        self.root = root
        self.root.title("Medimeisterschaften Ticket Tracker")
        #self.root.geometry("500x450")
        #self.root.columnconfigure(0, weight=1) 
        self.root.resizable(False, False)

        # Title Label
        self.title = tk.Label(root, text="Medimeisterschaften Ticket Tracker", font=("Arial", 18))
        self.title.grid(row=0, column=0, columnspan=2, pady=10)

        # Image
        self.image = Image.open("icon.jpg").resize((150, 150))
        self.image = ImageTk.PhotoImage(self.image)
        self.icon = tk.Label(root, image=self.image, width=150, height=150)
        self.icon.grid(row=1, column=0, columnspan=2, pady=5)

        # Description Label
        self.description = tk.Label(root, text="Diese Anwendung hilft beim Tracken von Ticketcodes für die Medimeisterschaften. "
                                               "Wählen Sie eine Excel-Datei aus, die alle Ticketcodes in der angegebenen "
                                               "Spalte (Standard = 'Code') enthält. Der Status der Codes wird in die Spalte 'Status' geschrieben.",
                                    wraplength=450, font=("Arial", 11), justify="center")
        self.description.grid(row=2, column=0, columnspan=2, pady=5, padx=10, sticky="w")

        # Separator Line
        self.line = ttk.Separator(root, orient="horizontal")
        self.line.grid(row=3, column=0, columnspan=2, sticky="ew", pady=10)

        # Column Label & Entry
        self.column_label = tk.Label(root, text="Gebe den Spaltennamen für die MediCodes ein:", font=("Arial", 10))
        self.column_label.grid(row=4, column=0, sticky="w", padx=10, pady=(0, 5))

        self.column_entry = tk.Entry(root, width=20, justify="left")
        self.column_entry.insert(0, "Code")
        self.column_entry.grid(row=4, column=1, pady=(0, 5), padx=10, sticky="e")

        self.line2 = ttk.Separator(root, orient="horizontal")
        self.line2.grid(row=5, column=0, columnspan=2, sticky="ew", pady=10)

        # File Selection Button / File Label
        self.file_title = tk.Label(root, text="Wähle die Excel-Datei mit den Ticketcodes und Teilnehmerdaten:", font=("Arial", 10))
        self.file_title.grid(row=6, column=0, columnspan=2, sticky="ew", padx=10, pady=5)
        self.file_button = tk.Button(root, text="Pfad für die Excel-Datei", command=self.select_file)
        self.file_button.grid(row=7, column=0, pady=5)

        self.file_label = tk.Label(root, text="Keine Datei ausgewählt", wraplength=380, font=("Arial", 8))
        self.file_label.grid(row=7, column=1, pady=5)

        # Process Button
        self.process_button = tk.Button(root, text="Teste die MediCodes", bg="#265e5f", fg='white', command=self.process_file)
        self.process_button.grid(row=8, column=0, columnspan=2, pady=10, padx=50, sticky="ew")

        # Store file path
        self.file_path = None

    def select_file(self):
        """Open file dialog to select an Excel file."""
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if file_path:
            self.file_path = file_path
            self.file_label.config(text=f"File: {os.path.basename(file_path)}")

    def process_file(self):
        """Process the selected Excel file based on the column name input."""
        if not self.file_path:
            messagebox.showerror("Error", "Wähle zuerst eine Excel-Datei aus!")
            return

        column_name = self.column_entry.get()
        if not column_name:
            messagebox.showerror("Error", "Gib bitte einen Spaltennamen ein.")
            return

        try:
            # Read Excel file and check if colum exists
            df = pd.read_excel(self.file_path)
            if column_name not in df.columns:
                messagebox.showerror("Error", f"Spalte '{column_name}' existiert nicht in der Datei.")
                return

            # Process the column (example: print its values)
            num_results = run(self.file_path, column_name)
            messagebox.showinfo("Success", f"{num_results} Ticketcodes erfolgreich geprüft!")

        except Exception as e:
            messagebox.showerror("Error", f"Fehler beim Verarbeiten der Medi-Codes: {e}")
        finally:
            print("Processing completed.")
            sys.exit(0)

# Run the application
if __name__ == "__main__":
    root = tk.Tk()
    app = TicketTrackerApp(root)
    root.mainloop()
