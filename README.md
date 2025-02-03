# Medimeisterschaften Ticket Tracker

## Overview
The **Medimeisterschaften Ticket Tracker** is a Python-based desktop application that helps manage and track ticket codes for Medimeisterschaften. Users can select an Excel file containing ticket codes, specify the column holding the codes, and process the data. The application updates the status of each code within the Excel file.

## Features
- User-friendly **Tkinter** interface
- Supports **Excel (.xlsx, .xls)** files
- Processes ticket codes asynchronously using **multiprocessing**

### Running the Application


1. Clone this repository:
   ```sh
   git clone https://github.com/your-repo/medimeisterschaften-tracker.git
   cd medimeisterschaften-tracker
   ```
2. Run the application: Click on the "tracker.exe" file to start the application

## Usage
1. Open the application.
2. Select the Excel file containing ticket codes.
3. Enter the column name where ticket codes are stored (default: `Code`).
4. Click **"Process File"** to check ticket statuses.
5. Wait for completion – results will be written to the previously defined excel file

## Technologies Used
- **Python** (Tkinter, Pandas, Multiprocessing)
- **Tkinter** (GUI Framework)
- **Pillow** (Image Handling)


## Contact
For any questions, please create an issue. 

