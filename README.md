# Report Generator

A Python desktop tool built with **Tkinter** and **Pandas** to
automatically generate filtered raw data Excel files from a large
multi‑sheet report.

The tool allows you to: - Select an Excel file (`.xlsx`) - Automatically
process multiple sheets - Filter rows based on defined business rules -
Save the generated raw datasets as new Excel files

------------------------------------------------------------------------

## \## Features

### **1. Raw Data 1--5**

Reads sheet: **3-Incident Resolved in Month**

Filters based on: - `Assigned_Group` in EUS group list\
- `Service_Type` contains `"User Service Restoration"` - `Status`
contains `"Closed"` or `"Resolved"`

Outputs a curated Excel file.

------------------------------------------------------------------------

### **2. Raw Data 6**

Reads sheet: **1-Open Incidents**

Filters: - EUS groups\
- User Service Restoration tickets

Outputs cleaned dataset.

------------------------------------------------------------------------

### **3. Raw Data 7 (User Provisioning)**

Reads sheet: **10-WorkOrder Completed in Month**

Filters: - EUS groups\
- Summary containing `"Computer or Accessories Request"` -
Categorization Tier 3 containing `"Desktop/Laptop"` or `"Loan"`

Generates a User Provisioning dataset.

------------------------------------------------------------------------

### **4. Raw Data 20**

Reads sheet: **10-WorkOrder Completed in Month**

Filters: - EUS groups\
- Status `"Closed"` or `"Completed"`

Adds fields: - `Type = "General"` - `Reopen = "No"`

Outputs a refined dataset.

------------------------------------------------------------------------

## \## Requirements

Install dependencies using:

``` bash
pip install -r requirements.txt
```

### requirements.txt

    openpyxl==3.1.5
    pandas==2.3.3

OpenPyXL is required for reading/writing Excel files.

------------------------------------------------------------------------

## \## Usage

1.  Run the script:

``` bash
python main.py
```

2.  Click **Open File**.
3.  Choose an Excel file (.xlsx).
4.  The button changes to **Generate Raw Data**.
5.  Click it to produce all four datasets.
6.  You will be prompted to save each generated Excel file.

------------------------------------------------------------------------

## \## File Structure

    project/
    │
    ├── main.py
    ├── requirements.txt
    └── README.md

------------------------------------------------------------------------

## \## Notes

-   Tkinter is part of the Python standard library and does not require
    installation.
-   All `.xlsx` reading/writing is handled by **pandas + openpyxl**.
-   The app uses a simple GUI for selecting files and saving results.

------------------------------------------------------------------------

## \## License

This project is free to use and modify.
