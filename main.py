import pandas as pd
from tkinter import *
from tkinter import filedialog, messagebox

excel_sheet = None

eus_groups = ["EUS Nitra", "EUS Hardware Nitra", "EUS Mobile Device Nitra",
            "EUS Office Printers Nitra", "EUS SCCM Nitra"]

def browse_file():
    try:
        global excel_sheet
        excel_sheet = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        file_button.config(text="Generate Raw Data", command=generate_data)
    except FileNotFoundError:
        messagebox.showerror("File Not Found!", "The selected file was not found.")

def generate_raw1_5():
    global excel_sheet
    df = pd.read_excel(excel_sheet, sheet_name="3-Incident Resolved in Month",
                        usecols=["Incident_Number", "Submit_Date",
                                 "Last_Resolved_Date", "Service_Type",
                                 "Priority", "SLM_Status",
                                 "Assigned_Group", "Assignee",
                                 "Status"])

    df["Stream"] = "EUS"

    data = df[df["Assigned_Group"].isin(eus_groups) &
                df["Status"].str.contains("Closed|Resolved", na=False)]

    data = data.loc[:, ["Incident_Number", "Submit_Date", "Last_Resolved_Date", "Service_Type", "Priority", "SLM_Status",
                          "Assigned_Group", "Assignee", "Status", "Stream"]]

    save_path = filedialog.asksaveasfilename(confirmoverwrite=True,
                                             defaultextension=".xlsx",
                                             initialfile="RawDataEus1-5",
                                             filetypes=[("Excel Files", "*.xlsx")])

    if save_path:
        data.to_excel(save_path, index=False)

def generate_raw6():
    global excel_sheet
    df = pd.read_excel(excel_sheet, sheet_name="1-Open Incidents",
                       usecols=["Incident_Number", "Submit_Date",
                                "Last_Resolved_Date", "Service_Type",
                                "Priority", "SLM_Status", "Assigned_Group",
                                "Assignee", "Status"])

    df["Stream"] = "EUS"

    data = df[df["Assigned_Group"].isin(eus_groups) &
              df["Service_Type"].str.contains("User Service Restoration")]


    data = data.loc[:, ["Incident_Number", "Submit_Date", "Last_Resolved_Date", "Service_Type", "Priority", "SLM_Status",
                        "Assigned_Group", "Assignee", "Status", "Stream"]]

    save_path = filedialog.asksaveasfilename(confirmoverwrite=True,
                                             defaultextension=".xlsx",
                                             initialfile="RawDataEus6",
                                             filetypes=[("Excel Files", "*.xlsx")])

    if save_path:
        data.to_excel(save_path, index=False)

def generate_raw7():
    global excel_sheet
    df = pd.read_excel(excel_sheet, sheet_name="10-WorkOrder Completed in Month",
                       usecols=["Work_Order_ID", "Categorization_Tier_1",
                                "Categorization_Tier_2", "Categorization_Tier_3",
                                "Summary", "SLM_Status", "First_Name",
                                "Last_Name", "AssignedGroup", "Status",
                                "Submitter", "Detailed_Description",
                                "CompletedDate", "Submit_Date"])

    data = df[df["AssignedGroup"].isin(eus_groups) &
              df["Summary"].str.contains("Computer or Accessories Request") &
              df["Categorization_Tier_3"].str.contains("Desktop/Laptop|Loan", na=False)]

    data = data.loc[:, ["Work_Order_ID", "Categorization_Tier_1", "Categorization_Tier_2", "Categorization_Tier_3", "Summary",
                        "SLM_Status", "First_Name", "Last_Name", "AssignedGroup", "Status", "Submitter", "Detailed_Description",
                        "CompletedDate", "Submit_Date"]]

    save_path = filedialog.asksaveasfilename(confirmoverwrite=True,
                                             defaultextension=".xlsx",
                                             initialfile="User Provisioning Data",
                                             filetypes=[("Excel Files", "*.xlsx")])

    if save_path:
        data.to_excel(save_path, index=False)

def generate_raw20():
    global excel_sheet
    df = pd.read_excel(excel_sheet, sheet_name="10-WorkOrder Completed in Month",
                       usecols=["Work_Order_ID", "Categorization_Tier_1", "Categorization_Tier_2",
                       "Categorization_Tier_3", "CompletedDate", "Submit_Date", "Status", "AssignedGroup"])

    df["Type"] = "General"
    df["Reopen"] = "No"

    data = df[df["AssignedGroup"].isin(eus_groups) &
              df["Status"].str.contains("Closed|Completed", na=False)]

    data.drop(columns="AssignedGroup")

    data = data.loc[:, ["Work_Order_ID", "Type", "Categorization_Tier_1", "Categorization_Tier_2", "Categorization_Tier_3",
                        "CompletedDate", "Submit_Date", "Status", "Reopen"]]

    save_path = filedialog.asksaveasfilename(confirmoverwrite=True,
                                             defaultextension=".xlsx",
                                             initialfile="RawDataEus20",
                                             filetypes=[("Excel Files", "*.xlsx")])

    if save_path:
        data.to_excel(save_path, index=False)

def generate_data():
    generate_raw1_5()
    generate_raw6()
    generate_raw7()
    generate_raw20()
    file_button.config(text="Open File", command=browse_file)

window = Tk()
window.title("Report Generator")
window.config(padx=100, pady=300, bg="grey")


file_button = Button(text="Open File", bg="white", highlightbackground="white", padx=20, pady=10, command=browse_file)
file_button.grid(column=1, row=1)

window.mainloop()