import pandas as pd
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

# Function to parse date values
def parse_date(date_value):
    try:
        if isinstance(date_value, str):
            return datetime.strptime(date_value, '%d/%m/%Y').date()
        elif isinstance(date_value, pd.Timestamp):
            return date_value.date()
        else:
            return pd.to_datetime(date_value).date()
    except Exception as e:
        print(f"Error processing date: {e}")
        return None

# Function to process the Excel file and check dates
def process_file(filepath):
    try:
        df = pd.read_excel(filepath, sheet_name='Sheet1')
        date_column = 'deadline'
        type_column = 'type'
        current_date = datetime.now().date()
        next_30days = current_date + timedelta(days=30)

        results = []
        for index, row in df.iterrows():
            date_value = row[date_column]
            type_value = row[type_column]

            if pd.notna(date_value):
                date_obj = parse_date(date_value)

                if date_obj:
                    if current_date <= date_obj <= next_30days:
                        results.append(f"{type_value}: {date_obj.strftime('%d/%m/%Y')}")

        return results
    except Exception as e:
        messagebox.showerror("Error", f"Failed to process file: {e}")
        return []

# Function to upload and process the Excel file
def upload_file():
    filepath = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx;*.xls")]
    )
    if filepath:
        results = process_file(filepath)
        if results:
            result_text.delete(1.0, tk.END)
            result_text.insert(tk.END, "Dates within the next 30 days:\n\n")
            for result in results:
                result_text.insert(tk.END, f"{result}\n")
        else:
            messagebox.showinfo("No Dates", "No dates within the next 30 days.")

# Create the main application window
app = tk.Tk()
app.title("Date Checker")
app.geometry("500x500")

# Create a label and button for uploading the Excel file
upload_label = tk.Label(app, text="Upload your Excel file to check dates", font=("Arial", 12))
upload_label.pack(pady=10)

upload_button = tk.Button(app, text="Upload Excel File", command=upload_file, font=("Arial", 12), bg="#4CAF50", fg="white")
upload_button.pack(pady=10)

# Create a scrollable text area for displaying results
result_text = scrolledtext.ScrolledText(app, width=60, height=20, font=("Arial", 12))
result_text.pack(pady=20)

# Run the application
app.mainloop()
