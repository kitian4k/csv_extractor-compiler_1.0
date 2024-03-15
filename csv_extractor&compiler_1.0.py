import tkinter as tk
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText
import pandas as pd
from datetime import datetime


# App
def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    x = (screen_width - width) // 2
    y = (screen_height - height) // 2

    window.geometry(f'{width}x{height}+{x}+{y}')

# python back-end
def upload_and_process():
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv")])
    num_files_selected = len(file_paths)  # Number of files selected
    master_df = pd.DataFrame(columns=['Security Symbol', 'Security Name'])  # Initialize a master DataFrame
    date_cols = set()  # Set to store unique date columns

    if num_files_selected > 0:
        try:
            for file_path in file_paths:
                if file_path.endswith('.csv'):
                    df = pd.read_csv(file_path)
                else:
                    df = pd.read_excel(file_path)

                # Extract date from the file name
                file_name = file_path.split("/")[-1].split(".")[0]
                report_date = datetime.strptime(file_name.replace("QUOTE", ""), "%Y%m%d").date()
                date_cols.add(report_date)

                # Extract unique combinations of 'Security Symbol' and 'Security Name' from the file
                unique_symbols_names = df[['Security Symbol', 'Security Name']].drop_duplicates()

                # Merge unique combinations into the master DataFrame
                master_df = pd.merge(master_df, unique_symbols_names, on=['Security Symbol', 'Security Name'], how='outer')

                # Merge total values from the file into the master DataFrame
                master_df = pd.merge(master_df, df[['Security Symbol', 'Security Name', 'Total Value']], on=['Security Symbol', 'Security Name'], how='left')
                master_df.rename(columns={'Total Value': report_date}, inplace=True)

            # Replace NaN values with 0
            master_df.fillna(0.0, inplace=True)

            # Sort the master DataFrame by 'Security Symbol' and 'Security Name'
            master_df.sort_values(by=['Security Symbol', 'Security Name'], inplace=True)

            # Reorder columns to have 'Security Symbol' and 'Security Name' first, followed by date columns
            cols = ['Security Symbol', 'Security Name'] + sorted(date_cols)
            master_df = master_df[cols]

            # Update the output text widget with the final output
            output_text.delete(1.0, tk.END)
            output_text.insert(tk.END, "Final Output:\n")
            output_text.insert(tk.END, master_df.to_string(index=False))

            # Save final DataFrame to CSV file
            save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
            master_df.to_csv(save_path, index=False)
            
            output_text.insert(tk.END, "\n\nFinal Output saved to: " + save_path)
            
            # Update label with saved file path
            saved_file_label.config(text="Saved File Path: " + save_path)

        except Exception as e:
            output_text.insert(tk.END, "\n\nError: " + str(e))
    else:
        output_text.insert(tk.END, "\n\nNo files selected.")

    # Update label with number of files selected
    num_files_label.config(text="Number of Files Uploaded: " + str(num_files_selected))

# front-end
# tkinter area
root = tk.Tk()
root.title("Luna Securities CSV Extractor")

# Size of App
window_width = 1200  # Adjusted window width
window_height = 600
center_window(root, window_width, window_height)

# Button
upload_button = tk.Button(root, text="Upload Excel/CSV", command=upload_and_process)
upload_button.pack(pady=20)

# Display Numbers of Files Uploaded
num_files_label = tk.Label(root, text="Number of Files Uploaded: ")
num_files_label.pack()

# Display 
output_text = ScrolledText(root, width=200, height=30, wrap=tk.NONE)  # Set wrap to NONE for horizontal scrolling
output_text.pack()

# File Path of Saved CSV File
saved_file_label = tk.Label(root, text="Saved File Path: ")
saved_file_label.pack()

root.mainloop()
