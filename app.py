import customtkinter as ctk
import threading
import queue
from tkinter import filedialog
from cleanup_logic import run_cleanup
from datetime import datetime
from pathlib import Path

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("GSPN Part Collection Report Cleaner")
app.geometry("410x470")

selected_file = ctk.StringVar(value="No file selected")
status_queue = queue.Queue()

title = ctk.CTkLabel(app, text="GSPN Part Report Cleaner", font=("Arial", 24, "bold"))
title.pack(pady=10)

def choose_file():
    file_path = filedialog.askopenfilename(
        title="Select GSPN Excel Report",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if file_path:
        selected_file.set(file_path)


def update_status(message):
    status_queue.put(message)


def process_status_queue():
    while not status_queue.empty():
        message = status_queue.get()

        status_box.configure(state="normal")
        status_box.insert("end", f"{message}\n")
        status_box.see("end")
        status_box.configure(state="disabled")

    app.after(100, process_status_queue)


def start_cleanup():
    input_file = selected_file.get()

    if input_file == "No file selected":
        update_status("Status: Please select a file first")
        return

    input_path = Path(input_file)
    date_stamp = datetime.now().strftime("%m_%d_%y")

    output_file = input_path.with_name(
        f"{input_path.stem}_Cleaned_{date_stamp}.xlsx"
    )

    start_button.configure(state="disabled")
    select_button.configure(state="disabled")

    update_status("Status: Running...")

    def worker():
        try:
            run_cleanup(input_file, output_file, update_status)
            update_status("Cleanup finished successfully.")
        except Exception as e:
            update_status(f"Error: {e}")
        finally:
            app.after(0, lambda: start_button.configure(state="normal"))
            app.after(0, lambda: select_button.configure(state="normal"))

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()

file_frame = ctk.CTkFrame(app)
file_frame.pack(pady=(10, 8), padx=20, fill="x")

file_title = ctk.CTkLabel(file_frame, text="Selected File", font=("Arial", 13, "bold"))
file_title.pack(anchor="w", padx=12, pady=(8, 0))

file_label = ctk.CTkLabel(file_frame, textvariable=selected_file, wraplength=360, justify="left")
file_label.pack(anchor="w", padx=12, pady=(2, 10))

select_button = ctk.CTkButton(app, text="Select Excel File", command=choose_file)
select_button.pack(pady=10)

start_button = ctk.CTkButton(app, text="Start Cleanup", command=start_cleanup)
start_button.pack(pady=10)

# log_label = ctk.CTkLabel(app, text="Progress Log")
# log_label.pack(pady=(10, 10))

status_box = ctk.CTkTextbox(app, width=340, height=200)
status_box.pack(pady=10)
status_box.insert("end", "Status: Ready\n")
status_box.configure(state="disabled")

process_status_queue()
app.mainloop()