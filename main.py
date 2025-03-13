import tkinter as tk
import subprocess
import sys
from tkinter import messagebox

sys.setrecursionlimit(2000)

# Functions to run corresponding files
def open_data_entry():
    try:
        subprocess.Popen([sys.executable, "dataEntry.py"])
    except FileNotFoundError:
        messagebox.showerror("Error", "dataEntry.py not found!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def open_customer_report():
    try:
        subprocess.Popen([sys.executable, "reportGenerateCustomer.py"])
    except FileNotFoundError:
        messagebox.showerror("Error", "reportGenerateCustomer.py not found!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def open_customer_payments():
    try:
        subprocess.Popen([sys.executable, "managePaymentsCustomer.py"])
    except FileNotFoundError:
        messagebox.showerror("Error", "managePaymentsCustomer.py not found!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def open_lacing_report():
    try:
        subprocess.Popen([sys.executable, "reportGenerateLacing.py"])
    except FileNotFoundError:
        messagebox.showerror("Error", "reportGenerateLacing.py not found!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def open_lacing_payments():
    try:
        subprocess.Popen([sys.executable, "managePaymentsLacing.py"])
    except FileNotFoundError:
        messagebox.showerror("Error", "managePaymentsLacing.py not found!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    # Create the main window
    root = tk.Tk()
    root.title("Shanmugaraj")
    root.geometry("500x400")
    root.configure(bg="#f4f4f9")  # Light background color

    # Set window icon
    # root.iconbitmap("app_icon.ico")

    # Title Label
    title_label = tk.Label(
        root,
        text="Welcome to the Main Menu",
        font=("Helvetica", 18, "bold"),
        bg="#f4f4f9",
        fg="#333"
    )
    title_label.pack(pady=20)

    # Button styles
    button_style = {
        "font": ("Helvetica", 12),
        "bg": "#007BFF",  # Blue background
        "fg": "#ffffff",  # White text
        "activebackground": "#0056b3",  # Darker blue when pressed
        "activeforeground": "#ffffff",  # White text when pressed
        "width": 20,
        "height": 2,
        "relief": "raised",
        "bd": 3
    }

    # Buttons for each functionality
    btn_data_entry = tk.Button(root, text="Data Entry", command=open_data_entry, **button_style)
    btn_data_entry.pack(pady=10)

    btn_customer_report = tk.Button(root, text="Customer Report", command=open_customer_report, **button_style)
    btn_customer_report.pack(pady=10)

    btn_customer_payments = tk.Button(root, text="Customer Payments", command=open_customer_payments, **button_style)
    btn_customer_payments.pack(pady=10)

    btn_lacing_report = tk.Button(root, text="Lacing Report", command=open_lacing_report, **button_style)
    btn_lacing_report.pack(pady=10)

    btn_lacing_payments = tk.Button(root, text="Lacing Payments", command=open_lacing_payments, **button_style)
    btn_lacing_payments.pack(pady=10)

    # Exit Button
    btn_exit = tk.Button(root, text="Exit", command=root.quit, **button_style)
    btn_exit.pack(pady=10)

    # Run the GUI event loop
    root.mainloop()
