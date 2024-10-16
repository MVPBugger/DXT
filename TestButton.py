import os
import tkinter as tk
from tkinter import messagebox
from tkinter import font
import threading
import json
import datetime
from PIL import Image, ImageTk  # Import these from Pillow

def run_script():
    btn.config(state=tk.DISABLED)
    status_label.config(text="Running script...", fg="blue")
    root.update()
    
    def execute_script():
        try:
            os.system('python "C:\\Users\\Rolando Pancho\\Documents\\EXTRACTEXCELFILEFINAL.py"')
            # Save the current extraction date
            last_extraction_date = datetime.datetime.now().date()
            with open('last_extraction.json', 'w') as f:
                json.dump({'last_extraction': last_extraction_date.strftime('%Y-%m-%d')}, f)
            status_label.config(text="Data extraction completed, check SharePoint.", fg="green")
            update_last_extraction_info()
        except Exception as e:
            status_label.config(text=f"Error: {e}", fg="red")
        finally:
            btn.config(state=tk.NORMAL)

    threading.Thread(target=execute_script).start()

def update_last_extraction_info():
    try:
        with open('last_extraction.json', 'r') as f:
            data = json.load(f)
            last_extraction_date = datetime.datetime.strptime(data['last_extraction'], '%Y-%m-%d').date()
            days_since = (datetime.datetime.now().date() - last_extraction_date).days
            last_extraction_label.config(text=f"Last extraction date: {last_extraction_date} ({days_since} days ago)", fg="black")
    except (FileNotFoundError, json.JSONDecodeError):
        last_extraction_label.config(text="No previous extraction found.", fg="red")

def on_exit():
    if messagebox.askyesno("Exit", "Are you sure you want to exit?"):
        root.quit()

# Create a themed root window
root = tk.Tk()
root.title("Greenprofi Projekte")
root.geometry("500x500")
root.configure(bg="#f0f0f0")

# Define fonts
header_font = font.Font(family="Helvetica", size=16, weight="bold")
button_font = font.Font(family="Helvetica", size=12)

from PIL import Image, ImageTk

# Load and resize the JPEG logo
logo_path = "C:\\Users\\Rolando Pancho\\Downloads\\ib_schmidt_gmbh_logo.jpg" # Replace with your actual logo file path
logo_image = Image.open(logo_path)
logo_image = logo_image.resize((200, 200), Image.LANCZOS)  
logo_photo = ImageTk.PhotoImage(logo_image)

# Add the logo to the root window, and place it in the top right corner
logo_label = tk.Label(root, image=logo_photo, bg="#f0f0f0")
logo_label.place(relx=0.5, rely=1.0, anchor="s", y=-10)  # Place in the top right corner with 10px padding from the right


# Header label
header_label = tk.Label(root, text="Greenprofi Data Extraction Tool", font=header_font, bg="#f0f0f0")
header_label.pack(pady=10)

# Last extraction info label
last_extraction_label = tk.Label(root, text="", bg="#f0f0f0", font=button_font)
last_extraction_label.pack(pady=5)
update_last_extraction_info()

# Extract Data button
btn = tk.Button(root, text="Extract Data", command=run_script, font=button_font, bg="#4CAF50", fg="white", padx=10, pady=5)
btn.pack(pady=20)

# Status label
status_label = tk.Label(root, text="", bg="#f0f0f0", font=button_font)
status_label.pack()

# Exit button
exit_btn = tk.Button(root, text="Exit", command=on_exit, font=button_font, bg="#f44336", fg="white", padx=10, pady=5)
exit_btn.pack(pady=10)

# Update extraction info on startup
update_last_extraction_info()

root.mainloop()
