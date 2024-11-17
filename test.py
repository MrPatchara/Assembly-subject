import tkinter as tk
from tkinter import messagebox
import pandas as pd
from PIL import Image, ImageTk  # Import the necessary parts of Pillow

# Load the Excel data
file_path = 'Book1.xlsx'  # ใส่ path ของไฟล์ Excel
assembly_data = pd.read_excel(file_path)

# Function to search and display all matching entries
def search_mnemonic():
    mnemonic_input = entry_mnemonic.get().upper()
    if not mnemonic_input:
        messagebox.showwarning("Input Error", "Please enter a mnemonic!")
        return

    # Search for all rows matching the mnemonic
    matching_rows = assembly_data[assembly_data['Mnemonic'].str.upper() == mnemonic_input]
    if matching_rows.empty:
        messagebox.showerror("Search Error", f"No matching entries found for '{mnemonic_input}'")
        return

    # Clear previous results
    listbox_results.delete(0, tk.END)

    # Add header
    header = f"{'Mnemonic':<10} {'Operands':<20} {'Opcode':<10} {'Bytes':<5}"
    separator = "-" * len(header)
    listbox_results.insert(tk.END, header)
    listbox_results.insert(tk.END, separator)

    # Add all matching rows to the results listbox
    for _, row in matching_rows.iterrows():
        mnemonic = row['Mnemonic']
        operands = row['Operands'] if pd.notna(row['Operands']) else "None"
        opcode = row['Opcode']
        bytes_count = row['Bytes']
        result = f"{mnemonic:<10} {operands:<20} {opcode:<10} {bytes_count:<5}"
        listbox_results.insert(tk.END, result)

# Function to reset the input and results
def reset():
    entry_mnemonic.delete(0, tk.END)
    listbox_results.delete(0, tk.END)

# Create the GUI window
root = tk.Tk()
root.title("Assembly Mnemonic Finder")

# Vintage mode colors
bg_color = "#2b2b2b"  # Warm dark brown
fg_color = "#d4af37"  # Golden text
btn_color = "#3c3c3c"  # Dark gray buttons
font_style = ("Courier", 12, "bold")

# Apply vintage mode styles
root.configure(bg=bg_color)

# Add IC 8051 image as an icon in the top of the window using Pillow
try:
    ic_image = Image.open("ic_8051.png")  # ปรับเป็น path ของไฟล์รูปที่คุณต้องการ
    ic_image = ic_image.resize((150, 150))  # ปรับขนาดของรูปภาพตามต้องการ
    ic_image_tk = ImageTk.PhotoImage(ic_image)  # แปลงเป็นรูปแบบที่ Tkinter รองรับ
    label_icon = tk.Label(root, image=ic_image_tk, bg=bg_color)
    label_icon.image = ic_image_tk  # เก็บ reference ของรูปภาพเพื่อไม่ให้ถูกลบ
    label_icon.pack(pady=10)
except Exception as e:
    print(f"Error loading image: {e}")

# Mnemonic input label and entry
label_mnemonic = tk.Label(root, text="Enter Mnemonic:", bg=bg_color, fg=fg_color, font=font_style)
label_mnemonic.pack(pady=10)
entry_mnemonic = tk.Entry(root, width=40, bg="#1e1e1e", fg=fg_color, font=font_style, insertbackground=fg_color)
entry_mnemonic.pack(pady=5)

# Search button
btn_search = tk.Button(root, text="Search", command=search_mnemonic, bg=btn_color, fg=fg_color, font=font_style)
btn_search.pack(pady=10)

# Reset button
btn_reset = tk.Button(root, text="Reset", command=reset, bg=btn_color, fg=fg_color, font=font_style)
btn_reset.pack(pady=10)

# Results listbox
label_results = tk.Label(root, text="Search Results:", bg=bg_color, fg=fg_color, font=font_style)
label_results.pack(pady=10)
listbox_results = tk.Listbox(root, width=80, height=15, bg="#1e1e1e", fg=fg_color, font=font_style, selectbackground="#6a6a6a")
listbox_results.pack(pady=5)

# Run the application
root.mainloop()
