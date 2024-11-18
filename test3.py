import tkinter as tk
from tkinter import messagebox
import pandas as pd
import os
from PIL import Image, ImageTk
import platform

# Load the Excel data
file_path = 'Book1.xlsx'  # ใส่ path ของไฟล์ Excel
assembly_data = pd.read_excel(file_path)

# Function to search and display all matching entries
def search_mnemonic():
    mnemonic_input = entry_mnemonic.get().strip().upper()
    operands_input = entry_operands.get().strip().upper()

    if not mnemonic_input and not operands_input:
        messagebox.showwarning("Input Error", "Please enter at least one search field!")
        return

    # Apply filters for Mnemonic and Operands
    matching_rows = assembly_data.copy()
    if mnemonic_input:  # Filter by Mnemonic if provided
        matching_rows = matching_rows[matching_rows['Mnemonic'].str.upper() == mnemonic_input]
    if operands_input:  # Filter by Operands if provided
        matching_rows = matching_rows[matching_rows['Operands'].str.upper().str.contains(operands_input, na=False)]

    if matching_rows.empty:
        messagebox.showerror("Search Error", "No matching entries found!")
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
    entry_operands.delete(0, tk.END)
    listbox_results.delete(0, tk.END)

# Function to open Excel file in Microsoft Excel
def open_excel():
    try:
        if platform.system() == 'Windows':  # On Windows
            os.startfile(file_path)
        elif platform.system() == 'Darwin':  # On macOS
            os.system(f"open {file_path}")
        else:  # On Linux
            os.system(f"xdg-open {file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to open file: {e}")

# Function to show developer information
def show_developer_info():
    # Create a new top-level window with the same theme as the main program
    developer_window = tk.Toplevel(root)
    developer_window.title("Contact Developer")
    developer_window.configure(bg=bg_color)
    # adjust size of the window
    developer_window.geometry("290x270")
    # add icon of the window
    developer_window.iconbitmap("dev.ico")

    # Label for developer information
    developer_info = """
   MR. Patchara Al-umaree
   Patcharaalumaree@gmail.com
   https://github.com/MrPatchara
    """
    
    # add picture of developer 
    try:
        ic_image = Image.open("pic.png")  # ปรับเป็น path ของไฟล์รูปที่คุณต้องการ
        ic_image = ic_image.resize((90, 90))  # ปรับขนาดของรูปภาพตามต้องการ
        ic_image_tk = ImageTk.PhotoImage(ic_image)  # แปลงเป็นรูปแบบที่ Tkinter รองรับ
        label_icon = tk.Label(developer_window, image=ic_image_tk, bg=bg_color)
        label_icon.image = ic_image_tk  # เก็บ reference ของรูปภาพเพื่อไม่ให้ถูกลบ
        label_icon.pack(pady=10) 
    except Exception as e:
        print(f"Error loading image: {e}")

    label_info = tk.Label(developer_window, text=developer_info, bg=bg_color, fg=fg_color, font=font_style)
    label_info.pack(padx=10, pady=10) 

    # adjust size text
    label_info.config(font=("Courier", 10, "bold"))

    # Close button for the developer info window
    btn_close = tk.Button(developer_window, text="Close", command=developer_window.destroy, bg=btn_color, fg=fg_color, font=font_style)
    btn_close.pack(pady=10)

# Create the GUI window
root = tk.Tk()
root.title("Assembly Opcode Finder (MCS-51)")
# adjust size of the window
root.geometry("490x710")
# add icon of the window
root.iconbitmap("cpu.ico")

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
    ic_image = ic_image.resize((160, 160))  # ปรับขนาดของรูปภาพตามต้องการ
    ic_image_tk = ImageTk.PhotoImage(ic_image)  # แปลงเป็นรูปแบบที่ Tkinter รองรับ
    label_icon = tk.Label(root, image=ic_image_tk, bg=bg_color)
    label_icon.image = ic_image_tk  # เก็บ reference ของรูปภาพเพื่อไม่ให้ถูกลบ
    label_icon.pack(pady=10)
except Exception as e:
    print(f"Error loading image: {e}")

# Create the menu bar
menu_bar = tk.Menu(root)

# File menu
file_menu = tk.Menu(menu_bar, tearoff=0)
file_menu.add_command(label="Open Excel File", command=open_excel)
menu_bar.add_cascade(label="Settings", menu=file_menu)

# Help menu
help_menu = tk.Menu(menu_bar, tearoff=0)
help_menu.add_command(label="Contact Developer", command=show_developer_info)
menu_bar.add_cascade(label="Help", menu=help_menu)

# Attach the menu to the root window
root.config(menu=menu_bar)

# Frame for Mnemonic and Operands
frame_inputs = tk.Frame(root, bg=bg_color)
frame_inputs.pack(pady=10)

# Mnemonic input label and entry
label_mnemonic = tk.Label(frame_inputs, text="Enter Mnemonic:", bg=bg_color, fg=fg_color, font=font_style)
label_mnemonic.grid(row=0, column=0, padx=10, pady=5)
entry_mnemonic = tk.Entry(frame_inputs, width=20, bg="#1e1e1e", fg=fg_color, font=font_style, insertbackground=fg_color)
entry_mnemonic.grid(row=1, column=0, padx=10, pady=5)

# Operands input label and entry
label_operands = tk.Label(frame_inputs, text="Enter Operands:", bg=bg_color, fg=fg_color, font=font_style)
label_operands.grid(row=0, column=1, padx=10, pady=5)
entry_operands = tk.Entry(frame_inputs, width=20, bg="#1e1e1e", fg=fg_color, font=font_style, insertbackground=fg_color)
entry_operands.grid(row=1, column=1, padx=10, pady=5)


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
