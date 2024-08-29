import os
import sys
import zipfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import shutil
from lxml import etree
from PIL import Image, ImageTk

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def select_files(file_ext):
    file_paths = filedialog.askopenfilenames(title=f"Select {file_ext} Files", filetypes=[(f"{file_ext.upper()} files", f"*.{file_ext}")])
    if not file_paths:
        messagebox.showerror("Error", f"No {file_ext} files selected.")
        return None
    return file_paths

def change_extension(file_path, new_extension):
    base = os.path.splitext(file_path)[0]
    new_file_path = f"{base}{new_extension}"
    os.rename(file_path, new_file_path)
    return new_file_path

def extract_workbook_xml(zip_file_path, temp_dir):
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        zip_ref.extract('xl/workbook.xml', temp_dir)

def modify_workbook_xml(temp_dir):
    workbook_xml_path = os.path.join(temp_dir, 'xl', 'workbook.xml')
    
    # Parse the XML with lxml and preserve namespaces
    parser = etree.XMLParser(remove_blank_text=True)
    tree = etree.parse(workbook_xml_path, parser)
    root = tree.getroot()

    # Define the namespaces used in the XML
    namespaces = {
        'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    }

    # Find and remove the <workbookProtection> element
    workbook_protection = root.find('main:workbookProtection', namespaces)
    if workbook_protection is not None:
        root.remove(workbook_protection)

    # Write the modified XML back to the file
    tree.write(workbook_xml_path, pretty_print=True, xml_declaration=True, encoding='UTF-8')

def replace_workbook_xml_in_zip(zip_file_path, temp_dir):
    workbook_xml_path = os.path.join(temp_dir, 'xl', 'workbook.xml')

    # Create a new zipfile and write everything from the original zipfile except the workbook.xml,
    # then add the modified workbook.xml back into the zip.
    temp_zip_path = os.path.join(temp_dir, 'temp.zip')
    with zipfile.ZipFile(zip_file_path, 'r') as original_zip:
        with zipfile.ZipFile(temp_zip_path, 'w') as temp_zip:
            for item in original_zip.infolist():
                if item.filename != 'xl/workbook.xml':
                    temp_zip.writestr(item, original_zip.read(item.filename))
            # Add the modified workbook.xml
            temp_zip.write(workbook_xml_path, 'xl/workbook.xml')

    # Replace the original zip with the modified one
    shutil.move(temp_zip_path, zip_file_path)

def process_files():
    file_ext = file_extension.get()
    file_paths = select_files(file_ext)
    if not file_paths:
        return

    for file_path in file_paths:
        # Step 1: Change the extension to .zip
        zip_file_path = change_extension(file_path, ".zip")

        # Step 2: Create a temporary directory to work in
        temp_dir = os.path.join(os.path.dirname(zip_file_path), 'temp_extract')
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        os.makedirs(temp_dir)

        try:
            # Step 3: Extract, modify, and replace the workbook.xml
            extract_workbook_xml(zip_file_path, temp_dir)
            modify_workbook_xml(temp_dir)
            replace_workbook_xml_in_zip(zip_file_path, temp_dir)

            # Step 4: Change the extension back
            final_file_path = change_extension(zip_file_path, f".{file_ext}")
            
            messagebox.showinfo("Success", f"WorkbookProtection removed and saved successfully to {os.path.basename(final_file_path)}.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while processing {os.path.basename(file_path)}: {str(e)}")
        finally:
            # Clean up the temporary directory
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)

def create_gui():
    root = tk.Tk()
    root.title("Workbook Unprotector")
    root.configure(bg='#FBE3D6')

    # Set the window and taskbar icon
    icon_path = resource_path("app.ico")
    if os.path.exists(icon_path):
        root.iconbitmap(icon_path)

    # Logo
    logo_path = resource_path("logo.png")
    if os.path.exists(logo_path):
        try:
            img = Image.open(logo_path)
            logo_img = ImageTk.PhotoImage(img)
            logo_label = tk.Label(root, image=logo_img, bg='#FBE3D6')
            logo_label.image = logo_img  # Keep a reference to avoid garbage collection
            logo_label.pack(pady=10)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load logo image: {str(e)}")

    # File Extension Dropdown
    tk.Label(root, text="Select file extension:", font=("Calibri", 14), bg='#FBE3D6').pack(pady=5)
    global file_extension
    file_extension = tk.StringVar(value="xlsx")
    extension_options = ["xlsx", "xlsm", "xlsb", "xltx", "xltm", "xls", "xlt"]
    extension_dropdown = ttk.Combobox(root, textvariable=file_extension, values=extension_options, state="readonly", font=("Calibri", 12))
    extension_dropdown.pack(pady=10)

    # Unprotect Button
    unprotect_button = tk.Button(root, text="UNPROTECT", command=process_files, bg="#104862", fg="white", font=("Calibri", 16), bd=0, padx=20, pady=10)
    unprotect_button.pack(pady=20)
    unprotect_button.config(relief="flat", overrelief="solid", cursor="hand2")

    # Run the application
    root.mainloop()

if __name__ == "__main__":
    create_gui()
