#!/usr/bin/env python
# coding: utf-8

# In[1]:


import tkinter
from tkinter import ttk, messagebox
from docxtpl import DocxTemplate
import datetime
import pandas as pd 
import os
import win32print

# Global variables
last_invoice_number = 0
INVOICE_FILE = "last_invoice_number.txt"
INVOICE_FOLDER = "D:\\Invoices"
invoice_list = []  # Define the invoice_list variable

# Create the invoice folder if it doesn't exist
if not os.path.exists(INVOICE_FOLDER):
    os.makedirs(INVOICE_FOLDER)

def save_last_invoice_number():
    with open(INVOICE_FILE, "w") as file:
        file.write(str(last_invoice_number))

def load_last_invoice_number():
    global last_invoice_number
    try:
        with open(INVOICE_FILE, "r") as file:
            last_invoice_number = int(file.read())
    except FileNotFoundError:
        pass

def clear_item():
    qty_spinbox.delete(0, tkinter.END)
    qty_spinbox.insert(0, "1")
    desc_entry.delete(0, tkinter.END)
    price_spinbox.delete(0, tkinter.END)
    price_spinbox.insert(0, "0.0")

def add_item():
    qty = float(qty_spinbox.get())
    desc = desc_entry.get()
    price = float(price_spinbox.get())
    line_total = qty * price
    invoice_item = [desc, qty, price, line_total]
    tree.insert('', 0, values=invoice_item)
    clear_item()
    invoice_list.append(invoice_item)

def new_invoice():
    global last_invoice_number
    last_invoice_number += 1
    invoice_number_entry.delete(0, tkinter.END)
    invoice_number_entry.insert(0, f"Dir{str(last_invoice_number).zfill(5)}")
    first_name_entry.delete(0, tkinter.END)
    last_name_entry.delete(0, tkinter.END)
    phone_entry.delete(0, tkinter.END)
    clear_item()
    tree.delete(*tree.get_children())
    invoice_list.clear()

def generate_invoice():
    global last_invoice_number
    invoice_number = invoice_number_entry.get()
    name = first_name_entry.get() + ' ' + last_name_entry.get()
    phone = phone_entry.get()
    date = date_entry.get()
    subtotal = sum(item[3] for item in invoice_list)
    salestax = 5
    total = subtotal + ((subtotal * salestax) / 100)

    # Generate invoice document
    doc = DocxTemplate("D:\\Invoice folder\\diraaz_invoice.docx")
    doc.render({
        "invoice_number": invoice_number,
        "name": name,
        "phone": phone,
        "date": date,
        "invoice_list": invoice_list,
        "subtotal": subtotal,
        "salestax": str(salestax) + "%",
        "total": total
    })

    doc_name = os.path.join(INVOICE_FOLDER, f"new_invoice_{name.replace(' ', '_')}_{datetime.datetime.now().strftime('%Y-%m-%d-%H%M%S')}.docx")
    doc.save(doc_name)

    # Create DataFrame for invoice details
    invoice_df = pd.DataFrame(columns=["Description", "Qty", "Unit Price", "Total", "Invoice Number", "Name", "Phone", "Date", "Subtotal"])
    invoice_df.loc[0] = [" ".join([item[0] for item in invoice_list]), sum(item[1] for item in invoice_list),
                         sum(item[2] for item in invoice_list), subtotal, invoice_number, name, phone, date, subtotal]

    excel_file = os.path.join(INVOICE_FOLDER, "invoice_details.xlsx")
    if os.path.exists(excel_file):
        # If the file already exists, append data to it
        existing_df = pd.read_excel(excel_file)
        new_df = pd.concat([existing_df, invoice_df], ignore_index=True)
        new_df.to_excel(excel_file, index=False)
    else:
        # If the file does not exist, create a new Excel file
        invoice_df.to_excel(excel_file, index=False)

    messagebox.showinfo("Invoice Generated", "Invoice has been generated successfully!")

    new_invoice()
    save_last_invoice_number()

def clear_invoice():
    tree.delete(*tree.get_children())
    invoice_list.clear()

def edit_invoice():
    selected_item = tree.selection()
    if selected_item:
        item_values = tree.item(selected_item)['values']
        desc_entry.delete(0, tkinter.END)
        desc_entry.insert(0, item_values[0])
        qty_spinbox.delete(0, tkinter.END)
        qty_spinbox.insert(0, str(item_values[1]))
        price_spinbox.delete(0, tkinter.END)
        price_spinbox.insert(0, str(item_values[2]))
        tree.delete(selected_item)
        invoice_list.remove(item_values)
    else:
        messagebox.showwarning("No Item Selected", "Please select an item to edit.")

def print_invoice():
    # Create a printer handle
    printer_name = win32print.GetDefaultPrinter()
    printer_handle = win32print.OpenPrinter(printer_name)

    # Set the printer properties
    properties = win32print.GetPrinter(printer_handle, 2)
    properties['pDevMode'].PaperSize = 9  # A4 paper size
    win32print.SetPrinter(printer_handle, 2, properties, 0)

    # Create a device context (DC)
    dc = win32print.GetDC(printer_name)

    # Start the printing job
    job = win32print.StartDocPrinter(printer_handle, 1, ("Invoice", None, "RAW"))

    try:
        win32print.StartPagePrinter(printer_handle)
        # Write your invoice content here, for simplicity, let's just print a test message
        message = "This is a test invoice content."
        win32print.TextOut(dc, 100, 100, message.encode('utf-8'))  # You can adjust the position

    finally:
        # End the printing job
        win32print.EndPagePrinter(printer_handle)
        win32print.EndDocPrinter(printer_handle)
        win32print.ClosePrinter(printer_handle)

window = tkinter.Tk()
window.title("Diraaz Invoice Generator")
window.geometry("800x600")

style = ttk.Style()
style.theme_use("clam")
style.configure("Treeview", background="white", foreground="black", fieldbackground="white")

frame = tkinter.Frame(window, bg="light blue")
frame.pack(padx=20, pady=10, fill=tkinter.BOTH, expand=True)

title_label = tkinter.Label(frame, text="Diraaz Invoice Generator", font=("Arial", 20), bg="light blue")
title_label.pack(pady=10)

invoice_frame = tkinter.Frame(frame, bg="light blue")
invoice_frame.pack(pady=10)

# Invoice Details Section
invoice_details_frame = tkinter.LabelFrame(invoice_frame, text="Invoice Details", bg="light blue")
invoice_details_frame.grid(row=0, column=0, padx=10, pady=10, sticky="w")

invoice_number_label = tkinter.Label(invoice_details_frame, text="Invoice Number:", bg="light blue")
invoice_number_label.grid(row=0, column=0, padx=5, pady=5)
invoice_number_entry = tkinter.Entry(invoice_details_frame)
invoice_number_entry.grid(row=0, column=1, padx=5, pady=5)

first_name_label = tkinter.Label(invoice_details_frame, text="First Name:", bg="light blue")
first_name_label.grid(row=1, column=0, padx=5, pady=5)
first_name_entry = tkinter.Entry(invoice_details_frame)
first_name_entry.grid(row=1, column=1, padx=5, pady=5)

last_name_label = tkinter.Label(invoice_details_frame, text="Last Name:", bg="light blue")
last_name_label.grid(row=1, column=2, padx=5, pady=5)
last_name_entry = tkinter.Entry(invoice_details_frame)
last_name_entry.grid(row=1, column=3, padx=5, pady=5)

phone_label = tkinter.Label(invoice_details_frame, text="Phone:", bg="light blue")
phone_label.grid(row=2, column=0, padx=5, pady=5)
phone_entry = tkinter.Entry(invoice_details_frame)
phone_entry.grid(row=2, column=1, padx=5, pady=5)

date_label = tkinter.Label(invoice_details_frame, text="Date:", bg="light blue")
date_label.grid(row=2, column=2, padx=5, pady=5)
date_entry = tkinter.Entry(invoice_details_frame)
date_entry.grid(row=2, column=3, padx=5, pady=5)

# Items Section
items_frame = tkinter.LabelFrame(invoice_frame, text="Items", bg="light blue")
items_frame.grid(row=0, column=1, padx=10, pady=10)

qty_label = tkinter.Label(items_frame, text="Qty:", bg="light blue")
qty_label.grid(row=0, column=0, padx=5, pady=5)
qty_spinbox = tkinter.Spinbox(items_frame, from_=1, to=100)
qty_spinbox.grid(row=0, column=1, padx=5, pady=5)
qty_spinbox.insert(0, "1")

desc_label = tkinter.Label(items_frame, text="Description:", bg="light blue")
desc_label.grid(row=0, column=2, padx=5, pady=5)
desc_entry = tkinter.Entry(items_frame)
desc_entry.grid(row=0, column=3, padx=5, pady=5)

price_label = tkinter.Label(items_frame, text="Unit Price:", bg="light blue")
price_label.grid(row=0, column=4, padx=5, pady=5)
price_spinbox = tkinter.Spinbox(items_frame, from_=0.0, to=500, increment=0.5)
price_spinbox.grid(row=0, column=5, padx=5, pady=5)
price_spinbox.insert(0, "0.0")

add_item_button = tkinter.Button(items_frame, text="Add Item", command=add_item)
add_item_button.grid(row=0, column=6, padx=5, pady=5)

# Invoice List Section
columns = ("Item", "Qty", "Unit Price", "Total")
tree = ttk.Treeview(frame, columns=columns, show="headings", height=10)
tree.heading('Item', text='Item')
tree.heading('Qty', text='Qty')
tree.heading('Unit Price', text='Unit Price')
tree.heading('Total', text='Total')
tree.pack(padx=20, pady=10)

# Buttons Section
buttons_frame = tkinter.Frame(frame, bg="light blue")
buttons_frame.pack(pady=10)

save_invoice_button = tkinter.Button(buttons_frame, text="Generate Invoice", command=generate_invoice)
save_invoice_button.grid(row=0, column=0, padx=5, pady=5)

new_invoice_button = tkinter.Button(buttons_frame, text="New Invoice", command=new_invoice)
new_invoice_button.grid(row=0, column=1, padx=5, pady=5)

clear_button = tkinter.Button(buttons_frame, text="Clear", command=clear_invoice)
clear_button.grid(row=0, column=2, padx=5, pady=5)

edit_button = tkinter.Button(buttons_frame, text="Edit", command=edit_invoice)
edit_button.grid(row=0, column=3, padx=5, pady=5)

print_button = tkinter.Button(buttons_frame, text="Print", command=print_invoice)
print_button.grid(row=0, column=4, padx=5, pady=5)

# Load the last used invoice number
load_last_invoice_number()
# Automatically set the initial invoice number
new_invoice()

window.mainloop()


# In[ ]:




