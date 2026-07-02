import tkinter as tk 
import sqlite3
from datetime import datetime
from tkinter import messagebox
import tkinter.ttk as ttk
import csv
from dateutil.relativedelta import relativedelta
from matplotlib import pyplot as plt
from openpyxl import Workbook

#Database Connection

conn = sqlite3.connect("expense.db")
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS expenses(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    title TEXT,
    amount INTEGER,
    category TEXT,
    date TEXT
)
""")
conn.commit()

print("Database Created Successfully")

#cursor.execute(
    #"ALTER TABLE expenses ADD COLUMN date TEXT"
#)

conn.commit()

def add_expense(): 
    
    title = entry_title.get()
    amount = entry_amount.get()
    category = entry_category.get()
    
    if title == "" or amount == "" or category == "":
        result_label.config(
            text="Please Fill All Fields"
        )
        return
    if not amount.isdigit():
        result_label.config(text="Amount must be a number!"
        )
        return
    
    try:
        amount = int(amount)
    except:
        result_label.config(
            text="Amount must be number"
        )
        return

    date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    cursor.execute(
        "INSERT INTO expenses(title,amount,category,date) VALUES(?,?,?,?)",
        (title,amount,category,date)
    )
    
    conn.commit()
    
    result_label.config(
        text="Expense Added Successfully"
    )
    show_expenses(  )
    
    show_expenses()
    check_budget()
    show_expenses()
    clear_fields()
    
def show_expenses():
    
    for row in table.get_children():
        table.delete(row)
        
    cursor.execute(
        "SELECT * FROM expenses"
    )
    
    rows = cursor.fetchall()
    
    for expense in rows:
        table.insert(
            "",
            tk.END,
            values = (
                expense[0],
                expense[1],
                expense[2],
                expense[3],
                expense[4]
            )
        )
    
def total_expense():
    
    cursor.execute(
        "SELECT SUM(amount) FROM expenses"
    )
    
    total = cursor.fetchone()[0]
    
    if total == None:
        total = 0
        
    result_label.config(
        text=f"Total Expense: ₹{total}"
    )
    
def delete_expense():
    confirm = messagebox.askyesno(
        "Delete",
        "Are you sure?"
    )
    if not confirm:
        return 
    
    expense_id = id_entry.get()
    
    cursor.execute(
        "DELETE FROM expenses WHERE id=?",
        (expense_id,)
    )
    
    conn.commit()
    
    result_label.config(
        text="Expense Deleted Successfully  "
    )
    id_entry.delete(0,tk.END)
    show_expenses()
    clear_fields()
    
def update_expense():
    
    expense_id = id_entry.get()

    title = entry_title.get()
    amount = entry_amount.get()
    category = entry_category.get()
    date = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    
    cursor.execute(
        "UPDATE expenses SET title=?, amount=?, category=?,date=? WHERE id=?",
        (title,amount,category,date,expense_id)
    )
    
    conn.commit()

    result_label.config(
        text="Expense Updated Successfully"
    )
    show_expenses()
    clear_fields()
    
def search_expense():
    keyword = search_entry.get()
    for row in table.get_children():
        table.delete(row)
    
    cursor.execute(
        "SELECT * FROM expenses WHERE title LIKE ? OR category LIKE ?",
        (f"%{keyword}%",f"%{keyword}%")
    )
    
    rows = cursor.fetchall()
    
    for expense in rows:
        table.insert(
            "",
            tk.END  ,
            values=(
                expense[0],
                expense[1],
                expense[2],
                expense[3],
                expense[4]      
            )
        )
            

def category_total():
    category = category_entry.get()
    
    cursor.execute(
        "SELECT SUM(amount) FROM expenses WHERE LOWER(category)=LOWER(?)", 
        (category,)
    )
    total=cursor.fetchone()[0]
    
    if total == None:
        total = 0
        
    result_label.config(
        text=f"{category} Expense: ₹{total}"
    )
    
def clear_fields():
    
    entry_title.delete(0, tk.END)
    entry_amount.delete(0, tk.END)
    entry_category.delete(0, tk.END)
    id_entry.delete(0, tk.END)
    search_entry.delete(0, tk.END)
    
    result_label.config(
        text=""
    )
    
def select_expense(event):
    
    selected = table.focus()

    data = table.item(selected)

    values = data["values"]

    if values:

        id_entry.delete(0, tk.END)
        id_entry.insert(0, values[0])

        entry_title.delete(0, tk.END)
        entry_title.insert(0, values[1])

        entry_amount.delete(0, tk.END)
        entry_amount.insert(0, values[2])

        entry_category.delete(0, tk.END)
        entry_category.insert(0, values[3])

def show_chart():
    
    cursor.execute(
        "SELECT category, SUM(amount) FROM expenses GROUP BY category"
    )

    data = cursor.fetchall()

    categories = []
    amounts = []

    for item in data:
        categories.append(item[0])
        amounts.append(item[1])

    plt.figure(figsize=(12,5))

    # Bar Chart
    plt.subplot(1,2,1)
    plt.bar(categories, amounts)
    plt.title("Expense Report")
    plt.xlabel("Category")
    plt.ylabel("Amount")

    # Pie Chart
    plt.subplot(1,2,2)
    plt.pie(
        amounts,
        labels=categories,
        autopct="%1.1f%%",
        startangle=90
    )
    plt.title("Category Distribution")

    plt.tight_layout()
    plt.show()
    
def export_csv():
    
    cursor.execute(
        "SELECT * FROM expenses"
    )
    rows = cursor.fetchall()
    
    file = open(
        "expense_report.csv",
        "w",
        newline=""
    )
    writer = csv.writer(file)
    
    writer.writerow(
        ["ID","Name","Amount","Category","Date"]
    )
    writer.writerows(rows)
    file.close()
    
    result_label.config(
        text="CSV Exported Successfully"
    )
    
def monthly_report():
    
    month = datetime.now().strftime("%m")
    
    cursor.execute(
        "SELECT SUM(amount) FROM expenses WHERE date LIKE ?",
        (f"%-{month}-%",)
    )
    
    total = cursor.fetchone ()[0]
    
    if total == None:
        total = 0
        
    result_label.config(
        text=f"This Month Expense: ₹{total}"
    )
    
def previous_month_report():
    
    cursor.execute(
        """
        SELECT SUM(amount) 
        FROM expenses 
        WHERE date LIKE '%-06-2026%'
        OR date LIKE '%-06-26%'
        """
    )

    total = cursor.fetchone()[0]

    if total is None:
        total = 0

    result_label.config(
        text=f"Previous Month Expense: ₹{total}"
    )
    
def check_budget():
    
    budget = budget_entry.get()
    if budget == "":
        return
    
    budget = int(budget)
    
    cursor.execute(
        "SELECT SUM(amount) FROM expenses"
    )
    total = cursor.fetchone()[0]
    
    if total is None:
        total = 0
    
    if total > budget:
        result_label.config(
            text=f"⚠ Budget Exceeded! Total = ₹{total}"
        )
    else:
        result_label.config(
            text=f"✅ Budget Remaining = ₹{budget-total}"
        )
        
def export_excel():
    workbook = Workbook()
    
    sheet = workbook.active
    sheet.title = "EXpenses"
    
    sheet.append([
        "ID",
        "Title",
        "Amount",
        "Category",
        "Date"
    ])
    
    cursor.execute(
        "SELECT * FROM expenses"
    )
    
    rows = cursor.fetchall()
    
    for row in rows:
        sheet.append(row)
        
    workbook.save("Expenses_Report.xlsx")
    
    result_label.config(
        text="Excel Exported Successfully"
    )

#GUI window
root = tk.Tk()
root.configure(bg="#EAF4FC")

root.title("Expense Tracker")
root.geometry("1000x750")

# Heading       
title_label = tk.Label(
    root,
    text="Expense Tracker",
    font=("Arial",20,"bold"),
    bg="#EAF4FC",
    fg="#1F4E79"
)

title_label.pack(pady=20)

#input frame            
input_frame = tk.Frame(
    root,
    bg="#EAF4FC"
    )
input_frame.pack(pady=10)


#Title Entry
tk.Label(input_frame,text="Expense Name",bg="#EAF4FC").grid(row=0,column=0,padx=5)
entry_title = tk.Entry(input_frame)
entry_title.grid(row=0,column=1,padx=5)


#Amount Entry
tk.Label(input_frame,text="Amount",bg="#EAF4FC").grid(row=1,column=0,padx=5)
entry_amount = tk.Entry(input_frame)
entry_amount.grid(row=1,column=1,padx=5)

#Category Entry
tk.Label(input_frame,text="Category",bg="#EAF4FC").grid(row=2,column=0,padx=5)
entry_category = tk.Entry(input_frame)
entry_category.grid(row=2,column=1,padx=5)

#Budget
tk.Label(root, text="Budget",bg="#EAF4FC").pack()

budget_entry = tk.Entry(root)
budget_entry.pack(pady=5)

#text frame
table_frame = tk.Frame(
    root,
    bg="#EAF4FC"
    )
table_frame.pack(pady=10)

columns = (
    "ID",
    "Name",
    "Amount",
    "Category",
    "Date"
)
table = ttk.Treeview(
    table_frame,
    columns=columns,
    show="headings"
)

for col in columns:
    table.heading(
        col,
        text=col
    )
    table.column(
        col,
        width=120
    )
    
scrollbar = ttk.Scrollbar(
    table_frame,
    orient="vertical",
    command=table.yview
)
table.configure(
    yscrollcommand=scrollbar.set
)
table.pack(
    side="left"
)

scrollbar.pack(
    side="right",
    fill="y"
)

table.bind(
    "<ButtonRelease-1>",
    select_expense
)

#button frame
button_frame = tk.Frame(
    root,
    bg="#EAF4FC"
    )
button_frame.pack(pady=10)

#view button
view_button = tk.Button(
    button_frame,
    text="View Expenses",
    command=show_expenses,
    bg="#4CAF50",
    fg="white"
)

view_button.pack(side="left",padx=5,pady=5)

#Add Expense
add_button = tk.Button(
    button_frame,
    text = "Add Expense",
    command=add_expense,
    bg="#4CAF50",
    fg="white"
)   

add_button.pack(side="left",padx=5,pady=5)

#total expense
total_button = tk.Button(
    button_frame,
    text = "Total Expense",
    command = total_expense,
    bg="#4CAF50",
    fg="white"
)

total_button.pack(side="left",padx=5,pady=5)

#category
tk.Label(root, text="Category for Total",bg="#EAF4FC").pack(  )
category_entry = tk.Entry(root)    
category_entry.pack(pady=5)

#category total
category_button = tk.Button(
    button_frame,
    text="Category Total",
    command=category_total,
    bg="#4CAF50",
    fg="white"
)
category_button.pack(side="left",padx=5,pady=5)

#id entry
tk.Label(root, text="Expense ID",bg="#EAF4FC").pack()
id_entry = tk.Entry(root)
id_entry.pack(pady=5)

#delete
delete_button = tk.Button(
    button_frame,
    text="Delete Expense",
    command=delete_expense,
    bg="#E53935",
    fg="white"
)
delete_button.pack(side="left",padx=5,pady=5)

#Update 
update_button = tk.Button(
    button_frame,
    text="Update Expense",
    command=update_expense,
    bg="#4CAF50",
    fg="white"  
)

update_button.pack(side="left",padx=5,pady=5)

#search entry
tk.Label(root, text="Search Expense",bg="#EAF4FC").pack()
search_entry = tk.Entry(root)
search_entry.pack(pady=5)

#search
search_button = tk.Button(
    button_frame,
    text="Search Expense",
    command=search_expense,
    bg="#4CAF50",
    fg="white"
)   
search_button.pack(side="left",padx=5,pady=5)

#Clear
clear_button = tk.Button(
    button_frame,
    text="Clear",
    command=clear_fields,
    bg="#4CAF50",
    fg="white"
)

clear_button.pack(side="left",padx=5,pady=5)

#chart
chart_button = tk.Button(
    button_frame,
    text="Show Chart",
    command=show_chart,
    bg="#4CAF50",
    fg="white"
)

chart_button.pack(
    side="left",
    padx=5,
    pady=5
)

export_button = tk.Button(
    button_frame,
    text="Export CSV",
    command = export_csv,
    bg="#4CAF50",
    fg="white"
)
export_button.pack(
    side="left",
    padx=5,
    pady=5
)

monthly_button = tk.Button(
    button_frame,
    text="Monthly_Report",
    command=monthly_report,
    bg="#4CAF50",
    fg="white"
)
monthly_button.pack(
    side="left",
    padx=5,
    pady=5
)

previous_button = tk.Button(
    button_frame,
    text="Previous Month",
    command=previous_month_report,
    bg="#4CAF50",
    fg="white"
)

previous_button.pack(
    side="left",
    padx=5
)

budget_button = tk.Button(
    button_frame,
    text="Check Budget",
    command=check_budget,
    bg="#4CAF50",
    fg="white"
)

budget_button.pack(side="left",padx=5)

excel_button = tk.Button(
    button_frame,
    text="Export Excel",
    command=export_excel,
    bg="#4CAF50",
    fg="white"
)
excel_button.pack(side="left",padx=5)

#result_label
result_label = tk.Label(
    root,   
    text="",
    bg="#EAF4FC"
)
result_label.pack(pady=10) 

show_expenses()
root.mainloop()