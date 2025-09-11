import pandas as pd
from tkinter import *
from tkinter import ttk
from Backend import *
from tkinter import filedialog

#Select file with open folder
folderData = r"C:\KP\Python ENV\ExcelData"
listFiles = list_files_in_folder(folderData)

def load_data(data):
    # Clear existing data
    tree.delete(*tree.get_children())

    # Insert rows into the treeview
    for _, row in data.iterrows():
        tree.insert("", "end", values=list(row))

root = Tk()
root.title("P2P Form Automation GUI")
root.configure(bg="#2e2e2e")

frame = Frame(root, bg="#2e2e2e")
frame.pack(padx=20, pady=20, fill="both", expand=True)

# Data from importing
columns = [
    "A", "Time Stamp", "Miliseconds", "System time stamp", "System Miliseconds",
    "RTU Name", "B1", "B2", "B3", "Element", "Description", "Status", "Source",
    "Priority", "Tag", "Operator", "Message Class", "Comment", "User Comment", "SoE"
]  

# Actual needed data for export telesignal (TC)
needed_data = [
    "Time Stamp","Miliseconds", "B1", "B2", "B3", "RTU Name","Element", "Description", "Status", "Source",
]


# Create Treeview inside a container frame
tree_frame = Frame(frame, bg="#2e2e2e")
tree_frame.pack()

# --- File selection frame ---
file_frame = Frame(frame, bg="#2e2e2e")
file_frame.pack(anchor="w", pady=(0, 15))

# Label
file_label = Label(file_frame, text="Select HIS File:", bg="#2e2e2e", fg="#fff")
file_label.pack(side=LEFT)

# File selection using file dialog

selected_file = StringVar()

def browse_file():
    file_path = filedialog.askopenfilename(
        initialdir=folderData,
        title="Select file",
        filetypes=(("Excel files", "*.ods"), ("All files", "*.*"))
    )
    if file_path:
        selected_file.set(file_path)

file_entry = Entry(file_frame, textvariable=selected_file, width=50, state="readonly")
file_entry.pack(side=LEFT, padx=(5, 5))

browse_btn = Button(file_frame, text="Browse...", command=browse_file, bg="#444", fg="#fff",
                    activebackground="#555", activeforeground="#fff")
browse_btn.pack(side=LEFT, padx=(0, 10))

# Load button (beside Combobox)
load_btn = Button(file_frame, text="Load Data", 
                  command=lambda: load_data(loadTelesignalData(selected_file.get())),
                  bg="#444", fg="#fff", activebackground="#555", activeforeground="#fff")
load_btn.pack(side=LEFT)

# --- "+" button: Move selected rows from first table to second table ---

# --- Prevent duplicates in second table ---
def move_selected_to_second_table():
    selected_items = tree.selection()
    existing = set(tuple(tree2.item(child, "values")) for child in tree2.get_children())
    for item in selected_items:
        row_values = tree.item(item, "values")
        col_indices = [columns.index(col) for col in needed_data]
        filtered_values = tuple(row_values[i] for i in col_indices)
        if filtered_values not in existing:
            tree2.insert("", "end", values=filtered_values)
            existing.add(filtered_values)

# Place "+" button beside Load button
add_btn = Button(
    file_frame, text="↓", bg="#2e2e2e", fg="#0f0", font=("Arial", 14, "bold"),
    activebackground="#555", activeforeground="#0f0", command=move_selected_to_second_table
)
add_btn.pack(side=LEFT, padx=(5,0))

# --- "-" button: Remove selected rows from second table ---
def remove_selected_from_second_table():
    selected_items = tree2.selection()
    for item in selected_items:
        tree2.delete(item)

remove_btn = Button(
    file_frame, text="-", bg="#2e2e2e", fg="#f00", font=("Arial", 14, "bold"),
    activebackground="#555", activeforeground="#f00", command=remove_selected_from_second_table
)
remove_btn.pack(side=LEFT, padx=(5,0))

dataTelesignal = loadTelesignalData(selected_file.get() if selected_file.get() else listFiles[0])

# Treeview aka table (for imported data)
tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=10)

for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=150)

# Attach scrollbars
vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
tree.configure(yscroll=vsb.set, xscroll=hsb.set)

# Layout with grid so scrollbars sit nicely
tree.grid(row=0, column=0, sticky="nsew")
vsb.grid(row=0, column=1, sticky="ns")
hsb.grid(row=1, column=0, sticky="ew")

tree_frame.grid_rowconfigure(0, weight=1)
tree_frame.grid_columnconfigure(0, weight=1)


# --- Second table (below the first one) ---
tree_frame2 = Frame(frame, bg="#2e2e2e")
tree_frame2.pack(fill="both", expand=True, pady=(10,0))

tree2 = ttk.Treeview(tree_frame2, columns=needed_data, show="headings", height=10)
for col in needed_data:
    tree2.heading(col, text=col)
    tree2.column(col, width=120)

vsb2 = ttk.Scrollbar(tree_frame2, orient="vertical", command=tree2.yview)
hsb2 = ttk.Scrollbar(tree_frame2, orient="horizontal", command=tree2.xview)
tree2.configure(yscroll=vsb2.set, xscroll=hsb2.set)

tree2.grid(row=0, column=0, sticky="nsew")
vsb2.grid(row=0, column=1, sticky="ns")
hsb2.grid(row=1, column=0, sticky="ew")

tree_frame2.grid_rowconfigure(0, weight=1)
tree_frame2.grid_columnconfigure(0, weight=1)

# --- Make cells editable on double-click ---
def on_tree2_double_click(event):
    region = tree2.identify("region", event.x, event.y)
    if region != "cell":
        return
    row_id = tree2.identify_row(event.y)
    col_id = tree2.identify_column(event.x)
    if not row_id or not col_id:
        return

    col_index = int(col_id.replace("#", "")) - 1
    col_name = needed_data[col_index]
    x, y, width, height = tree2.bbox(row_id, col_id)
    value = tree2.set(row_id, col_name)

    # Create Entry widget overlay
    entry = Entry(tree2, width=width//8, bg="#333", fg="#fff", borderwidth=0, highlightthickness=1, relief="flat")
    entry.place(x=x, y=y, width=width, height=height)
    entry.insert(0, value)
    entry.focus_set()

    def save_edit(event=None):
        new_value = entry.get()
        tree2.set(row_id, col_name, new_value)
        entry.destroy()

    def cancel_edit(event=None):
        entry.destroy()

    entry.bind("<Return>", save_edit)
    entry.bind("<FocusOut>", save_edit)
    entry.bind("<Escape>", cancel_edit)

tree2.bind("<Double-1>", on_tree2_double_click)

# --- Export button (does nothing yet) ---
export_btn = Button(
    frame, text="Export", bg="#444", fg="#fff",
    activebackground="#555", activeforeground="#fff",
    font=("Arial", 12, "bold")
)
export_btn.pack(anchor="e", pady=(10, 0))


# Style
style = ttk.Style()
style.theme_use("default")
style.configure("Treeview",
                background="#222",
                foreground="#fff",
                fieldbackground="#222",
                rowheight=25)
style.map("Treeview",
          background=[("selected", "#444")],
          foreground=[("selected", "#fff")])

root.mainloop()
