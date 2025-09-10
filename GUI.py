import pandas as pd
from tkinter import *
from tkinter import ttk
from Backend import *

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

# Assuming each item in dataTelesignal is a tuple/list of columns
columns = [
    "A", "Time Stamp", "Miliseconds", "System time stamp", "System Miliseconds",
    "RTU Name", "B1", "B2", "B3", "Element", "Description", "Status", "Source",
    "Priority", "Tag", "Operator", "Message Class", "Comment", "User Comment", "SoE"
]  # Replace with your actual column names

# Create Treeview inside a container frame
tree_frame = Frame(frame, bg="#2e2e2e")
tree_frame.pack()


# --- File selection frame ---
file_frame = Frame(frame, bg="#2e2e2e")
file_frame.pack(anchor="w", pady=(0, 15))

# Label
file_label = Label(file_frame, text="Select File:", bg="#2e2e2e", fg="#fff")
file_label.pack(side=LEFT)

# Combobox
selected_file = StringVar()
file_combo = ttk.Combobox(file_frame, textvariable=selected_file, values=listFiles, state="readonly", width=50)
file_combo.pack(side=LEFT, padx=(5, 10))

# Load button (beside Combobox)
load_btn = Button(file_frame, text="Load Data", 
                  command=lambda: load_data(loadTelesignalData(selected_file.get())),
                  bg="#444", fg="#fff", activebackground="#555", activeforeground="#fff")
load_btn.pack(side=LEFT)


dataTelesignal = loadTelesignalData(selected_file.get() if selected_file.get() else listFiles[0])

tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=30)

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
