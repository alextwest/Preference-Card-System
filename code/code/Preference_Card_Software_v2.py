# Preference Card Software
# Designed for LA County Hospital Surgical Unit
# Authored By Alex West

import glob
import os
import os.path
import sys
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, simpledialog

import openpyxl
import pandas as pd


class Context:

    soft_goods_key = "soft goods"
    instruments_key = "instruments"
    sheet_columns = ['Quantity', 'Service', 'Container Name', 'Item Description', 'Vendor Part #', 'Hold']
    instrument_columns = ['Quantity', 'Service', 'Container Name', 'Hold']
    soft_goods_columns = ['Quantity', 'Item Description', 'Vendor Part #', 'Hold']

    root = None
    data_in_directory = "."
    data_out_directory = "."
    data = {}
    windows = {}
    validation = None
    container_file = None
    soft_goods_file = None

    def __init__(self, window, in_directory, out_directory):
        self.root = window
        self.validation = window.register(only_digits)
        window.withdraw()
        self.data_in_directory = in_directory
        self.data_out_directory = out_directory

    def initialize(self):
        self.data = {}
        self.remove_all_windows()

    def get_soft_goods_key(self):
        return self.soft_goods_key

    def get_instruments_key(self):
        return self.instruments_key

    def get_sheet_columns(self):
        return self.sheet_columns

    def get_instrument_columns(self):
        return self.instrument_columns

    def get_soft_goods_columns(self):
        return self.soft_goods_columns

    def get_validation(self):
        return self.validation

    def get_data(self, key):
        if key in self.data:
            return self.data[key]
        return None

    def set_data(self, key, data):
        self.data[key] = data

    def get_root_window(self):
        return self.root

    def get_container_file(self):
        return self.container_file

    def get_soft_goods_file(self):
        return self.soft_goods_file

    def set_container_file(self, container_file):
        self.container_file = container_file

    def set_soft_goods_file(self, soft_goods_file):
        self.soft_goods_file = soft_goods_file

    def start(self):
        self.root.mainloop()

    def user_start(self):

        self.data = {}

        self.root.deiconify()

        self.root.title("Preference Card System")
        self.root.geometry("500x300")

        tk.Label(self.root,
                 text="Would you like to create a new preference card or edit an existing one?").pack(pady=10)
        tk.Button(self.root,
                  text="Create New", command=select_surgery_service).pack(pady=5)
        tk.Button(self.root,
                  text="Edit Existing", command=select_editable_preference_card_file).pack(pady=5)
        tk.Button(self.root,
                  text="Exit", command=exit_app).pack(pady=5)

    def get_in_directory(self):
        return self.data_in_directory

    def get_out_directory(self):
        return self.data_out_directory

    def new_window(self, key=None):
        window = tk.Toplevel(self.get_root_window())
        if key is not None:
            if key in self.windows.keys():
                self.windows[key].destroy()
            self.windows[key] = window
        return window

    def remove_all_windows(self):
        for key in self.windows.keys():
            try:
                # was 'withdraw' but I think will actually kill the window which we don't need anymore,
                # just don't destroy the root.
                self.windows[key].destroy()
            # maybe find out the Exception class that can be thrown
            except Exception:
                pass
        self.windows = {}

    def exit(self):
        self.root.destroy()
        exit_app()


global context


def make_quantity_entry_widget(frame, width=5):
    # link this to the associated checkbox so that the checkbox xan be deactivated/activated based on having content
    return tk.Entry(frame, width=width, validate="key", validatecommand=(context.get_validation(), "%S"))


def only_digits(char):
    # for number entry fields (see Context)
    return char.isdigit()


def show_error(message):
    messagebox.showerror("Error", message)


def read_excel_file_as_dataframe(file_path):
    try:
        return pd.read_excel(file_path)
    except Exception as e:
        show_error(f"Failed to read Excel file: {e}")
    return pd.DataFrame()


def load_excel_file(file_path, group, container_name):
    # Read the Excel file
    df = read_excel_file_as_dataframe(file_path)
    if df.empty:
        return df
        # Group values in column B based on categories in column A
    try:
        return df.groupby(group)[container_name].apply(list).reset_index()
    except Exception as e:
        show_error(f"Failed to parse Excel file: {e}")


def select_excel_file(directory, title="Select Excel File", default_file=None):
    file_path = filedialog.askopenfilename(initialdir=directory, title=title,
                                           filetypes=(("Excel", "*.xlsx"), ("Excel", "*.xls")),
                                           initialfile=default_file)
    return file_path


# Function to update scroll region when resizing
def on_frame_configure(canvas):
    canvas.configure(scrollregion=canvas.bbox("all"))


def convert_instrument_data(instrument_entries):
    instruments_data = []
    for container_label, (an_entry, checkbox, frame) in instrument_entries.items():
        container_name = container_label["text"].split(': ')[1]
        service_name = container_label["text"].split(': ')[0]
        quantity = an_entry.get()
        hold = checkbox.get()
        if quantity.strip():
            instruments_data.append([int(quantity), service_name, container_name, None, None, hold])
    return pd.DataFrame(instruments_data, columns=context.get_sheet_columns())


# Convert the data contained within UI widgets
def convert_soft_goods_data(soft_goods_entries):
    soft_goods_data = []
    if soft_goods_entries:
        for (item_descr, vendor_part_num), (an_entry, checkbox, frame) in soft_goods_entries.items():
            quantity = an_entry.get()
            hold = checkbox.get()
            if quantity.strip():
                soft_goods_data.append([int(quantity), None, None, item_descr, vendor_part_num, hold])
    return pd.DataFrame(soft_goods_data, columns=context.get_sheet_columns())


def export_to_excel(directory, instrument_dataframe, soft_goods_dataframe):

    # Concatenate the DataFrames
    combined_df = pd.concat([instrument_dataframe, soft_goods_dataframe], ignore_index=True)

    if combined_df.empty:
        show_error("Nothing to save!")
        return

    # Prompt user for doctor's name
    doctor_name = simpledialog.askstring("Enter Doctor's Name", "Please enter the doctor's name:")
    if doctor_name is not None:
        # Create the file path based on doctor's name
        file_path = os.path.join(directory, f"{doctor_name}.xlsx")

        # Check if the file already exists
        if os.path.isfile(file_path):
            # File exists, ask user if they want to append to existing file
            choice = messagebox.askyesno(
                "File Exists",
                f"File '{doctor_name}.xlsx' already exists. Do you want to append to this file?")
            if choice:
                # Append to existing file
                with pd.ExcelWriter(file_path, mode='a') as writer:
                    # Prompt user for surgery name
                    surgery_name = \
                        simpledialog.askstring("Enter Name of Surgery",
                                               "Please enter the name of the surgery performed:")
                    sheet_name = f"{surgery_name}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}"
                    combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
                messagebox.showinfo(
                    "Export Successful",
                    f"Selected instruments appended to '{sheet_name}' in '{doctor_name}.xlsx' successfully.")

                ask_restart()
            else:
                # Ask for new file name
                file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                         filetypes=(("Excel files", "*.xlsx"),
                                                                    ("All files", "*.*")),
                                                         initialfile=f"{doctor_name}.xlsx",
                                                         title="Save As")
                if file_path:
                    combined_df.to_excel(file_path, index=False)
                    messagebox.showinfo("Export Successful", "Selected instruments exported successfully.")

                    ask_restart()
                else:
                    messagebox.showwarning("Warning", "Operation canceled.")
        else:
            # File does not exist, ask for confirmation to create new file
            choice = messagebox.askyesno(
                "File Not Found",
                f"File '{doctor_name}.xlsx' does not exist. Do you want to create a new file?")
            if choice:
                # Prompt user for surgery name
                surgery_name = \
                    simpledialog.askstring("Enter Name of Surgery",
                                           "Please enter the name of the surgery performed:")
                sheet_name = f"{surgery_name}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}"
                combined_df.to_excel(file_path, sheet_name=sheet_name, index=False)
                messagebox.showinfo(
                    "Export Successful",
                    f"Selected instruments exported to '{sheet_name}' in '{doctor_name}.xlsx' successfully.")

                ask_restart()
            else:
                messagebox.showwarning("Warning", "Operation canceled.")


def filter_soft_goods(event, widgets):
    text = event.widget.get().lower()
    for label, frame in reversed(widgets):
        if text not in label["text"].lower():
            if frame.winfo_ismapped():
                frame.grid_remove()
        else:
            if not frame.winfo_ismapped():
                frame.grid()


def select_soft_goods(filename, selected_soft_goods_data=None):
    soft_goods_df = read_excel_file_as_dataframe(filename)
    soft_goods_frame, soft_goods_window = create_container_window("Select Soft Goods")
    if selected_soft_goods_data is not None:
        preselected = selected_soft_goods_data.set_index(['Item Description', 'Vendor Part #'])
    else:
        preselected = None

    entries = {}

    count = 0
    widgets = []
    for index, row in soft_goods_df.sort_values(by="ITEM DESCRIPTION").iterrows():
        item_description = row['ITEM DESCRIPTION']
        vendor_part_number = row['VENDOR PART#']

        entry_frame = tk.Frame(soft_goods_frame)

        label = tk.Label(entry_frame, text=f"{item_description}, Vendor Part #: {vendor_part_number}",
                         width=50, anchor="w", wraplength=400, justify="left")
        entry = make_quantity_entry_widget(entry_frame, width=5)

        check_var = tk.BooleanVar()
        checkbox = tk.Checkbutton(entry_frame, width=1, height=1, variable=check_var, onvalue=True, offvalue=False)
        label.grid(row=0, column=0)
        entry.grid(row=0, column=1)
        checkbox.grid(row=0, column=2)
        entries[(item_description, vendor_part_number)] = (entry, check_var, entry_frame)
        entry_frame.grid(row=count, column=0)
        widgets.append([label, entry_frame])
        key = (item_description, vendor_part_number)
        if preselected is not None:
            if key in preselected.index:
                quantity = preselected.loc[key, 'Quantity']
                hold = preselected.loc[key, 'Hold']
                entry.insert(0, str(quantity))
                if hold:
                    checkbox.select()
        count = count+1

    def done_command():
        context.set_data(context.get_soft_goods_key(), convert_soft_goods_data(entries))
        hide_window(soft_goods_window)

    done_soft_goods_button = tk.Button(soft_goods_window, text="Done", command=done_command)
    done_soft_goods_button.pack(padx=10, pady=10)
    cancel_soft_goods_button = tk.Button(soft_goods_window, text="Cancel",
                                         command=lambda: hide_window(soft_goods_window))
    cancel_soft_goods_button.pack(padx=10, pady=10)

    search_frame = tk.Frame(soft_goods_window)
    search_frame.pack(fill=tk.Y, expand=True)

    search_label = tk.Label(search_frame, width=8, text="Search: ")
    search_label.grid(row=0, column=0)
    search_entry = tk.Entry(search_frame, width=12)
    search_entry.grid(row=0, column=1)
    search_entry.bind("<Return>", lambda event: filter_soft_goods(event, widgets))


def filter_instruments(event, instrument_entries):
    text = event.widget.get().lower()
    for key in reversed(instrument_entries.keys()):
        entry, checkbox, frame = instrument_entries[key]
        if text not in key["text"].lower():
            if frame.winfo_ismapped():
                frame.grid_remove()
        else:
            if not frame.winfo_ismapped():
                frame.grid()


def select_instruments(grouped_data, types, selected_instrument_data=None):
    container_window, instrument_entries = (
        layout_window("Select Containers", grouped_data, types, selected_instrument_data))

    def cancel_window():
        hide_window(container_window)

    # display select soft goods button on window with container selections to move on to the next step
    soft_goods_button = tk.Button(container_window, text="Select Soft Goods",
                                  command=lambda: select_soft_goods(context.get_soft_goods_file(),
                                                                    context.get_data(context.get_soft_goods_key())))
    soft_goods_button.pack(padx=10, pady=10)
    export_to_excel_button = tk.Button(container_window, text="Export to Excel",
                                       command=lambda: export_to_excel(context.get_out_directory(),
                                                                       convert_instrument_data(instrument_entries),
                                                                       context.get_data(context.get_soft_goods_key())))
    export_to_excel_button.pack(padx=10, pady=10)
    cancel_instruments_button = tk.Button(container_window, text="Cancel", command=cancel_window)
    cancel_instruments_button.pack(padx=10, pady=10)

    search_frame = tk.Frame(container_window)
    search_frame.pack(fill=tk.Y, expand=True)

    search_label = tk.Label(search_frame, width=8, text="Search: ")
    search_label.grid(row=0, column=0)
    search_entry = tk.Entry(search_frame, width=12)
    search_entry.grid(row=0, column=1)
    search_entry.bind("<Return>", lambda event: filter_instruments(event, instrument_entries))


def layout_window(title, grouped_data, types, selected_instrument_data, dimensions="800x800"):

    selected_containers_frame, container_window = create_container_window(title, dimensions)

    entries = {}

    grouped_data_keys = grouped_data.keys()

    if selected_instrument_data is not None:
        preselected = selected_instrument_data.set_index(['Service', 'Container Name'])
    else:
        preselected = None

    row = 0
    for a_type in types:
        # Get the index of the row where 'Service' matches selected_service
        index = grouped_data[grouped_data[grouped_data_keys[0]] == a_type].index.tolist()
        # Retrieve container names for the specified service index
        container_names = grouped_data.loc[index, grouped_data_keys[1]].iloc[0]
        # Insert each container name into the listbox
        #        row = 0
        for container_name in container_names:
            container_label = f"{a_type}: {container_name}"
            # instrument_frame = tk.Frame(selected_containers_frame)
            # instrument_frame.pack(fill=tk.X)

            instrument_frame = tk.Frame(selected_containers_frame)

            label = tk.Label(instrument_frame, text=container_label,
                             width=50, anchor="w", wraplength=400, justify="left")
            quantity_entry = make_quantity_entry_widget(instrument_frame, width=5)
            check_var = tk.BooleanVar()
            checkbox = tk.Checkbutton(instrument_frame, width=1, height=1, variable=check_var,
                                      onvalue=True, offvalue=False)
            label.grid(row=0, column=0)
            quantity_entry.grid(row=0, column=1)
            checkbox.grid(row=0, column=2)
            instrument_frame.grid(row=row, column=0)
            entries[label] = (quantity_entry, check_var, instrument_frame)
            key = (a_type, container_name)
            if preselected is not None:
                if key in preselected.index:
                    quantity = preselected.loc[key, 'Quantity']
                    hold = preselected.loc[key, 'Hold']
                    quantity_entry.insert(0, str(quantity))
                    if hold:
                        checkbox.select()
            row = row + 1

    return container_window, entries


def create_container_window(title, dimensions="800x800"):
    container_window = new_window(title)
    container_window.title(title)

    # Set the size of the window
    container_window.geometry(dimensions)

    # Create a canvas for scrollable view
    canvas = tk.Canvas(container_window)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # Add a scrollbar to the canvas
    scrollbar = tk.Scrollbar(container_window, orient=tk.VERTICAL, command=canvas.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # Configure the canvas to scroll with the scrollbar
    canvas.configure(yscrollcommand=scrollbar.set)

    # Create a frame to contain the list of instruments
    selected_containers_frame = tk.Frame(canvas)
    selected_containers_frame.pack(fill=tk.BOTH, expand=True)

    # Add the frame to the canvas
    canvas.create_window((0, 0), window=selected_containers_frame, anchor=tk.NW)

    selected_containers_frame.bind("<Configure>", lambda event: on_frame_configure(canvas))

    return selected_containers_frame, container_window


def select_surgery_service(selected_types=None):
    if selected_types is None:
        selected_types = []
    grouped_data = load_excel_file(context.get_container_file(), 'Service', 'Container Name')

    context.remove_all_windows()
    title = "Select Service"
    service_window = new_window(title)
    service_window.title(title)

    # Set the size of the window
    service_window.geometry("400x500")

    selected_services = tk.Listbox(service_window, selectmode=tk.MULTIPLE)
    selected_services.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

    services = grouped_data['Service'].unique()  # Access unique values in 'Service' column directly
    # Insert service options into the listbox\
    for service in services:
        selected_services.insert(tk.END, service)
        if service in selected_types:
            selected_services.select_set(tk.END)

    def select():
        # Retrieve the indices of selected items
        selected_indices = selected_services.curselection()
        # Retrieve the selected services based on the indices
        selected = [selected_services.get(idx) for idx in selected_indices]
        # Pass the selected services to the next function
        if selected:
            select_instruments(grouped_data, selected, context.get_data(context.get_instruments_key()))
        else:
            messagebox.showinfo("Nothing selected", "Please select at least one service first.")

    select_button = tk.Button(service_window, text="Select", command=select)
    select_button.pack(padx=10, pady=10)
    exit_button = tk.Button(service_window, text="Cancel", command=lambda: hide_window(service_window))
    exit_button.pack(padx=10, pady=10)


def select_editable_preference_card_file():
    select_surgery_sheet(context.get_out_directory())


def select_surgery_sheet(directory):
    file_path = filedialog.askopenfilename(initialdir=directory,
                                           title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        sheet_names = get_sheet_names(file_path)
        show_sheet_selection_window(file_path, sheet_names)


def get_sheet_names(excel_file):
    try:
        workbook = openpyxl.load_workbook(excel_file, read_only=True)
        sheet_names = workbook.sheetnames
        workbook.close()
        return sheet_names
    except Exception as e:
        show_error(f"Failed to read Excel file: {e}")
        return []


def show_sheet_selection_window(file_path, sheet_names):
    def on_select(event):
        selected_sheet.set(sheet_listbox.get(sheet_listbox.curselection()))
        selection_window.destroy()
        messagebox.showinfo("Sheet Selected", f"You have selected the sheet: {selected_sheet.get()}")
        types, selected_container_data, selected_soft_goods_data = process_sheet(file_path, selected_sheet.get())
        context.set_data(context.get_instruments_key(), selected_container_data)
        context.set_data(context.get_soft_goods_key(), selected_soft_goods_data)
        select_surgery_service(types)

    selection_window = tk.Toplevel()
    selection_window.title("Select Sheet")
    selection_window.geometry("400x300")

    label = tk.Label(selection_window, text="Available sheets:")
    label.pack(pady=10)

    selected_sheet = tk.StringVar()

    sheet_listbox = tk.Listbox(selection_window, selectmode=tk.SINGLE)
    for sheet in sheet_names:
        sheet_listbox.insert(tk.END, sheet)
    sheet_listbox.pack(pady=10)
    sheet_listbox.bind('<<ListboxSelect>>', on_select)


def process_sheet(file_path, sheet_name):

    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # Assuming the columns are in the order:
        # Quantity, Service, Container Name, Item Description, Vendor Part #, Hold
        selected_container_data = df[context.get_instrument_columns()]
        selected_soft_goods_data = df[context.get_soft_goods_columns()]

        # Filter selected_container_data to drop rows where 'Service' is NaN
        selected_container_data = selected_container_data.dropna(subset=['Service'])
        selected_soft_goods_data = selected_soft_goods_data.dropna(subset=['Vendor Part #'])

        # Get unique types excluding NaN
        types = selected_container_data['Service'].dropna().unique()

        return types, selected_container_data, selected_soft_goods_data

    except Exception as e:
        show_error(f"Failed to process selected preference card: {e}")


def new_window(key=None):
    return context.new_window(key)


def hide_window(window):
    window.withdraw()


def ask_restart():
    choice = messagebox.askquestion("Restart", "Would you like to select another preference card?", icon="question")
    if choice == 'yes':
        # Restart the program
        context.initialize()
    else:
        # Quit the program
        context.exit()


def exit_app():
    sys.exit(0)


def user_start():
    context.user_start()


def confirm_files():
    in_directory = context.get_in_directory()
    title = "Confirm Files"
    confirmation_window = new_window(title)
    confirmation_window.title(title)
    confirmation_window.geometry("800x800")
    tk.Label(confirmation_window, text="Are these the files you expected?").pack(pady=10)

    canvas = tk.Canvas(confirmation_window)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar = tk.Scrollbar(confirmation_window, orient=tk.VERTICAL, command=canvas.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    confirm_frame = tk.Frame(canvas)
    canvas.create_window((0, 0), window=confirm_frame, anchor=tk.NW)
    confirm_frame.bind("<Configure>", lambda event: on_frame_configure(canvas))

    def update_preview():
        container_file = context.get_container_file()
        soft_goods_file = context.get_soft_goods_file()
        preview_text = ""
        for (file_type, file_path) in [('Container File', container_file), ('Soft Goods File', soft_goods_file)]:
            if file_path is not None:
                df = pd.read_excel(file_path, nrows=10)
                preview_text += f"Preview of {file_type}: {file_path}:\n\n{df.head(10)}\n\n\n"
            else:
                preview_text += f"No {file_type} found!\n\n\n"
        preview_label.config(text=preview_text)

    preview_label = tk.Label(confirm_frame, text="", justify=tk.LEFT)
    preview_label.pack(pady=10)
    update_preview()

    def on_confirm():

        def missing_file(message):
            show_error(f"No {message} identified!")

        container_file = context.get_container_file()
        soft_goods_file = context.get_soft_goods_file()
        if container_file is None:
            if soft_goods_file is None:
                missing_file('container file or soft goods file')
            else:
                missing_file('container file')
        else:
            if soft_goods_file is None:
                missing_file('soft goods file')
            else:
                confirmation_window.destroy()
                user_start()

    def not_confirm_container_file():
        corrected_file = select_excel_file(in_directory)
        if corrected_file:
            df = pd.read_excel(corrected_file)
            if not {'Service', 'Container Name', 'Reference ID'}.issubset(df.columns):
                show_error("The selected file does not have the expected columns for a container file.")
                return
            context.set_container_file(corrected_file)
            update_preview()

    def not_confirm_soft_goods_file():
        corrected_file = select_excel_file(in_directory)
        if corrected_file:
            df = pd.read_excel(corrected_file)
            if not {'ITEM DESCRIPTION', 'VENDOR PART#'}.issubset(df.columns):
                show_error("The selected file does not have the expected columns for a soft goods file.")
                return
            context.set_soft_goods_file(corrected_file)
            update_preview()

    change_container_button = (
        tk.Button(confirm_frame, text="Select Correct Instrument Container File", command=not_confirm_container_file))
    change_container_button.pack(pady=10)

    change_soft_goods_button = (
        tk.Button(confirm_frame, text="Select Correct Soft Goods File", command=not_confirm_soft_goods_file))
    change_soft_goods_button.pack(pady=10)

    confirm_button = tk.Button(confirm_frame, text="Confirm", command=on_confirm)
    confirm_button.pack(pady=10)

    exit_button = tk.Button(confirm_frame, text="Exit", command=exit_app)
    exit_button.pack(pady=10)


def select_and_confirm_files():
    directory = context.get_in_directory()  # Directory where information files live
    files = glob.glob(f"{directory}/*.xlsx")  # Find the full path for the in directory

    container_file = None
    soft_goods_file = None
    # Search for the 2 files with the specific colum headers to determine which file contains which information
    for file in files:
        df = pd.read_excel(file)
        if {'Service', 'Container Name', 'Reference ID'}.issubset(df.columns):
            container_file = file
            if soft_goods_file is not None:
                break
        elif {'ITEM DESCRIPTION', 'VENDOR PART#'}.issubset(df.columns):
            soft_goods_file = file
            if container_file is not None:
                break

    # If both files are found, move on to user confirmation
    if container_file and soft_goods_file:
        context.set_container_file(container_file)
        context.set_soft_goods_file(soft_goods_file)
    confirm_files()


def main(args):
    in_directory = "."
    out_directory = "."
    if args:
        in_directory = os.path.abspath(args[0])
        os.makedirs(in_directory, exist_ok=True)
        if args[1:]:
            out_directory = os.path.abspath(args[1])
            os.makedirs(out_directory, exist_ok=True)
        else:
            out_directory = in_directory
    global context
    root = tk.Tk()
    context = Context(root, in_directory, out_directory)
    select_and_confirm_files()
    context.start()


if __name__ == "__main__":
    main(sys.argv[1:])
