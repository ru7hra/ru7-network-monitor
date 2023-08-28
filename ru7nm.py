import psutil
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from tkinter import font

class NetworkConnectionsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ru7 Network Monitor")

        self.root.geometry("800x600")  # Set initial window size

        self.style = ttk.Style()
        self.style.theme_use("clam")  # Set the theme (can be changed)

        self.connections_list = ttk.Treeview(root, columns=("Local", "Remote", "Status", "Process"))
        self.connections_list.bind("<ButtonRelease-1>", self.on_tree_click)  # Single click
        self.connections_list.bind("<Double-Button-1>", self.on_double_click)  # Double click

        for column in ("Local", "Remote", "Status", "Process"):
            self.connections_list.heading(column, text=column, command=lambda col=column: self.sort_column(col))

        self.scrollbar = tk.Scrollbar(root, orient="vertical", command=self.connections_list.yview)
        self.connections_list.configure(yscrollcommand=self.scrollbar.set)

        self.connections_list.pack(fill="both", expand=True, padx=20, pady=(10, 0))
        self.scrollbar.pack(side="right", fill="y")

        self.system_processes = []
        self.user_processes = []

        self.refresh_button = ttk.Button(root, text="\u21BB", command=self.update_connections)
        self.refresh_button.pack(pady=(10, 5))

        self.export_button = ttk.Button(root, text="\u2B07 Download", command=self.export_data)
        self.export_button.pack(pady=(0, 10))

        self.search_entry = ttk.Entry(root)
        self.search_entry.pack(fill="x", padx=20, pady=5)

        self.search_button = ttk.Button(root, text="Search", command=self.search_connections)
        self.search_button.pack(pady=(0, 10))

        self.selected_items = set()  # Set to store selected item IDs

        self.update_connections()

        # Set Theme and font
        self.custom_font = font.Font(family="Courier New", size=12)
        self.connections_list.option_add("*Treeview*Font", self.custom_font)

    def update_connections(self):
        self.system_processes.clear()
        self.user_processes.clear()

        for conn in psutil.net_connections(kind="inet"):
            local_address = self.format_address(conn.laddr)
            remote_address = "N/A" if conn.raddr is None else self.format_address(conn.raddr)
            status = conn.status
            process_name = self.get_process_name(conn.pid)

            if self.is_system_process(process_name):
                self.system_processes.append((local_address, remote_address, status, process_name))
            else:
                self.user_processes.append((local_address, remote_address, status, process_name))

        self.sort_connections()

    def sort_connections(self):
        self.connections_list.delete(*self.connections_list.get_children())

        for i, conn in enumerate(self.system_processes + self.user_processes):
            item_id = f"I{i:03}"  # Creating unique item ID
            self.connections_list.insert("", "end", item_id, values=conn)
            self.connections_list.item(item_id, tags=(item_id,))

    def sort_column(self, column):
        if column == "Process":
            self.system_processes.sort(key=lambda x: x[3])
            self.user_processes.sort(key=lambda x: x[3])
        else:
            col_index = self.column_index(column)
            self.system_processes.sort(key=lambda x: x[col_index])
            self.user_processes.sort(key=lambda x: x[col_index])

        self.sort_connections()

    def column_index(self, column):
        return {"Local": 0, "Remote": 1, "Status": 2}.get(column, 0)

    def get_process_name(self, pid):
        try:
            process = psutil.Process(pid)
            return process.name()
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            return "N/A"

    def is_system_process(self, process_name):
        return process_name.lower() in ["system", "system idle process"]

    def export_data(self):
        if not self.selected_items:
            messagebox.showerror("Error", "No items selected for export.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[
            ("Text Files", "*.txt"), ("CSV Files", "*.csv"), ("Excel Files", "*.xlsx"), ("HTML Files", "*.html")
        ])

        if file_path:
            export_format = file_path.split(".")[-1]
            data = []

            for item_id in self.selected_items:
                try:
                    values = self.connections_list.item(item_id, "values")
                    data.append(values)
                except tk.TclError:
                    pass  # Skip items that might have been deleted

            df = pd.DataFrame(data, columns=["Local", "Remote", "Status", "Process"])

            try:
                if export_format == "txt":
                    df.to_csv(file_path, sep="\t", index=False)
                elif export_format == "csv":
                    df.to_csv(file_path, index=False)
                elif export_format == "xlsx":
                    df.to_excel(file_path, index=False)
                elif export_format == "html":
                    df.to_html(file_path, index=False)

                messagebox.showinfo("Exported", "Data exported successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred during export: {str(e)}")

    def format_address(self, address):
        if address:
            ip, port = address
            return f"{ip}:{port}"
        else:
            return "N/A"

    def search_connections(self):
        search_text = self.search_entry.get().strip().lower()
        self.system_processes.clear()
        self.user_processes.clear()

        for conn in psutil.net_connections(kind="inet"):
            local_address = self.format_address(conn.laddr)
            remote_address = "N/A" if conn.raddr is None else self.format_address(conn.raddr)
            status = conn.status
            process_name = self.get_process_name(conn.pid)

            if search_text in process_name.lower():
                if self.is_system_process(process_name):
                    self.system_processes.append((local_address, remote_address, status, process_name))
                else:
                    self.user_processes.append((local_address, remote_address, status, process_name))

        self.sort_connections()

    def on_tree_click(self, event):
        selected_item = self.connections_list.selection()
        if not selected_item:
            return

        self.selected_items.clear()
        for item in selected_item:
            item_id = self.connections_list.item(item, "tags")[0]
            self.selected_items.add(item_id)

    def on_double_click(self, event):
        selected_item = self.connections_list.selection()
        if not selected_item:
            return

        item_id = self.connections_list.item(selected_item[0], "tags")[0]
        values = self.connections_list.item(selected_item[0], "values")
        process_name = values[3]

        self.system_processes.clear()
        self.user_processes.clear()

        for conn in psutil.net_connections(kind="inet"):
            local_address = self.format_address(conn.laddr)
            remote_address = "N/A" if conn.raddr is None else self.format_address(conn.raddr)
            status = conn.status
            conn_process_name = self.get_process_name(conn.pid)

            if process_name == conn_process_name:
                if self.is_system_process(conn_process_name):
                    self.system_processes.append((local_address, remote_address, status, conn_process_name))
                else:
                    self.user_processes.append((local_address, remote_address, status, conn_process_name))

        self.sort_connections()

if __name__ == "__main__":
    root = tk.Tk()
    app = NetworkConnectionsApp(root)
    root.mainloop()
