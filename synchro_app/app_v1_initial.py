import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import customtkinter as ctk
import pyvisa
import openpyxl
import threading
import time

# Set Appearance
ctk.set_appearance_mode("System")  # Modes: "System", "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue", "green", "dark-blue"

class SynchroApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window Config
        self.title("Synchro Voltage Measurement System - Agilent 34401A")
        self.geometry("950x650")
        self.resizable(True, True)

        # Variables
        self.visa_rm = None
        self.multimeter = None
        self.excel_file_path = ""
        self.sheet_name = "Sheet1"
        self.current_tp = 1
        self.total_tp = 11
        self.start_row = 12  # Default Excel start row for TP 1

        # Columns in Excel: P=16 (X), Q=17 (Y), R=18 (Z)
        self.col_x = 16
        self.col_y = 17
        self.col_z = 18

        # Data structure for Table
        self.data_records = {i: {"X": "-", "Y": "-", "Z": "-"} for i in range(1, self.total_tp + 1)}

        # Layout Setup
        self.create_widgets()
        self.refresh_com_ports()

    def create_widgets(self):
        # Grid Configuration
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=2)
        self.grid_rowconfigure(0, weight=1)

        # ==========================================
        # LEFT PANEL: CONTROLS & CONNECTIONS
        # ==========================================
        left_frame = ctk.CTkFrame(self, corner_radius=10)
        left_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # Title
        title_label = ctk.CTkLabel(left_frame, text="Control Panel", font=ctk.CTkFont(size=18, weight="bold"))
        title_label.pack(padx=10, pady=(15, 10))

        # 1. Connection Section
        conn_group = ctk.CTkFrame(left_frame)
        conn_group.pack(padx=10, pady=10, fill="x")

        ctk.CTkLabel(conn_group, text="Multimeter Connection", font=ctk.CTkFont(size=14, weight="bold")).pack(padx=10, pady=(10, 5), anchor="w")
        
        com_layout = ctk.CTkFrame(conn_group, fg_color="transparent")
        com_layout.pack(padx=10, pady=5, fill="x")

        self.com_combo = ctk.CTkComboBox(com_layout, values=["Searching..."])
        self.com_combo.pack(side="left", fill="x", expand=True, padx=(0, 5))

        btn_refresh = ctk.CTkButton(com_layout, text="🔄", width=35, command=self.refresh_com_ports)
        btn_refresh.pack(side="right")

        self.btn_connect = ctk.CTkButton(conn_group, text="Connect Multimeter", fg_color="green", hover_color="darkgreen", command=self.toggle_connection)
        self.btn_connect.pack(padx=10, pady=10, fill="x")

        # 2. Excel Setup Section
        excel_group = ctk.CTkFrame(left_frame)
        excel_group.pack(padx=10, pady=10, fill="x")

        ctk.CTkLabel(excel_group, text="Excel Target File", font=ctk.CTkFont(size=14, weight="bold")).pack(padx=10, pady=(10, 5), anchor="w")

        self.lbl_excel_file = ctk.CTkLabel(excel_group, text="No File Selected", text_color="gray", wraplength=220)
        self.lbl_excel_file.pack(padx=10, pady=5)

        btn_browse = ctk.CTkButton(excel_group, text="Browse Excel File", command=self.browse_excel)
        btn_browse.pack(padx=10, pady=(5, 10), fill="x")

        # 3. Test Point Selector
        tp_group = ctk.CTkFrame(left_frame)
        tp_group.pack(padx=10, pady=10, fill="x")

        ctk.CTkLabel(tp_group, text="Active Test Point", font=ctk.CTkFont(size=14, weight="bold")).pack(padx=10, pady=(10, 5), anchor="w")

        self.tp_spinbox_frame = ctk.CTkFrame(tp_group, fg_color="transparent")
        self.tp_spinbox_frame.pack(padx=10, pady=5, fill="x")

        self.lbl_tp_val = ctk.CTkLabel(self.tp_spinbox_frame, text=f"Test Point: {self.current_tp} / {self.total_tp}", font=ctk.CTkFont(size=16, weight="bold"))
        self.lbl_tp_val.pack(pady=5)

        # 4. Measurement Actions
        measure_group = ctk.CTkFrame(left_frame)
        measure_group.pack(padx=10, pady=10, fill="x")

        ctk.CTkLabel(measure_group, text="Trigger Measurement", font=ctk.CTkFont(size=14, weight="bold")).pack(padx=10, pady=(10, 5), anchor="w")

        self.btn_meas_x = ctk.CTkButton(measure_group, text="Measure X to GND", command=lambda: self.measure_channel('X'))
        self.btn_meas_x.pack(padx=10, pady=5, fill="x")

        self.btn_meas_y = ctk.CTkButton(measure_group, text="Measure Y to GND", command=lambda: self.measure_channel('Y'))
        self.btn_meas_y.pack(padx=10, pady=5, fill="x")

        self.btn_meas_z = ctk.CTkButton(measure_group, text="Measure Z to GND", command=lambda: self.measure_channel('Z'))
        self.btn_meas_z.pack(padx=10, pady=5, fill="x")

        # ==========================================
        # RIGHT PANEL: DATA TABLE & STATUS LOG
        # ==========================================
        right_frame = ctk.CTkFrame(self, corner_radius=10)
        right_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        right_frame.grid_rowconfigure(1, weight=1)
        right_frame.grid_columnconfigure(0, weight=1)

        # Table Header
        ctk.CTkLabel(right_frame, text="Measurement Results Dashboard", font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, padx=15, pady=(15, 10), sticky="w")

        # Treeview Table (Tkinter Standard embedded in CTk)
        table_frame = ctk.CTkFrame(right_frame)
        table_frame.grid(row=1, column=0, padx=15, pady=5, sticky="nsew")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", font=('Calibri', 11), rowheight=25)
        style.configure("Treeview.Heading", font=('Calibri', 12, 'bold'))

        self.tree = ttk.Treeview(table_frame, columns=("TP", "X_GND", "Y_GND", "Z_GND"), show="headings")
        self.tree.heading("TP", text="Test Point")
        self.tree.heading("X_GND", text="X to GND (VAC)")
        self.tree.heading("Y_GND", text="Y to GND (VAC)")
        self.tree.heading("Z_GND", text="Z to GND (VAC)")

        self.tree.column("TP", anchor="center", width=100)
        self.tree.column("X_GND", anchor="center", width=150)
        self.tree.column("Y_GND", anchor="center", width=150)
        self.tree.column("Z_GND", anchor="center", width=150)

        self.tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.grid(row=0, column=1, sticky="ns")

        # Populate initial empty rows in treeview
        self.update_table_view()

        # Navigation Controls for Test Points
        nav_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        nav_frame.grid(row=2, column=0, padx=15, pady=5, sticky="ew")

        btn_prev_tp = ctk.CTkButton(nav_frame, text="◀ Previous TP", width=120, command=self.prev_tp)
        btn_prev_tp.pack(side="left", padx=5)

        btn_next_tp = ctk.CTkButton(nav_frame, text="Next TP ▶", width=120, command=self.next_tp)
        btn_next_tp.pack(side="right", padx=5)

        # Log Window
        ctk.CTkLabel(right_frame, text="System Log", font=ctk.CTkFont(size=12, weight="bold")).grid(row=3, column=0, padx=15, pady=(10, 0), sticky="w")
        
        self.txt_log = ctk.CTkTextbox(right_frame, height=100)
        self.txt_log.grid(row=4, column=0, padx=15, pady=(5, 15), sticky="ew")
        self.log_message("System Started. Please connect multimeter and select Excel file.")

    # ==========================================
    # LOGIC & FUNCTIONS
    # ==========================================
    def log_message(self, msg):
        timestamp = time.strftime("[%H:%M:%S] ")
        self.txt_log.insert("end", timestamp + msg + "\n")
        self.txt_log.see("end")

    def refresh_com_ports(self):
        try:
            self.visa_rm = pyvisa.ResourceManager()
            ports = self.visa_rm.list_resources()
            com_list = [p for p in ports if "ASRL" in p or "GPIB" in p]
            if com_list:
                self.com_combo.configure(values=com_list)
                self.com_combo.set(com_list[0])
                self.log_message(f"Detected ports: {', '.join(com_list)}")
            else:
                self.com_combo.configure(values=["No COM Found"])
                self.com_combo.set("No COM Found")
                self.log_message("No VISA instrument ports found.")
        except Exception as e:
            self.log_message(f"Error listing ports: {e}")

    def toggle_connection(self):
        if self.multimeter is None:
            # Connect
            selected_port = self.com_combo.get()
            if "No COM" in selected_port or not selected_port:
                messagebox.showwarning("Warning", "Please select a valid COM port.")
                return

            try:
                self.multimeter = self.visa_rm.open_resource(selected_port)
                self.multimeter.baud_rate = 9600
                self.multimeter.data_bits = 8
                self.multimeter.parity = pyvisa.constants.Parity.none
                self.multimeter.stop_bits = pyvisa.constants.StopBits.two
                self.multimeter.timeout = 5000

                self.multimeter.write("*RST")
                time.sleep(0.3)
                self.multimeter.write("CONF:VOLT:AC")

                idn = self.multimeter.query("*IDN?").strip()
                self.log_message(f"Connected to: {idn}")
                self.btn_connect.configure(text="Disconnect Multimeter", fg_color="red", hover_color="darkred")
            except Exception as e:
                self.multimeter = None
                self.log_message(f"Connection Failed: {e}")
                messagebox.showerror("Error", f"Could not connect to instrument:\n{e}")
        else:
            # Disconnect
            try:
                self.multimeter.close()
            except:
                pass
            self.multimeter = None
            self.log_message("Multimeter Disconnected.")
            self.btn_connect.configure(text="Connect Multimeter", fg_color="green", hover_color="darkgreen")

    def browse_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            self.excel_file_path = file_path
            self.lbl_excel_file.configure(text=file_path.split("/")[-1], text_color="white")
            self.log_message(f"Excel file linked: {file_path}")

    def measure_channel(self, channel):
        if self.multimeter is None:
            messagebox.showwarning("Warning", "Please connect to the multimeter first!")
            return

        def run_measurement():
            try:
                self.log_message(f"Measuring {channel} to GND for Test Point {self.current_tp}...")
                raw_val = self.multimeter.query("READ?")
                val = round(float(raw_val), 2)

                # Update internal data
                self.data_records[self.current_tp][channel] = val
                self.update_table_view()
                self.log_message(f"Result {channel}-GND: {val} VAC")

                # Auto Save to Excel
                if self.excel_file_path:
                    self.save_single_to_excel(self.current_tp, channel, val)
                else:
                    self.log_message("[Info] Data updated in app table (Excel not linked).")

            except Exception as e:
                self.log_message(f"Measurement Error: {e}")
                messagebox.showerror("Error", f"Failed to read data from Multimeter:\n{e}")

        # Threading to prevent GUI freeze
        threading.Thread(target=run_measurement, daemon=True).start()

    def save_single_to_excel(self, tp, channel, value):
        try:
            wb = openpyxl.load_workbook(self.excel_file_path)
            ws = wb[self.sheet_name]

            target_row = self.start_row + (tp - 1)
            target_col = self.col_x if channel == 'X' else (self.col_y if channel == 'Y' else self.col_z)

            ws.cell(row=target_row, column=target_col, value=value)
            wb.save(self.excel_file_path)
            self.log_message(f"Saved {value} V to Excel Row {target_row}, Col {target_col}.")
        except Exception as e:
            self.log_message(f"Excel Write Error: {e}")
            messagebox.showerror("Excel Error", f"Could not write to Excel file. Make sure it is closed!\nError: {e}")

    def update_table_view(self):
        # Clear Table
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Repopulate
        for tp in range(1, self.total_tp + 1):
            row = self.data_records[tp]
            self.tree.insert("", "end", values=(f"TP {tp}", row["X"], row["Y"], row["Z"]))

    def next_tp(self):
        if self.current_tp < self.total_tp:
            self.current_tp += 1
            self.lbl_tp_val.configure(text=f"Test Point: {self.current_tp} / {self.total_tp}")

    def prev_tp(self):
        if self.current_tp > 1:
            self.current_tp -= 1
            self.lbl_tp_val.configure(text=f"Test Point: {self.current_tp} / {self.total_tp}")

if __name__ == "__main__":
    app = SynchroApp()
    app.mainloop()
