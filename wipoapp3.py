import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
import datetime

# WIPO ST.3 Country Code Mapping
COUNTRY_MAP = {
    'JP': {'code': 'JP', 'name': 'Japan'},
    'CN': {'code': 'CN', 'name': 'China'},
    'AU': {'code': 'AU', 'name': 'Australia'},
    'AT': {'code': 'AT', 'name': 'Austria'},
    'DE': {'code': 'DE', 'name': 'Germany'},
    'FR': {'code': 'FR', 'name': 'France'},
    'KR': {'code': 'KR', 'name': 'Republic of Korea'},
    'CA': {'code': 'CA', 'name': 'Canada'},
    'UK': {'code': 'GB', 'name': 'United Kingdom (GB)'},
    'IT': {'code': 'IT', 'name': 'Italy'}
}

# Pre-populated sample data including both Applicants and Representing Law Firms / Agents
MOCK_DATA = [
    # Japan (JP)
    {'pub_num': 'WO/2026/012345', 'country': 'JP', 'applicant': 'Toyota Motor Corp', 'law_firm': 'Aoyama & Partners', 'date': '2026-10-15', 'title': 'Solid-State Battery Thermal Management System'},
    {'pub_num': 'WO/2026/012360', 'country': 'JP', 'applicant': 'Toyota Motor Corp', 'law_firm': 'Aoyama & Partners', 'date': '2026-10-15', 'title': 'Fuel Cell Vehicle Pressure Control Valve'},
    {'pub_num': 'WO/2026/012346', 'country': 'JP', 'applicant': 'Sony Group Corp', 'law_firm': 'Seiwa Patent & Law', 'date': '2026-10-15', 'title': 'Image Sensor Signal Processing Apparatus'},
    {'pub_num': 'WO/2026/012361', 'country': 'JP', 'applicant': 'Sony Group Corp', 'law_firm': 'Seiwa Patent & Law', 'date': '2026-10-15', 'title': 'Haptic Feedback Actuator for Controllers'},
    {'pub_num': 'WO/2026/012362', 'country': 'JP', 'applicant': 'Panasonic Holdings Corp', 'law_firm': 'Aoyama & Partners', 'date': '2026-10-15', 'title': 'Lithium-Ion Cathode Material Coating'},

    # China (CN)
    {'pub_num': 'WO/2026/012347', 'country': 'CN', 'applicant': 'Tech Innovations Ltd', 'law_firm': 'China Patent Agent (H.K.) Ltd.', 'date': '2026-10-15', 'title': 'Wireless Communications Scheduling Algorithm'},
    {'pub_num': 'WO/2026/012363', 'country': 'CN', 'applicant': 'Tech Innovations Ltd', 'law_firm': 'China Patent Agent (H.K.) Ltd.', 'date': '2026-10-15', 'title': 'MIMO Beamforming Antenna Array'},
    {'pub_num': 'WO/2026/012348', 'country': 'CN', 'applicant': 'Automotive Tech Co', 'law_firm': 'King & Wood Mallesons', 'date': '2026-10-15', 'title': 'Electric Vehicle Drive Inverter Assembly'},
    {'pub_num': 'WO/2026/012364', 'country': 'CN', 'applicant': 'BYD Co Ltd', 'law_firm': 'CCPIT Patent & Trademark Law Office', 'date': '2026-10-15', 'title': 'Blade Battery Thermal Disconnect Mechanism'},

    # Australia (AU)
    {'pub_num': 'WO/2026/012349', 'country': 'AU', 'applicant': 'CSIRO Innovations', 'law_firm': 'Davies Collison Cave', 'date': '2026-10-15', 'title': 'Wireless Agriculture Soil Moisture Monitoring Node'},
    {'pub_num': 'WO/2026/012365', 'country': 'AU', 'applicant': 'CSIRO Innovations', 'law_firm': 'FB Rice Patent Attorneys', 'date': '2026-10-15', 'title': 'Solar Cell Passivation Layer Synthesis'},

    # Austria (AT)
    {'pub_num': 'WO/2026/012350', 'country': 'AT', 'applicant': 'Steelmaking AG', 'law_firm': 'Sonn & Partner Patentanwälte', 'date': '2026-10-15', 'title': 'Zero-Carbon Hydrogen Steelmaking Reduction Process'},

    # Germany (DE)
    {'pub_num': 'WO/2026/012351', 'country': 'DE', 'applicant': 'Siemens AG', 'law_firm': 'Siemens IP Department & Law', 'date': '2026-10-15', 'title': 'Industrial Microgrid Edge Controller'},
    {'pub_num': 'WO/2026/012366', 'country': 'DE', 'applicant': 'Siemens AG', 'law_firm': 'Grünecker Patent- und Rechtsanwälte', 'date': '2026-10-15', 'title': 'Predictive Maintenance Sensor Network'},
    {'pub_num': 'WO/2026/012352', 'country': 'DE', 'applicant': 'BASF SE', 'law_firm': 'Grünecker Patent- und Rechtsanwälte', 'date': '2026-10-15', 'title': 'Bio-Based Polymer Catalyst Composition'},

    # France (FR)
    {'pub_num': 'WO/2026/012353', 'country': 'FR', 'applicant': 'TotalEnergies SE', 'law_firm': 'Cabinet Lavoix', 'date': '2026-10-15', 'title': 'Offshore Floating Wind Turbine Mooring Control'},

    # Republic of Korea (KR)
    {'pub_num': 'WO/2026/012354', 'country': 'KR', 'applicant': 'Samsung Electronics Co Ltd', 'law_firm': 'Kim & Chang', 'date': '2026-10-15', 'title': 'High-Bandwidth Memory Architecture'},
    {'pub_num': 'WO/2026/012367', 'country': 'KR', 'applicant': 'Samsung Electronics Co Ltd', 'date': '2026-10-15', 'law_firm': 'First Law Offices', 'title': 'EUV Lithography Substrate Patterning'},

    # Canada (CA)
    {'pub_num': 'WO/2026/012355', 'country': 'CA', 'applicant': 'Ballard Power Systems Inc', 'law_firm': 'Bereskin & Parr LLP', 'date': '2026-10-15', 'title': 'Proton Exchange Membrane Fuel Cell Stack'},

    # United Kingdom (UK -> GB)
    {'pub_num': 'WO/2026/012356', 'country': 'UK', 'applicant': 'Rolls-Royce PLC', 'law_firm': 'Mewburn Ellis LLP', 'date': '2026-10-15', 'title': 'Hybrid Electric Aircraft Propulsion Drive'},

    # Italy (IT)
    {'pub_num': 'WO/2026/012357', 'country': 'IT', 'applicant': 'Enel S.p.A.', 'law_firm': 'Studio Torta S.p.A.', 'date': '2026-10-15', 'title': 'Smart Power Grid Load Balancing System'},
]

class WipoPctApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WIPO PCT Weekly Publication Search (Applicant & Law Firm Sorting)")
        self.root.geometry("1080x700")
        self.root.minsize(880, 540)

        self.style = ttk.Style()
        if "clam" in self.style.theme_names():
            self.style.theme_use("clam")

        self.current_results = []
        self.sort_column = "applicant"  # Default sort column: applicant
        self.sort_reverse = False

        self._build_ui()

    def _build_ui(self):
        # Header
        header_frame = ttk.Frame(self.root, padding=15)
        header_frame.pack(fill=tk.X)

        title_lbl = ttk.Label(
            header_frame,
            text="WIPO PCT Weekly Publications Search",
            font=("Helvetica", 16, "bold")
        )
        title_lbl.pack(anchor=tk.W)

        subtitle_lbl = ttk.Label(
            header_frame,
            text="Filter international PCT applications by country of origin (WIPO ST.3 standard codes) - Sort by Applicant or Law Firm",
            font=("Helvetica", 10)
        )
        subtitle_lbl.pack(anchor=tk.W, pady=(2, 0))

        # Filter Section
        filter_frame = ttk.LabelFrame(self.root, text=" Search Filters & Sorting Options ", padding=12)
        filter_frame.pack(fill=tk.X, padx=15, pady=5)

        ttk.Label(filter_frame, text="Country of Origin:", font=("Helvetica", 10, "bold")).grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)

        country_options = [f"{key} - {info['name']}" for key, info in COUNTRY_MAP.items()]
        self.country_combo = ttk.Combobox(filter_frame, values=country_options, state="readonly", width=28)
        self.country_combo.set(country_options[0])
        self.country_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

        ttk.Label(filter_frame, text="Publication Date (YYYY-MM-DD):", font=("Helvetica", 10)).grid(row=0, column=2, sticky=tk.W, padx=(20, 5), pady=5)
        
        self.date_entry = ttk.Entry(filter_frame, width=15)
        self.date_entry.insert(0, "2026-10-15")
        self.date_entry.grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)

        # Quick Sort Control Buttons
        ttk.Label(filter_frame, text="Quick Sort By:", font=("Helvetica", 10, "bold")).grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)

        sort_btn_frame = ttk.Frame(filter_frame)
        sort_btn_frame.grid(row=1, column=1, columnspan=3, sticky=tk.W, padx=5, pady=5)

        self.btn_sort_applicant = ttk.Button(sort_btn_frame, text="🏢 Sort by Applicant", command=lambda: self.set_sort_key("applicant"))
        self.btn_sort_applicant.pack(side=tk.LEFT, padx=(0, 5))

        self.btn_sort_lawfirm = ttk.Button(sort_btn_frame, text="⚖️ Sort by Law Firm", command=lambda: self.set_sort_key("law_firm"))
        self.btn_sort_lawfirm.pack(side=tk.LEFT, padx=5)

        ttk.Label(sort_btn_frame, text=" (Click header columns or use buttons above)").pack(side=tk.LEFT, padx=10)

        # Action Buttons
        btn_frame = ttk.Frame(filter_frame)
        btn_frame.grid(row=2, column=0, columnspan=4, sticky=tk.W, pady=(10, 0))

        search_btn = ttk.Button(btn_frame, text="Search Publications", command=self.perform_search)
        search_btn.pack(side=tk.LEFT, padx=5)

        clear_btn = ttk.Button(btn_frame, text="Reset Filters", command=self.reset_filters)
        clear_btn.pack(side=tk.LEFT, padx=5)

        copy_query_btn = ttk.Button(btn_frame, text="Generate WIPO Query String", command=self.generate_wipo_query)
        copy_query_btn.pack(side=tk.LEFT, padx=5)

        # Results Table Section
        results_frame = ttk.LabelFrame(self.root, text=" Results List ", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)

        top_bar = ttk.Frame(results_frame)
        top_bar.pack(fill=tk.X, pady=(0, 5))

        self.count_label = ttk.Label(top_bar, text="Found 0 publication(s)", font=("Helvetica", 10, "bold"))
        self.count_label.pack(side=tk.LEFT)

        export_btn = ttk.Button(top_bar, text="Export Sorted Results...", command=self.export_results)
        export_btn.pack(side=tk.RIGHT)

        tree_scroll_y = ttk.Scrollbar(results_frame, orient=tk.VERTICAL)
        tree_scroll_x = ttk.Scrollbar(results_frame, orient=tk.HORIZONTAL)

        columns = ("pub_num", "country", "st3_code", "applicant", "law_firm", "date", "title")
        self.tree = ttk.Treeview(
            results_frame,
            columns=columns,
            show="headings",
            yscrollcommand=tree_scroll_y.set,
            xscrollcommand=tree_scroll_x.set,
            selectmode="extended"
        )

        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        tree_scroll_x.config(command=self.tree.xview)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree.pack(fill=tk.BOTH, expand=True)

        self.headers_def = {
            "applicant": ("Applicant Name ▲", 210),
            "law_firm": ("Law Firm / Representative", 220),
            "pub_num": ("Publication Number", 150),
            "country": ("Selected Country", 100),
            "st3_code": ("WIPO ST.3 Code", 100),
            "date": ("Pub Date", 90),
            "title": ("Application Title", 280)
        }

        for col, (heading, width) in self.headers_def.items():
            self.tree.heading(col, text=heading, command=lambda c=col: self.sort_by_column(c))
            self.tree.column(col, width=width, anchor=tk.W)

        self.status_var = tk.StringVar(value="Ready. Select a country and click Search.")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding=5)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

        self.perform_search()

    def get_selected_country(self):
        val = self.country_combo.get()
        user_key = val.split(" - ")[0]
        return user_key, COUNTRY_MAP.get(user_key, {"code": user_key, "name": user_key})

    def set_sort_key(self, col):
        if self.sort_column == col:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = col
            self.sort_reverse = False
        self.refresh_table_display()

    def sort_by_column(self, col):
        self.set_sort_key(col)

    def perform_search(self):
        user_key, country_info = self.get_selected_country()

        filtered = [
            item for item in MOCK_DATA
            if item["country"] == user_key
        ]

        self.current_results = filtered
        self.refresh_table_display()

    def refresh_table_display(self):
        user_key, country_info = self.get_selected_country()
        st3_code = country_info["code"]

        # Sort current results by active sort_column
        self.current_results.sort(
            key=lambda x: (x.get(self.sort_column, "").lower(), x.get("pub_num", "")),
            reverse=self.sort_reverse
        )

        for row in self.tree.get_children():
            self.tree.delete(row)

        headers_base = {
            "applicant": "Applicant Name",
            "law_firm": "Law Firm / Representative",
            "pub_num": "Publication Number",
            "country": "Selected Country",
            "st3_code": "WIPO ST.3 Code",
            "date": "Pub Date",
            "title": "Application Title"
        }

        arrow = " ▼" if self.sort_reverse else " ▲"
        for col, title in headers_base.items():
            header_text = title + (arrow if col == self.sort_column else "")
            self.tree.heading(col, text=header_text)

        for item in self.current_results:
            self.tree.insert(
                "",
                tk.END,
                values=(
                    item["pub_num"],
                    user_key,
                    st3_code,
                    item["applicant"],
                    item["law_firm"],
                    item["date"],
                    item["title"]
                )
            )

        unique_applicants = len(set(i["applicant"] for i in self.current_results))
        unique_law_firms = len(set(i["law_firm"] for i in self.current_results))
        active_sort_name = headers_base.get(self.sort_column, self.sort_column)

        self.count_label.config(
            text=f"Found {len(self.current_results)} publication(s) | {unique_applicants} Applicant(s) | {unique_law_firms} Law Firm(s) [{user_key} / {st3_code}]"
        )
        self.status_var.set(
            f"Currently sorted by: {active_sort_name} ({'Descending' if self.sort_reverse else 'Ascending'})."
        )

    def reset_filters(self):
        country_options = [f"{key} - {info['name']}" for key, info in COUNTRY_MAP.items()]
        self.country_combo.set(country_options[0])
        self.sort_column = "applicant"
        self.sort_reverse = False
        self.date_entry.delete(0, tk.END)
        self.date_entry.insert(0, "2026-10-15")
        self.perform_search()

    def generate_wipo_query(self):
        user_key, country_info = self.get_selected_country()
        st3_code = country_info["code"]
        pub_date = self.date_entry.get().strip()

        date_query = pub_date.replace("-", "")
        query = f"PA:{st3_code} AND DP:{date_query}"

        top = tk.Toplevel(self.root)
        top.title("WIPO PATENTSCOPE Query Syntax")
        top.geometry("580x160")
        top.transient(self.root)

        ttk.Label(top, text="WIPO PATENTSCOPE Query Syntax:", font=("Helvetica", 10, "bold")).pack(anchor=tk.W, padx=15, pady=(15, 5))
        
        q_entry = ttk.Entry(top, width=65)
        q_entry.insert(0, query)
        q_entry.pack(padx=15, pady=5)

        def copy_url():
            self.root.clipboard_clear()
            self.root.clipboard_append(query)
            messagebox.showinfo("Copied", "Query string copied to clipboard!", parent=top)

        close_btn = ttk.Button(top, text="Copy Query & Close", command=copy_url)
        close_btn.pack(pady=10)

    def export_results(self):
        if not self.current_results:
            messagebox.showwarning("Export Warning", "No results available to export. Perform a search first.")
            return

        user_key, country_info = self.get_selected_country()
        sort_label = "applicant" if self.sort_column == "applicant" else ("law_firm" if self.sort_column == "law_firm" else self.sort_column)

        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[
                ("CSV Files (*.csv)", "*.csv"),
                ("Text Files (*.txt)", "*.txt"),
                ("All Files (*.*)", "*.*")
            ],
            title=f"Export WIPO Search Results (Sorted by {sort_label.title()})",
            initialfile=f"wipo_pct_publications_{user_key}_sorted_by_{sort_label}.csv"
        )

        if not file_path:
            return

        try:
            if file_path.endswith(".txt"):
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(f"WIPO PCT Weekly Publications Export\n")
                    f.write(f"Country Selection: {user_key} | WIPO ST.3 Code: {country_info['code']}\n")
                    f.write(f"Active Sort Metric: {sort_label.upper()}\n")
                    f.write(f"Export Date: {datetime.date.today()}\n")
                    f.write("=" * 80 + "\n\n")

                    current_group = None
                    group_field = self.sort_column if self.sort_column in ("applicant", "law_firm") else "applicant"

                    for item in self.current_results:
                        val = item.get(group_field, "N/A")
                        if val != current_group:
                            current_group = val
                            header_title = "LAW FIRM / REPRESENTATIVE" if group_field == "law_firm" else "APPLICANT"
                            f.write(f"\n--- {header_title}: {current_group} ---\n")
                        
                        f.write(f"  Publication Number: {item['pub_num']}\n")
                        f.write(f"  Applicant Name:     {item['applicant']}\n")
                        f.write(f"  Law Firm / Agent:   {item['law_firm']}\n")
                        f.write(f"  Publication Date:   {item['date']}\n")
                        f.write(f"  Title:              {item['title']}\n\n")
            else:
                with open(file_path, "w", newline="", encoding="utf-8") as f:
                    writer = csv.writer(f)
                    writer.writerow(["Applicant Name", "Law Firm / Agent", "Publication Number", "Country Choice", "ST.3 Code", "Pub Date", "Title"])
                    for item in self.current_results:
                        writer.writerow([
                            item["applicant"],
                            item["law_firm"],
                            item["pub_num"],
                            user_key,
                            country_info["code"],
                            item["date"],
                            item["title"]
                        ])

            messagebox.showinfo("Export Success", f"Successfully exported {len(self.current_results)} publication record(s) sorted by {sort_label} to:\n{file_path}")
            self.status_var.set(f"Exported {len(self.current_results)} rows to {file_path}")

        except Exception as e:
            messagebox.showerror("Export Error", f"An error occurred while saving file:\n{str(e)}")

def main():
    root = tk.Tk()
    app = WipoPctApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
