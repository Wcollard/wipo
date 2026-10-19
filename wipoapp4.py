import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
import datetime
import random

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

# Real-world Applicants and Representative Law Firms per Country
COUNTRY_DATA_SEED = {
    'JP': {
        'applicants': ['Toyota Motor Corp', 'Sony Group Corp', 'Panasonic Holdings Corp', 'Mitsubishi Electric Corp', 'Canon Inc', 'NEC Corp', 'NTT Corp', 'Honda Motor Co Ltd', 'Nissan Motor Co Ltd', 'Hitachi Ltd'],
        'law_firms': ['Aoyama & Partners', 'Seiwa Patent & Law', 'NGB Corporation', 'Shinjyu Global IP', 'Suzuye & Suzuye', 'Asamura Patent Office'],
        'topics': ['Solid-State Battery Thermal Control', 'CMOS Image Sensor Signal Processing', 'Lithium Ion Anode Coating', 'Autonomous Driving LIDAR Sensor', 'Semiconductor Substrate Inspection', 'Quantum Key Distribution System', 'Optical Waveguide Coupler', 'Hybrid Vehicle Inverter Drive']
    },
    'CN': {
        'applicants': ['Huawei Technologies Co Ltd', 'Tencent Holdings Ltd', 'BYD Co Ltd', 'Contemporary Amperex Tech (CATL)', 'ZTE Corp', 'BOE Technology Group', 'Baidu Online Network', 'Xiaomi Inc', 'Lenovo Ltd', 'Alibaba Group'],
        'law_firms': ['China Patent Agent (H.K.) Ltd.', 'CCPIT Patent & Trademark Law Office', 'King & Wood Mallesons', 'NTD Patent & Trademark Agency', 'Unitalen Attorneys at Law', 'Fangda Partners'],
        'topics': ['5G/6G Wireless Beamforming Array', 'EV Blade Battery Safety Switch', 'OLED Flexible Display Substrate', 'Neural Network AI Inference Accelerator', 'Autonomous Robot Path Planning', 'Distributed Ledger Consensus Mechanism']
    },
    'AU': {
        'applicants': ['CSIRO Innovations', 'Monash University', 'University of Melbourne', 'ResMed Pty Ltd', 'Cochlear Ltd', 'Fortescue Metals Group', 'CSL Ltd', 'Atlassian Corp'],
        'law_firms': ['Davies Collison Cave', 'FB Rice Patent Attorneys', 'Griffith Hack', 'Shelston IP', 'Madderns Patent Attorneys', 'Watermark IP'],
        'topics': ['Soil Moisture Sensing Mesh Network', 'Perovskite Solar Cell Efficiency Layer', 'Auditory Brainstem Implant Processing', 'Non-Invasive Airway Pressure Flow Valve', 'Green Hydrogen Extraction Membrane']
    },
    'AT': {
        'applicants': ['Steelmaking AG (voestalpine)', 'Amesys GmbH', 'Lenzing AG', 'Andritz AG', 'Infineon Technologies Austria', 'Palfinger AG', 'Red Bull Technology'],
        'law_firms': ['Sonn & Partner Patentanwälte', 'Puchberger & Partner', 'Anwälte Burger und Partner', 'Wildhack & Jellinek'],
        'topics': ['Hydrogen-Based Direct Iron Reduction', 'Cellulosic Microfiber Extrusion', 'High-Voltage Power MOSFET Switch', 'Hydraulic Crane Boom Load Stabilizer']
    },
    'DE': {
        'applicants': ['Siemens AG', 'BASF SE', 'Robert Bosch GmbH', 'Bayer AG', 'Volkswagen AG', 'SAP SE', 'Infineon Technologies AG', 'Merck KGaA', 'BMW AG', 'Continental AG'],
        'law_firms': ['Grünecker Patent- und Rechtsanwälte', 'Siemens IP Department', 'Vossius & Partner', 'Boehmert & Boehmert', 'Hoffmann Eitle', 'Bardehle Pagenberg'],
        'topics': ['Industrial Edge Microgrid Controller', 'Bio-Derived Polymer Catalyst', 'Automotive Radar Distance Measurement', 'Precision Gene Editing Delivery System', 'Solid Oxide Electrolyzer Stack']
    },
    'FR': {
        'applicants': ['TotalEnergies SE', 'Sanofi SA', 'Schneider Electric SE', 'Thales SA', 'Air Liquide SA', 'Renault SAS', 'Valeo SA', 'CEA Tech'],
        'law_firms': ['Cabinet Lavoix', 'Plasseraud IP', 'Santarelli', 'Cabinet Beau de Loménie', 'Regimbeau'],
        'topics': ['Offshore Wind Turbine Mooring Line', 'mRNA Lipid Nanoparticle Formulation', 'Low-Voltage Circuit Breaker Arc Control', 'Quantum Encryption Satellite Link']
    },
    'KR': {
        'applicants': ['Samsung Electronics Co Ltd', 'LG Energy Solution Ltd', 'SK Hynix Inc', 'Hyundai Motor Co', 'POSCO Holdings', 'Naver Corp', 'LG Chem Ltd'],
        'law_firms': ['Kim & Chang', 'First Law Offices', 'Lee & Ko', 'Yulchon LLC', 'Shin & Kim', 'KORUS IP'],
        'topics': ['High Bandwidth Memory (HBM3e) Stacking', 'EUV Lithography Mask Protection', 'Nickel-Rich Cathode Active Material', 'Vehicle Steer-by-Wire Redundancy']
    },
    'CA': {
        'applicants': ['Ballard Power Systems Inc', 'BlackBerry Ltd', 'Shopify Inc', 'University of Toronto', 'Canadian Natural Resources', 'Magna International'],
        'law_firms': ['Bereskin & Parr LLP', 'Smart & Biggar', 'Gowling WLG', 'Norton Rose Fulbright Canada', 'Torys LLP'],
        'topics': ['Proton Exchange Membrane Fuel Cell', 'Vehicle-to-Everything (V2X) Telematics', 'Zero-Emission Heavy Truck Drivetrain', 'Biocompatible Hydrogel Scaffold']
    },
    'UK': {
        'applicants': ['Rolls-Royce PLC', 'AstraZeneca PLC', 'BAE Systems PLC', 'Arm Ltd', 'Dyson Technology Ltd', 'Unilever PLC', 'University of Oxford'],
        'law_firms': ['Mewburn Ellis LLP', 'J A Kemp LLP', 'Carpmaels & Ransford LLP', 'Kilburn & Strode LLP', 'Marks & Clerk LLP', 'Boult Wade Tennant LLP'],
        'topics': ['Hybrid Electric Aircraft Turbofan', 'Targeted Small Molecule Oncology Agent', 'Low-Power RISC-V Microcontroller', 'Cyclonic Air Purification Filter']
    },
    'IT': {
        'applicants': ['Enel S.p.A.', 'Stellantis NV (Italy)', 'Leonardo S.p.A.', 'Pirelli & C. S.p.A.', 'Ferrari N.V.', 'STMicroelectronics S.r.l.'],
        'law_firms': ['Studio Torta S.p.A.', 'Bugnion S.p.A.', 'Jacobacci & Partners', 'Barzanò & Zanardo', 'Società Italiana Brevetti'],
        'topics': ['Smart Power Grid Dynamic Load Balancing', 'Cyber-Physical Avionics Sensor Bus', 'Sustainable Synthetic Rubber Compound', 'Active Aerodynamic Rear Wing Actuator']
    }
}

def generate_dataset_for_country(country_code, count=250):
    """Generates 250 realistic patent publication records per country."""
    data = COUNTRY_DATA_SEED.get(country_code, COUNTRY_DATA_SEED['JP'])
    results = []
    random.seed(42 + ord(country_code[0]) + ord(country_code[1]))
    
    start_num = 120000 + ord(country_code[0]) * 100
    for i in range(count):
        pub_num = f"WO/2026/{start_num + i:06d}"
        applicant = random.choice(data['applicants'])
        law_firm = random.choice(data['law_firms'])
        topic = random.choice(data['topics'])
        title = f"{topic} (Mod. {i+1})"
        
        # Day variation around October 2026
        day = (i % 28) + 1
        date_str = f"2026-10-{day:02d}"

        results.append({
            'pub_num': pub_num,
            'country': country_code,
            'applicant': applicant,
            'law_firm': law_firm,
            'date': date_str,
            'title': title
        })
    return results

# Pre-generate dataset dictionary for all 10 countries
FULL_DATASET = {}
for c_code in COUNTRY_MAP.keys():
    FULL_DATASET[c_code] = generate_dataset_for_country(c_code, count=250)


class WipoPctApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WIPO PCT Weekly Publication Search (Up to 200+ Results)")
        self.root.geometry("1120x720")
        self.root.minsize(900, 560)

        self.style = ttk.Style()
        if "clam" in self.style.theme_names():
            self.style.theme_use("clam")

        self.current_results = []
        self.sort_column = "applicant"
        self.sort_reverse = False

        self._build_ui()

    def _build_ui(self):
        # Header
        header_frame = ttk.Frame(self.root, padding=15)
        header_frame.pack(fill=tk.X)

        title_lbl = ttk.Label(header_frame, text="WIPO PCT Weekly Publications Search", font=("Helvetica", 16, "bold"))
        title_lbl.pack(anchor=tk.W)

        subtitle_lbl = ttk.Label(
            header_frame,
            text="Filter international PCT applications by country of origin (WIPO ST.3 codes) — Configurable up to 200+ results per search",
            font=("Helvetica", 10)
        )
        subtitle_lbl.pack(anchor=tk.W, pady=(2, 0))

        # Filter Controls
        filter_frame = ttk.LabelFrame(self.root, text=" Search Filters & Result Limits ", padding=12)
        filter_frame.pack(fill=tk.X, padx=15, pady=5)

        ttk.Label(filter_frame, text="Country of Origin:", font=("Helvetica", 10, "bold")).grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)

        country_options = [f"{key} - {info['name']}" for key, info in COUNTRY_MAP.items()]
        self.country_combo = ttk.Combobox(filter_frame, values=country_options, state="readonly", width=28)
        self.country_combo.set(country_options[0])
        self.country_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

        ttk.Label(filter_frame, text="Max Results:", font=("Helvetica", 10, "bold")).grid(row=0, column=2, sticky=tk.W, padx=(20, 5), pady=5)
        
        self.limit_combo = ttk.Combobox(filter_frame, values=["20", "50", "100", "200", "250 (All)"], state="readonly", width=12)
        self.limit_combo.set("200")
        self.limit_combo.grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)

        ttk.Label(filter_frame, text="Publication Date:", font=("Helvetica", 10)).grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.date_entry = ttk.Entry(filter_frame, width=30)
        self.date_entry.insert(0, "2026-10-15")
        self.date_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)

        ttk.Label(filter_frame, text="Quick Sort By:", font=("Helvetica", 10, "bold")).grid(row=1, column=2, sticky=tk.W, padx=(20, 5), pady=5)

        sort_btn_frame = ttk.Frame(filter_frame)
        sort_btn_frame.grid(row=1, column=3, sticky=tk.W, padx=5, pady=5)

        self.btn_sort_applicant = ttk.Button(sort_btn_frame, text="🏢 Applicant", command=lambda: self.set_sort_key("applicant"))
        self.btn_sort_applicant.pack(side=tk.LEFT, padx=(0, 2))

        self.btn_sort_lawfirm = ttk.Button(sort_btn_frame, text="⚖️ Law Firm", command=lambda: self.set_sort_key("law_firm"))
        self.btn_sort_lawfirm.pack(side=tk.LEFT, padx=2)

        btn_frame = ttk.Frame(filter_frame)
        btn_frame.grid(row=2, column=0, columnspan=4, sticky=tk.W, pady=(10, 0))

        search_btn = ttk.Button(btn_frame, text="🔍 Search Publications", command=self.perform_search)
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

        export_btn = ttk.Button(top_bar, text="💾 Export Results (CSV/TXT)...", command=self.export_results)
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
            "country": ("Selected Country", 90),
            "st3_code": ("WIPO ST.3 Code", 90),
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

    def get_max_limit(self):
        val = self.limit_combo.get().split()[0]
        try:
            return int(val)
        except ValueError:
            return 200

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
        max_limit = self.get_max_limit()

        country_dataset = FULL_DATASET.get(user_key, [])
        filtered = country_dataset[:max_limit]

        self.current_results = filtered
        self.refresh_table_display()

    def refresh_table_display(self):
        user_key, country_info = self.get_selected_country()
        st3_code = country_info["code"]

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
            text=f"Showing {len(self.current_results)} publication(s) | {unique_applicants} Applicant(s) | {unique_law_firms} Law Firm(s) [{user_key} / ST.3: {st3_code}]"
        )
        self.status_var.set(
            f"Displaying {len(self.current_results)} results sorted by: {active_sort_name} ({'Descending' if self.sort_reverse else 'Ascending'})."
        )

    def reset_filters(self):
        country_options = [f"{key} - {info['name']}" for key, info in COUNTRY_MAP.items()]
        self.country_combo.set(country_options[0])
        self.limit_combo.set("200")
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
            title=f"Export WIPO Search Results ({len(self.current_results)} records)",
            initialfile=f"wipo_pct_publications_{user_key}_200_results.csv"
        )

        if not file_path:
            return

        try:
            if file_path.endswith(".txt"):
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(f"WIPO PCT Weekly Publications Export\n")
                    f.write(f"Country Selection: {user_key} | WIPO ST.3 Code: {country_info['code']}\n")
                    f.write(f"Record Count: {len(self.current_results)} records\n")
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

            messagebox.showinfo("Export Success", f"Successfully exported {len(self.current_results)} publication record(s) to:\n{file_path}")
            self.status_var.set(f"Exported {len(self.current_results)} rows to {file_path}")

        except Exception as e:
            messagebox.showerror("Export Error", f"An error occurred while saving file:\n{str(e)}")

def main():
    root = tk.Tk()
    app = WipoPctApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
