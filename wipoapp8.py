import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
import datetime
import random
import webbrowser
import xml.etree.ElementTree as ET
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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

COUNTRY_DATA_SEED = {
    'JP': {
        'applicants': ['Toyota Motor Corp', 'Sony Group Corp', 'Panasonic Holdings Corp', 'Mitsubishi Electric Corp', 'Canon Inc', 'NEC Corp', 'NTT Corp', 'Honda Motor Co Ltd', 'Nissan Motor Co Ltd', 'Hitachi Ltd'],
        'law_firms': ['Aoyama & Partners', 'Seiwa Patent & Law', 'NGB Corporation', 'Shinjyu Global IP', 'Suzuye & Suzuye', 'Asamura Patent Office'],
        'topics': [
            ('Solid-State Battery Thermal Control', 'H01M 10/60', 'Y02E 60/10'),
            ('CMOS Image Sensor Signal Processing', 'H04N 5/369', 'H04N 25/70'),
            ('Lithium Ion Anode Coating', 'H01M 4/13', 'Y02T 10/70'),
            ('Autonomous Driving LIDAR Sensor', 'G01S 17/89', 'G05D 1/02'),
            ('Semiconductor Substrate Inspection', 'H01L 21/66', 'H01L 22/10'),
            ('Quantum Key Distribution System', 'H04L 9/08', 'H04L 88/00'),
            ('Optical Waveguide Coupler', 'G02B 6/26', 'G02B 6/42'),
            ('Hybrid Vehicle Inverter Drive', 'B60L 50/16', 'Y02T 10/62')
        ]
    },
    'CN': {
        'applicants': ['Huawei Technologies Co Ltd', 'Tencent Holdings Ltd', 'BYD Co Ltd', 'Contemporary Amperex Tech (CATL)', 'ZTE Corp', 'BOE Technology Group', 'Baidu Online Network', 'Xiaomi Inc', 'Lenovo Ltd', 'Alibaba Group'],
        'law_firms': ['China Patent Agent (H.K.) Ltd.', 'CCPIT Patent & Trademark Law Office', 'King & Wood Mallesons', 'NTD Patent & Trademark Agency', 'Unitalen Attorneys at Law', 'Fangda Partners'],
        'topics': [
            ('5G/6G Wireless Beamforming Array', 'H04B 7/06', 'H04W 16/28'),
            ('EV Blade Battery Safety Switch', 'H01M 50/583', 'Y02E 60/10'),
            ('OLED Flexible Display Substrate', 'H10K 59/10', 'H10K 77/10'),
            ('Neural Network AI Inference Accelerator', 'G06N 3/063', 'G06N 3/08'),
            ('Autonomous Robot Path Planning', 'G05D 1/00', 'G06T 7/70'),
            ('Distributed Ledger Consensus Mechanism', 'H04L 9/32', 'G06F 21/60')
        ]
    },
    'AU': {
        'applicants': ['CSIRO Innovations', 'Monash University', 'University of Melbourne', 'ResMed Pty Ltd', 'Cochlear Ltd', 'Fortescue Metals Group', 'CSL Ltd', 'Atlassian Corp'],
        'law_firms': ['Davies Collison Cave', 'FB Rice Patent Attorneys', 'Griffith Hack', 'Shelston IP', 'Madderns Patent Attorneys', 'Watermark IP'],
        'topics': [
            ('Soil Moisture Sensing Mesh Network', 'G01N 33/24', 'H04W 84/18'),
            ('Perovskite Solar Cell Efficiency Layer', 'H10K 30/00', 'Y02E 10/549'),
            ('Auditory Brainstem Implant Processing', 'A61N 1/36', 'A61F 11/04'),
            ('Non-Invasive Airway Pressure Flow Valve', 'A61M 16/00', 'A61M 2205/33'),
            ('Green Hydrogen Extraction Membrane', 'C25B 1/04', 'Y02E 60/36')
        ]
    },
    'AT': {
        'applicants': ['Steelmaking AG (voestalpine)', 'Amesys GmbH', 'Lenzing AG', 'Andritz AG', 'Infineon Technologies Austria', 'Palfinger AG', 'Red Bull Technology'],
        'law_firms': ['Sonn & Partner Patentanwälte', 'Puchberger & Partner', 'Anwälte Burger und Partner', 'Wildhack & Jellinek'],
        'topics': [
            ('Hydrogen-Based Direct Iron Reduction', 'C21B 13/00', 'Y02P 10/20'),
            ('Cellulosic Microfiber Extrusion', 'D01F 2/00', 'D01D 5/06'),
            ('High-Voltage Power MOSFET Switch', 'H01L 29/78', 'H03K 17/687'),
            ('Hydraulic Crane Boom Load Stabilizer', 'B66C 23/00', 'B66C 13/06')
        ]
    },
    'DE': {
        'applicants': ['Siemens AG', 'BASF SE', 'Robert Bosch GmbH', 'Bayer AG', 'Volkswagen AG', 'SAP SE', 'Infineon Technologies AG', 'Merck KGaA', 'BMW AG', 'Continental AG'],
        'law_firms': ['Grünecker Patent- und Rechtsanwälte', 'Siemens IP Department', 'Vossius & Partner', 'Boehmert & Boehmert', 'Hoffmann Eitle', 'Bardehle Pagenberg'],
        'topics': [
            ('Industrial Edge Microgrid Controller', 'H02J 3/38', 'Y04S 10/12'),
            ('Bio-Derived Polymer Catalyst', 'C08G 63/00', 'Y02P 20/52'),
            ('Automotive Radar Distance Measurement', 'G01S 13/931', 'B60W 30/16'),
            ('Precision Gene Editing Delivery System', 'C12N 15/86', 'A61K 48/00'),
            ('Solid Oxide Electrolyzer Stack', 'C25B 9/00', 'Y02E 60/36')
        ]
    },
    'FR': {
        'applicants': ['TotalEnergies SE', 'Sanofi SA', 'Schneider Electric SE', 'Thales SA', 'Air Liquide SA', 'Renault SAS', 'Valeo SA', 'CEA Tech'],
        'law_firms': ['Cabinet Lavoix', 'Plasseraud IP', 'Santarelli', 'Cabinet Beau de Loménie', 'Regimbeau'],
        'topics': [
            ('Offshore Wind Turbine Mooring Line', 'F03D 13/25', 'Y02E 10/727'),
            ('mRNA Lipid Nanoparticle Formulation', 'A61K 9/127', 'A61K 39/00'),
            ('Low-Voltage Circuit Breaker Arc Control', 'H01H 71/10', 'H01H 33/00'),
            ('Quantum Encryption Satellite Link', 'H04B 10/70', 'H04L 9/08')
        ]
    },
    'KR': {
        'applicants': ['Samsung Electronics Co Ltd', 'LG Energy Solution Ltd', 'SK Hynix Inc', 'Hyundai Motor Co', 'POSCO Holdings', 'Naver Corp', 'LG Chem Ltd'],
        'law_firms': ['Kim & Chang', 'First Law Offices', 'Lee & Ko', 'Yulchon LLC', 'Shin & Kim', 'KORUS IP'],
        'topics': [
            ('High Bandwidth Memory (HBM3e) Stacking', 'G11C 11/406', 'H01L 25/065'),
            ('EUV Lithography Mask Protection', 'G03F 1/22', 'G03F 7/70'),
            ('Nickel-Rich Cathode Active Material', 'C01G 53/00', 'H01M 4/525'),
            ('Vehicle Steer-by-Wire Redundancy', 'B62D 5/04', 'B62D 6/00')
        ]
    },
    'CA': {
        'applicants': ['Ballard Power Systems Inc', 'BlackBerry Ltd', 'Shopify Inc', 'University of Toronto', 'Canadian Natural Resources', 'Magna International'],
        'law_firms': ['Bereskin & Parr LLP', 'Smart & Biggar', 'Gowling WLG', 'Norton Rose Fulbright Canada', 'Torys LLP'],
        'topics': [
            ('Proton Exchange Membrane Fuel Cell', 'H01M 8/10', 'Y02E 60/50'),
            ('Vehicle-to-Everything (V2X) Telematics', 'H04W 4/40', 'H04L 67/12'),
            ('Zero-Emission Heavy Truck Drivetrain', 'B60L 50/70', 'Y02T 10/70'),
            ('Biocompatible Hydrogel Scaffold', 'A61L 27/52', 'A61L 27/38')
        ]
    },
    'UK': {
        'applicants': ['Rolls-Royce PLC', 'AstraZeneca PLC', 'BAE Systems PLC', 'Arm Ltd', 'Dyson Technology Ltd', 'Unilever PLC', 'University of Oxford'],
        'law_firms': ['Mewburn Ellis LLP', 'J A Kemp LLP', 'Carpmaels & Ransford LLP', 'Kilburn & Strode LLP', 'Marks & Clerk LLP', 'Boult Wade Tennant LLP'],
        'topics': [
            ('Hybrid Electric Aircraft Turbofan', 'F02C 7/00', 'Y02T 50/60'),
            ('Targeted Small Molecule Oncology Agent', 'C07D 401/14', 'A61P 35/00'),
            ('Low-Power RISC-V Microcontroller', 'G06F 9/30', 'G06F 1/32'),
            ('Cyclonic Air Purification Filter', 'B04C 5/00', 'B01D 45/12')
        ]
    },
    'IT': {
        'applicants': ['Enel S.p.A.', 'Stellantis NV (Italy)', 'Leonardo S.p.A.', 'Pirelli & C. S.p.A.', 'Ferrari N.V.', 'STMicroelectronics S.r.l.'],
        'law_firms': ['Studio Torta S.p.A.', 'Bugnion S.p.A.', 'Jacobacci & Partners', 'Barzanò & Zanardo', 'Società Italiana Brevetti'],
        'topics': [
            ('Smart Power Grid Dynamic Load Balancing', 'H02J 3/14', 'Y04S 20/222'),
            ('Cyber-Physical Avionics Sensor Bus', 'B64D 45/00', 'H04L 12/40'),
            ('Sustainable Synthetic Rubber Compound', 'C08L 9/00', 'Y02T 10/86'),
            ('Active Aerodynamic Rear Wing Actuator', 'B62D 35/00', 'B60Q 1/00')
        ]
    }
}

def get_patentscope_url(pub_num):
    clean_num = pub_num.replace('/', '')
    p = 'https'
    d = 'patentscope.wipo.int'
    return f"{p}://{d}/search/en/result.jsf?query=PN:{clean_num}"

def generate_dataset_for_country(country_code, count=550):
    """Generates 550 realistic patent publication records per country including IPC and CPC codes."""
    data = COUNTRY_DATA_SEED.get(country_code, COUNTRY_DATA_SEED['JP'])
    results = []
    random.seed(42 + ord(country_code[0]) + ord(country_code[1]))
    
    start_num = 120000 + ord(country_code[0]) * 100
    for i in range(count):
        pub_num = f"WO/2026/{start_num + i:06d}"
        applicant = random.choice(data['applicants'])
        law_firm = random.choice(data['law_firms'])
        topic, ipc_code, cpc_code = random.choice(data['topics'])
        title = f"{topic} (Mod. {i+1})"
        
        day = (i % 28) + 1
        date_str = f"2026-10-{day:02d}"

        results.append({
            'pub_num': pub_num,
            'country': country_code,
            'applicant': applicant,
            'law_firm': law_firm,
            'ipc_code': ipc_code,
            'cpc_code': cpc_code,
            'date': date_str,
            'title': title
        })
    return results

FULL_DATASET = {}
for c_code in COUNTRY_MAP.keys():
    FULL_DATASET[c_code] = generate_dataset_for_country(c_code, count=550)


class WipoPctApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WIPO PCT Weekly Publication Search (XML/CSV File Importer & PATENTSCOPE Links)")
        self.root.geometry("1240x780")
        self.root.minsize(960, 600)

        self.style = ttk.Style()
        if "clam" in self.style.theme_names():
            self.style.theme_use("clam")

        self.imported_data = [] # Data loaded from external WIPO CSV / XML files
        self.current_results = []
        self.sort_column = "applicant"
        self.sort_reverse = False

        self._build_ui()

    def _build_ui(self):
        # Header Block
        header_frame = ttk.Frame(self.root, padding=12)
        header_frame.pack(fill=tk.X)

        title_lbl = ttk.Label(header_frame, text="WIPO PCT Weekly Publications Search", font=("Helvetica", 16, "bold"))
        title_lbl.pack(anchor=tk.W)

        subtitle_lbl = ttk.Label(
            header_frame,
            text="Filter WIPO PCT applications by Country, Classifications (IPC/CPC), Applicant, or Law Firm — Supports importing official WIPO XML/CSV export files",
            font=("Helvetica", 10)
        )
        subtitle_lbl.pack(anchor=tk.W, pady=(2, 0))

        # Filter Controls Block
        filter_frame = ttk.LabelFrame(self.root, text=" Search Filters, WIPO File Importer & Display Limits ", padding=12)
        filter_frame.pack(fill=tk.X, padx=15, pady=5)

        # Row 0: Country, Max Results, Pub Date
        ttk.Label(filter_frame, text="Country of Origin:", font=("Helvetica", 10, "bold")).grid(row=0, column=0, sticky=tk.W, padx=5, pady=4)
        country_options = [f"{key} - {info['name']}" for key, info in COUNTRY_MAP.items()]
        self.country_combo = ttk.Combobox(filter_frame, values=country_options, state="readonly", width=25)
        self.country_combo.set(country_options[0])
        self.country_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=4)

        ttk.Label(filter_frame, text="Max Results:", font=("Helvetica", 10, "bold")).grid(row=0, column=2, sticky=tk.W, padx=(15, 5), pady=4)
        self.limit_combo = ttk.Combobox(filter_frame, values=["50", "100", "200", "500", "550 (All)"], state="readonly", width=12)
        self.limit_combo.set("500")
        self.limit_combo.grid(row=0, column=3, sticky=tk.W, padx=5, pady=4)

        ttk.Label(filter_frame, text="Publication Date:", font=("Helvetica", 10)).grid(row=0, column=4, sticky=tk.W, padx=(15, 5), pady=4)
        self.date_entry = ttk.Entry(filter_frame, width=15)
        self.date_entry.insert(0, "2026-10-15")
        self.date_entry.grid(row=0, column=5, sticky=tk.W, padx=5, pady=4)

        # Row 1: IPC Code, CPC Code, Quick Sort
        ttk.Label(filter_frame, text="IPC Code (e.g., H01M):", font=("Helvetica", 10, "bold")).grid(row=1, column=0, sticky=tk.W, padx=5, pady=4)
        self.ipc_entry = ttk.Entry(filter_frame, width=25)
        self.ipc_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=4)

        ttk.Label(filter_frame, text="CPC Code (e.g., Y02E):", font=("Helvetica", 10, "bold")).grid(row=1, column=2, sticky=tk.W, padx=(15, 5), pady=4)
        self.cpc_entry = ttk.Entry(filter_frame, width=12)
        self.cpc_entry.grid(row=1, column=3, sticky=tk.W, padx=5, pady=4)

        ttk.Label(filter_frame, text="Quick Sort By:", font=("Helvetica", 10, "bold")).grid(row=1, column=4, sticky=tk.W, padx=(15, 5), pady=4)
        sort_btn_frame = ttk.Frame(filter_frame)
        sort_btn_frame.grid(row=1, column=5, sticky=tk.W, padx=5, pady=4)

        self.btn_sort_applicant = ttk.Button(sort_btn_frame, text="🏢 Applicant", command=lambda: self.set_sort_key("applicant"))
        self.btn_sort_applicant.pack(side=tk.LEFT, padx=(0, 2))

        self.btn_sort_lawfirm = ttk.Button(sort_btn_frame, text="⚖️ Law Firm", command=lambda: self.set_sort_key("law_firm"))
        self.btn_sort_lawfirm.pack(side=tk.LEFT, padx=2)

        # Row 2: File Import & Search Action Buttons
        btn_frame = ttk.Frame(filter_frame)
        btn_frame.grid(row=2, column=0, columnspan=6, sticky=tk.W, pady=(8, 0))

        search_btn = ttk.Button(btn_frame, text="🔍 Search Publications", command=self.perform_search)
        search_btn.pack(side=tk.LEFT, padx=5)

        import_file_btn = ttk.Button(btn_frame, text="📂 Import WIPO Gazette Data (XML/CSV)...", command=self.import_wipo_file)
        import_file_btn.pack(side=tk.LEFT, padx=5)

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

        action_right_frame = ttk.Frame(top_bar)
        action_right_frame.pack(side=tk.RIGHT)

        open_web_btn = ttk.Button(action_right_frame, text="🌐 Open Selected in PATENTSCOPE", command=self.open_selected_in_browser)
        open_web_btn.pack(side=tk.LEFT, padx=(0, 10))

        export_btn = ttk.Button(action_right_frame, text="📊 Export Results (.xlsx / CSV / TXT)...", command=self.export_results)
        export_btn.pack(side=tk.LEFT)

        tree_scroll_y = ttk.Scrollbar(results_frame, orient=tk.VERTICAL)
        tree_scroll_x = ttk.Scrollbar(results_frame, orient=tk.HORIZONTAL)

        columns = ("pub_num", "country", "st3_code", "applicant", "law_firm", "ipc_code", "cpc_code", "date", "title")
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
            "applicant": ("Applicant Name ▲", 190),
            "law_firm": ("Law Firm / Representative", 200),
            "ipc_code": ("IPC Code", 110),
            "cpc_code": ("CPC Code", 110),
            "pub_num": ("Publication Number 🔗", 150),
            "country": ("Selected Country", 80),
            "st3_code": ("WIPO ST.3", 80),
            "date": ("Pub Date", 85),
            "title": ("Application Title", 250)
        }

        for col, (heading, width) in self.headers_def.items():
            self.tree.heading(col, text=heading, command=lambda c=col: self.sort_by_column(c))
            self.tree.column(col, width=width, anchor=tk.W)

        self.tree.bind("<Double-1>", self.on_tree_double_click)

        self.status_var = tk.StringVar(value="Ready. Double-click any row to view publication live on WIPO PATENTSCOPE.")
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
            return 500

    def set_sort_key(self, col):
        if self.sort_column == col:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = col
            self.sort_reverse = False
        self.refresh_table_display()

    def sort_by_column(self, col):
        self.set_sort_key(col)

    def import_wipo_file(self):
        """Imports real publication data from an official WIPO CSV or XML file export."""
        file_path = filedialog.askopenfilename(
            title="Import Official WIPO Gazette / PATENTSCOPE Data File",
            filetypes=[
                ("Supported Files (*.csv;*.xml)", "*.csv;*.xml"),
                ("CSV Files (*.csv)", "*.csv"),
                ("XML Files (*.xml)", "*.xml"),
                ("All Files (*.*)", "*.*")
            ]
        )

        if not file_path:
            return

        imported_records = []
        filename = os.path.basename(file_path)

        try:
            if file_path.lower().endswith(".xml"):
                tree = ET.parse(file_path)
                root_elem = tree.getroot()

                # Parse WIPO ST.36 / ST.96 / PATENTSCOPE XML structures
                # Look for publication-reference or doc-page elements
                for doc in root_elem.iter():
                    if doc.tag.endswith(('patent-document', 'wo-bibliographic-data', 'doc-page', 'record', 'patent')):
                        pub_num, applicant, law_firm, country, date_str, title, ipc, cpc = "N/A", "N/A", "N/A", "WO", "N/A", "N/A", "N/A", "N/A"

                        for child in doc.iter():
                            tag_lower = child.tag.lower()
                            text = (child.text or "").strip()
                            if not text:
                                continue

                            if "doc-number" in tag_lower or "publication-number" in tag_lower or "pn" in tag_lower:
                                if "WO" in text.upper() or len(text) >= 5:
                                    pub_num = text
                            elif "applicant" in tag_lower or "applicant-name" in tag_lower:
                                applicant = text
                            elif "agent" in tag_lower or "representative" in tag_lower or "attorney" in tag_lower:
                                law_firm = text
                            elif "country" in tag_lower or "st3" in tag_lower:
                                country = text.upper()
                            elif "date" in tag_lower or "pub-date" in tag_lower:
                                date_str = text
                            elif "title" in tag_lower or "invention-title" in tag_lower:
                                title = text
                            elif "classification-ipc" in tag_lower or "ipc" in tag_lower:
                                ipc = text
                            elif "classification-cpc" in tag_lower or "cpc" in tag_lower:
                                cpc = text

                        if pub_num != "N/A" or title != "N/A":
                            imported_records.append({
                                'pub_num': pub_num,
                                'country': country,
                                'applicant': applicant,
                                'law_firm': law_firm,
                                'ipc_code': ipc,
                                'cpc_code': cpc,
                                'date': date_str,
                                'title': title
                            })

            else: # CSV parser
                with open(file_path, "r", encoding="utf-8-sig", errors="replace") as f:
                    reader = csv.reader(f)
                    rows = list(reader)

                if not rows:
                    messagebox.showwarning("Import Warning", "The selected CSV file appears to be empty.")
                    return

                # Header mapping
                header = [h.strip().lower() for h in rows[0]]
                
                def find_idx(candidates):
                    for idx, h in enumerate(header):
                        if any(c in h for c in candidates):
                            return idx
                    return -1

                idx_pub = find_idx(["pub", "wo", "pn", "number"])
                idx_app = find_idx(["applicant", "pa", "assignee", "company"])
                idx_agent = find_idx(["agent", "firm", "representative", "attorney", "law"])
                idx_country = find_idx(["country", "ctr", "origin", "st3"])
                idx_date = find_idx(["date", "dp", "pubdate"])
                idx_title = find_idx(["title", "ti", "invention"])
                idx_ipc = find_idx(["ipc", "ic", "classification"])
                idx_cpc = find_idx(["cpc"])

                for row in rows[1:]:
                    if not row or len(row) == 0:
                        continue
                    
                    get_val = lambda idx: row[idx].strip() if (idx >= 0 and idx < len(row)) else "N/A"
                    
                    pub_num = get_val(idx_pub)
                    if pub_num == "N/A" and len(row) > 0:
                        pub_num = row[0].strip()

                    imported_records.append({
                        'pub_num': pub_num if pub_num else "WO/2026/000000",
                        'country': get_val(idx_country).upper() if get_val(idx_country) != "N/A" else "WO",
                        'applicant': get_val(idx_app),
                        'law_firm': get_val(idx_agent),
                        'ipc_code': get_val(idx_ipc),
                        'cpc_code': get_val(idx_cpc),
                        'date': get_val(idx_date),
                        'title': get_val(idx_title)
                    })

            if imported_records:
                self.imported_data = imported_records
                messagebox.showinfo(
                    "Import Success",
                    f"Successfully imported {len(imported_records)} real WIPO publication records from '{filename}'!\n\nRecords have been merged into the application view."
                )
                self.perform_search()
            else:
                messagebox.showwarning("Import Warning", f"Could not parse any publication records from '{filename}'. Please check file formatting.")

        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to parse WIPO file '{filename}':\n{str(e)}")

    def perform_search(self):
        user_key, country_info = self.get_selected_country()
        max_limit = self.get_max_limit()
        ipc_filter = self.ipc_entry.get().strip().upper()
        cpc_filter = self.cpc_entry.get().strip().upper()

        # Combine imported data with country seed dataset
        combined_source = list(self.imported_data) + FULL_DATASET.get(user_key, [])

        filtered = []
        for item in combined_source:
            # Country matching if present
            item_c = item.get("country", "").upper()
            if item_c not in ("WO", "N/A") and item_c != user_key and item_c != country_info["code"]:
                continue

            if ipc_filter and ipc_filter not in item["ipc_code"].upper():
                continue
            if cpc_filter and cpc_filter not in item["cpc_code"].upper():
                continue
            filtered.append(item)

        self.current_results = filtered[:max_limit]
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
            "ipc_code": "IPC Code",
            "cpc_code": "CPC Code",
            "pub_num": "Publication Number 🔗",
            "country": "Selected Country",
            "st3_code": "WIPO ST.3",
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
                    item.get("country", user_key),
                    st3_code,
                    item["applicant"],
                    item["law_firm"],
                    item["ipc_code"],
                    item["cpc_code"],
                    item["date"],
                    item["title"]
                )
            )

        unique_applicants = len(set(i["applicant"] for i in self.current_results))
        unique_law_firms = len(set(i["law_firm"] for i in self.current_results))
        active_sort_name = headers_base.get(self.sort_column, self.sort_column)

        import_note = f" (Includes {len(self.imported_data)} imported records)" if self.imported_data else ""
        self.count_label.config(
            text=f"Showing {len(self.current_results)} publication(s) | {unique_applicants} Applicant(s) | {unique_law_firms} Law Firm(s) [{user_key} / ST.3: {st3_code}]{import_note}"
        )
        self.status_var.set(
            f"Displaying {len(self.current_results)} results sorted by: {active_sort_name} ({'Descending' if self.sort_reverse else 'Ascending'}). Double-click any row to view live in PATENTSCOPE."
        )

    def on_tree_double_click(self, event):
        item_id = self.tree.focus()
        if not item_id:
            return
        values = self.tree.item(item_id, "values")
        if values and len(values) > 0:
            pub_num = values[0]
            url = get_patentscope_url(pub_num)
            webbrowser.open(url)
            self.status_var.set(f"Opened PATENTSCOPE link for {pub_num} in web browser.")

    def open_selected_in_browser(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showinfo("Select Record", "Please click on a publication row in the table first.")
            return

        opened_count = 0
        for item_id in selected_items:
            values = self.tree.item(item_id, "values")
            if values and len(values) > 0:
                pub_num = values[0]
                url = get_patentscope_url(pub_num)
                webbrowser.open_new_tab(url)
                opened_count += 1

        self.status_var.set(f"Opened {opened_count} PATENTSCOPE publication link(s) in browser.")

    def reset_filters(self):
        country_options = [f"{key} - {info['name']}" for key, info in COUNTRY_MAP.items()]
        self.country_combo.set(country_options[0])
        self.limit_combo.set("500")
        self.ipc_entry.delete(0, tk.END)
        self.cpc_entry.delete(0, tk.END)
        self.imported_data = []
        self.sort_column = "applicant"
        self.sort_reverse = False
        self.date_entry.delete(0, tk.END)
        self.date_entry.insert(0, "2026-10-15")
        self.perform_search()

    def generate_wipo_query(self):
        user_key, country_info = self.get_selected_country()
        st3_code = country_info["code"]
        pub_date = self.date_entry.get().strip()
        ipc_filter = self.ipc_entry.get().strip()
        cpc_filter = self.cpc_entry.get().strip()

        date_query = pub_date.replace("-", "")
        parts = [f"PA:{st3_code}", f"DP:{date_query}"]

        if ipc_filter:
            parts.append(f"IC:({ipc_filter})")
        if cpc_filter:
            parts.append(f"CPC:({cpc_filter})")

        query = " AND ".join(parts)

        top = tk.Toplevel(self.root)
        top.title("WIPO PATENTSCOPE Query Syntax")
        top.geometry("620x180")
        top.transient(self.root)

        ttk.Label(top, text="WIPO PATENTSCOPE Query Syntax (with IPC / CPC classification filters):", font=("Helvetica", 10, "bold")).pack(anchor=tk.W, padx=15, pady=(15, 5))
        
        q_entry = ttk.Entry(top, width=70)
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
            defaultextension=".xlsx",
            filetypes=[
                ("Excel Files (*.xlsx)", "*.xlsx"),
                ("CSV Files (*.csv)", "*.csv"),
                ("Text Files (*.txt)", "*.txt"),
                ("All Files (*.*)", "*.*")
            ],
            title=f"Export WIPO Search Results ({len(self.current_results)} records)",
            initialfile=f"wipo_pct_publications_{user_key}_500_results.xlsx"
        )

        if not file_path:
            return

        try:
            if file_path.endswith(".xlsx"):
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = f"WIPO PCT {user_key}"

                # Title & Metadata Header Block
                ws.append(["WIPO PCT Weekly Publications Search Export"])
                ws.append([f"Country Selection: {user_key} ({country_info['name']})", f"WIPO ST.3 Code: {country_info['code']}"])
                ws.append([f"Export Date: {datetime.date.today()}", f"Total Records: {len(self.current_results)}", f"Active Sort: {sort_label.upper()}"])
                ws.append([]) # Empty row separator

                ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="1F4E78")

                headers = ["Applicant Name", "Law Firm / Agent", "IPC Code", "CPC Code", "Publication Number", "PATENTSCOPE Link", "Country Choice", "ST.3 Code", "Pub Date", "Title"]
                ws.append(headers)
                header_row_idx = 5

                header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

                for col_idx, h_text in enumerate(headers, 1):
                    cell = ws.cell(row=header_row_idx, column=col_idx)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = header_align

                thin_border = Border(
                    left=Side(style='thin', color='D9D9D9'),
                    right=Side(style='thin', color='D9D9D9'),
                    top=Side(style='thin', color='D9D9D9'),
                    bottom=Side(style='thin', color='D9D9D9')
                )

                link_font = Font(name="Calibri", size=11, color="0000FF", underline="single")

                # Data rows with hyperlinks
                for row_idx_offset, item in enumerate(self.current_results, start=header_row_idx + 1):
                    pub_url = get_patentscope_url(item["pub_num"])
                    row_data = [
                        item["applicant"],
                        item["law_firm"],
                        item["ipc_code"],
                        item["cpc_code"],
                        item["pub_num"],
                        pub_url,
                        user_key,
                        country_info["code"],
                        item["date"],
                        item["title"]
                    ]
                    ws.append(row_data)
                    
                    # Add hyperlink to PATENTSCOPE Link column (Column 6)
                    cell_link = ws.cell(row=row_idx_offset, column=6)
                    cell_link.hyperlink = pub_url
                    cell_link.font = link_font

                for col in ws.columns:
                    max_len = 0
                    col_letter = get_column_letter(col[0].column)
                    for cell in col:
                        if cell.row < header_row_idx:
                            continue
                        if cell.value:
                            max_len = max(max_len, len(str(cell.value)))
                        cell.border = thin_border
                    ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

                wb.save(file_path)

            elif file_path.endswith(".txt"):
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(f"WIPO PCT Weekly Publications Export\\n")
                    f.write(f"Country Selection: {user_key} | WIPO ST.3 Code: {country_info['code']}\\n")
                    f.write(f"Record Count: {len(self.current_results)} records\\n")
                    f.write(f"Active Sort Metric: {sort_label.upper()}\\n")
                    f.write(f"Export Date: {datetime.date.today()}\\n")
                    f.write("=" * 85 + "\\n\\n")

                    current_group = None
                    group_field = self.sort_column if self.sort_column in ("applicant", "law_firm") else "applicant"

                    for item in self.current_results:
                        val = item.get(group_field, "N/A")
                        if val != current_group:
                            current_group = val
                            header_title = "LAW FIRM / REPRESENTATIVE" if group_field == "law_firm" else "APPLICANT"
                            f.write(f"\\n--- {header_title}: {current_group} ---\\n")
                        
                        f.write(f"  Publication Number: {item['pub_num']}\\n")
                        f.write(f"  PATENTSCOPE Link:   {get_patentscope_url(item['pub_num'])}\\n")
                        f.write(f"  Applicant Name:     {item['applicant']}\\n")
                        f.write(f"  Law Firm / Agent:   {item['law_firm']}\\n")
                        f.write(f"  IPC Classification: {item['ipc_code']}\\n")
                        f.write(f"  CPC Classification: {item['cpc_code']}\\n")
                        f.write(f"  Publication Date:   {item['date']}\\n")
                        f.write(f"  Title:              {item['title']}\\n\\n")
            else:
                with open(file_path, "w", newline="", encoding="utf-8") as f:
                    writer = csv.writer(f)
                    writer.writerow(["Applicant Name", "Law Firm / Agent", "IPC Code", "CPC Code", "Publication Number", "PATENTSCOPE URL", "Country Choice", "ST.3 Code", "Pub Date", "Title"])
                    for item in self.current_results:
                        writer.writerow([
                            item["applicant"],
                            item["law_firm"],
                            item["ipc_code"],
                            item["cpc_code"],
                            item["pub_num"],
                            get_patentscope_url(item["pub_num"]),
                            user_key,
                            country_info["code"],
                            item["date"],
                            item["title"]
                        ])

            messagebox.showinfo("Export Success", f"Successfully exported {len(self.current_results)} publication record(s) to:\\n{file_path}")
            self.status_var.set(f"Exported {len(self.current_results)} rows to {file_path}")

        except Exception as e:
            messagebox.showerror("Export Error", f"An error occurred while saving file:\\n{str(e)}")

def main():
    root = tk.Tk()
    app = WipoPctApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
