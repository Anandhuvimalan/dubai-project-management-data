"""
Dubai Civil Engineering Company - Project Management Data Generator
Generates comprehensive, realistic CSV data for a Dubai-based civil engineering firm
"""

import pandas as pd
import random
from faker import Faker
import datetime
import os

fake = Faker()
Faker.seed(42)
random.seed(42)

# === CONFIGURATION ===
NUM_CLIENTS = 50
NUM_EMPLOYEES = 400
NUM_CONTRACTORS = 100
NUM_SUPPLIERS = 80
NUM_EQUIPMENT = 150
NUM_PROJECTS = 100
START_DATE = datetime.date(2020, 1, 1)
END_DATE = datetime.date(2026, 12, 31)

# === DUBAI CIVIL ENGINEERING CONTEXT ===
DUBAI_AREAS = [
    "Business Bay", "Downtown Dubai", "Dubai Marina", "DIFC", "JLT",
    "Palm Jumeirah", "Dubai Hills Estate", "Meydan", "Dubai Creek Harbour",
    "Dubai South", "JAFZA", "Al Quoz Industrial", "DIP", "Expo City",
    "JVC", "Al Barsha", "Deira", "Bur Dubai", "Sheikh Zayed Road"
]

PROJECT_TYPES = [
    "High-Rise Residential Tower", "Commercial Office Building", "Mixed-Use Development",
    "Villa Development", "Infrastructure - Roads", "Infrastructure - Bridges",
    "Industrial Warehouse", "Hotel & Hospitality", "Retail Mall",
    "Marine & Coastal Works", "Utility Infrastructure", "School/Educational"
]

ENGINEERING_ROLES = {
    "PMO": ["Project Director", "Project Manager", "Senior Project Manager", "Project Coordinator", "Planning Engineer", "Scheduler"],
    "Engineering": ["Chief Engineer", "Senior Civil Engineer", "Civil Engineer", "Structural Engineer", "MEP Coordinator", "Design Engineer", "CAD Technician"],
    "Construction": ["Construction Manager", "Site Engineer", "Site Supervisor", "Quantity Surveyor", "QS Manager", "Contracts Manager"],
    "HSE": ["HSE Manager", "HSE Officer", "Safety Engineer", "Environmental Officer"],
    "QA/QC": ["QA/QC Manager", "QA/QC Engineer", "QC Inspector", "Materials Engineer"],
    "Commercial": ["Commercial Manager", "Estimation Engineer", "Procurement Manager", "Buyer"],
    "Finance": ["Finance Manager", "Accountant", "Cost Controller"],
    "Admin": ["HR Manager", "PRO", "Office Administrator", "Document Controller"]
}

CONTRACTOR_TYPES = [
    "Piling & Foundations", "Structural Steel", "Concrete Works", "MEP Works",
    "Facade & Cladding", "Waterproofing", "HVAC", "Electrical", "Plumbing",
    "Fire Fighting", "Landscaping", "Road Works", "Demolition", "Fit-Out"
]

SUPPLIER_CATEGORIES = [
    "Ready-Mix Concrete", "Steel Reinforcement", "Structural Steel", "Cement & Aggregates",
    "Formwork & Scaffolding", "MEP Materials", "Facade Materials", "Finishing Materials",
    "Safety Equipment", "Construction Chemicals"
]

EQUIPMENT_TYPES = [
    ("Tower Crane", 3500, 5500), ("Mobile Crane", 2000, 4000), ("Excavator", 800, 1500),
    ("Wheel Loader", 600, 1200), ("Backhoe Loader", 500, 900), ("Dump Truck", 400, 700),
    ("Concrete Pump", 1500, 2500), ("Batching Plant", 2000, 3500), ("Compactor", 300, 600),
    ("Generator", 200, 500), ("Forklift", 250, 450), ("Boom Lift", 400, 800),
    ("Scaffolding Set", 100, 200), ("Welding Machine", 50, 150)
]

PERMIT_TYPES = [
    ("Dubai Municipality - Building Permit", "DM-BP"), ("Dubai Municipality - NOC", "DM-NOC"),
    ("DEWA Connection", "DEWA"), ("RTA Access Permit", "RTA"), ("Civil Defense NOC", "CD"),
    ("Etisalat NOC", "ETIS"), ("DDA Approval", "DDA"), ("Trakhees Permit", "TRK"),
    ("Environmental Permit", "ENV"), ("Crane Permit", "CRANE")
]

DOCUMENT_TYPES = [
    "Shop Drawing", "As-Built Drawing", "Method Statement", "Material Submittal",
    "RFI", "NCR", "Site Instruction", "Variation Order", "Meeting Minutes",
    "Inspection Request", "Test Report", "Progress Report"
]

PHASE_NAMES = [
    ("Enabling Works", 0.05), ("Excavation & Shoring", 0.10), ("Piling", 0.12),
    ("Substructure", 0.15), ("Superstructure", 0.25), ("MEP Rough-In", 0.10),
    ("Facade", 0.08), ("Finishes", 0.10), ("Testing & Commissioning", 0.03), ("Handover", 0.02)
]

NATIONALITIES = ["UAE", "India", "Pakistan", "Philippines", "Egypt", "UK", "Jordan", "Lebanon", "Bangladesh", "Sri Lanka"]

# === HELPER FUNCTIONS ===
def random_date(start, end):
    if start >= end: return start
    return fake.date_between(start_date=start, end_date=end)

def generate_uae_phone():
    return f"+971-{random.choice(['50','52','54','55','56'])}-{random.randint(1000000,9999999)}"

def generate_trn():
    return f"100{random.randint(100000000, 999999999)}"

def generate_license():
    return f"{random.randint(100000, 999999)}"

# === DATA GENERATORS ===
def generate_clients():
    clients = []
    client_names = [
        "Emaar Properties", "DAMAC Properties", "Nakheel", "Dubai Holding", "Meraas",
        "Azizi Developments", "Sobha Realty", "Al Futtaim Group", "Majid Al Futtaim",
        "Dubai RTA", "Dubai Municipality", "DEWA", "Expo 2020 LLC", "DP World",
        "JAFZA", "Dubai South", "Etihad Rail", "Aldar Properties", "Mubadala",
        "Abu Dhabi Ports", "Bloom Holding", "Eagle Hills", "Select Group",
        "Danube Properties", "Binghatti Developers", "Omniyat", "Al Wasl Properties",
        "Dubai Properties", "Deyaar Development", "Union Properties",
        "Tilal Properties", "Al Habtoor Group", "Jumeirah Group", "Arabtec Construction",
        "Drake & Scull", "Al Ghurair Properties", "Limitless", "Tabreed", "Enoc",
        "Emirates NBD Properties", "FAB Properties", "RAK Properties", "Sharjah Holding",
        "Al Ain Properties", "Fujairah Development", "Ajman Properties", "Palma Holding",
        "Reportage Properties", "Mag Property Development", "Ellington Properties", "Samana Developers"
    ]
    for i, name in enumerate(client_names[:NUM_CLIENTS], 1):
        # Generate contact person first, then derive email from their name
        contact_first = fake.first_name()
        contact_last = fake.last_name()
        contact_person = f"{contact_first} {contact_last}"
        company_domain = name.lower().replace(' ', '').replace('&', '').replace('.', '')
        email = f"{contact_first.lower()}.{contact_last.lower()}@{company_domain}.ae"
        
        clients.append({
            "client_id": f"CL-{i:03d}",
            "client_name": name,
            "client_type": random.choice(["Government", "Semi-Government", "Private Developer", "Contractor"]),
            "trade_license": generate_license(),
            "trn": generate_trn(),
            "contact_person": contact_person,
            "designation": random.choice(["Project Director", "Contracts Manager", "Procurement Manager"]),
            "email": email,
            "phone": generate_uae_phone(),
            "address": f"{random.choice(DUBAI_AREAS)}, Dubai, UAE",
            "payment_terms": random.choice(["Net 30", "Net 45", "Net 60", "Net 90"]),
            "credit_limit_aed": random.choice([500000, 1000000, 2000000, 5000000, 10000000])
        })
    return pd.DataFrame(clients)

def generate_employees():
    employees = []
    emp_id = 1
    for dept, roles in ENGINEERING_ROLES.items():
        count = random.randint(15, 35) if dept in ["Engineering", "Construction"] else random.randint(5, 15)
        for _ in range(count):
            if emp_id > NUM_EMPLOYEES: break
            role = random.choice(roles)
            # Generate name first, then derive email from the actual name
            first_name = fake.first_name()
            last_name = fake.last_name()
            full_name = f"{first_name} {last_name}"
            email = f"{first_name.lower()}.{last_name.lower()}@alrashdeng.ae"
            
            employees.append({
                "employee_id": f"EMP-{emp_id:04d}",
                "full_name": full_name,
                "department": dept,
                "designation": role,
                "nationality": random.choice(NATIONALITIES),
                "joining_date": random_date(datetime.date(2018, 1, 1), datetime.date(2025, 6, 1)),
                "salary_aed": random.randint(5000, 45000) if "Manager" not in role else random.randint(25000, 70000),
                "mmup_license": f"MMUP-{random.randint(10000, 99999)}" if "Engineer" in role else None,
                "visa_status": random.choice(["Employment Visa", "Golden Visa", "Mission Visa"]),
                "email": email,
                "phone": generate_uae_phone()
            })
            emp_id += 1
    return pd.DataFrame(employees[:NUM_EMPLOYEES])

def generate_contractors():
    contractors = []
    for i in range(1, NUM_CONTRACTORS + 1):
        specialty = random.choice(CONTRACTOR_TYPES)
        company_name = f"{fake.company()} {specialty}" if random.random() > 0.5 else fake.company()
        
        # Generate contact person first, then derive email from their name
        contact_first = fake.first_name()
        contact_last = fake.last_name()
        contact_person = f"{contact_first} {contact_last}"
        # Create domain from company name (simplified)
        company_domain = company_name.split()[0].lower().replace(',', '').replace('.', '') + "contracting"
        email = f"{contact_first.lower()}.{contact_last.lower()}@{company_domain}.ae"
        
        contractors.append({
            "contractor_id": f"CON-{i:03d}",
            "company_name": company_name,
            "specialty": specialty,
            "trade_license": generate_license(),
            "cicpa_registration": f"CICPA-{random.randint(10000, 99999)}",
            "trn": generate_trn(),
            "contact_person": contact_person,
            "email": email,
            "phone": generate_uae_phone(),
            "address": f"{random.choice(['Al Quoz', 'DIP', 'JAFZA', 'Ras Al Khor', 'Sharjah Industrial'])}, UAE",
            "rating": random.choice(["A", "A", "B", "B", "B", "C"]),
            "prequalified": random.choice([True, True, True, False])
        })
    return pd.DataFrame(contractors)

def generate_suppliers():
    suppliers = []
    for i in range(1, NUM_SUPPLIERS + 1):
        category = random.choice(SUPPLIER_CATEGORIES)
        company_name = fake.company()
        
        # Generate contact person first, then derive email from their name
        contact_first = fake.first_name()
        contact_last = fake.last_name()
        contact_person = f"{contact_first} {contact_last}"
        # Create domain from company name
        company_domain = company_name.split()[0].lower().replace(',', '').replace('.', '') + "trading"
        email = f"{contact_first.lower()}.{contact_last.lower()}@{company_domain}.ae"
        
        suppliers.append({
            "supplier_id": f"SUP-{i:03d}",
            "company_name": company_name,
            "category": category,
            "trade_license": generate_license(),
            "trn": generate_trn(),
            "contact_person": contact_person,
            "email": email,
            "phone": generate_uae_phone(),
            "address": f"{random.choice(['Sharjah', 'Ajman', 'RAK', 'Dubai', 'Abu Dhabi'])}, UAE",
            "lead_time_days": random.randint(3, 30),
            "payment_terms": random.choice(["Cash", "Net 30", "Net 45", "Net 60"]),
            "approved": random.choice([True, True, True, False])
        })
    return pd.DataFrame(suppliers)

def generate_equipment():
    equipment = []
    for i in range(1, NUM_EQUIPMENT + 1):
        eq_type, min_rate, max_rate = random.choice(EQUIPMENT_TYPES)
        equipment.append({
            "equipment_id": f"EQ-{i:04d}",
            "equipment_type": eq_type,
            "model": f"{random.choice(['CAT', 'Komatsu', 'Liebherr', 'Volvo', 'JCB', 'Potain', 'Hitachi'])} {fake.bothify('??-###')}",
            "serial_number": fake.bothify('SN-########'),
            "ownership": random.choice(["Owned", "Owned", "Rented", "Leased"]),
            "daily_rate_aed": random.randint(min_rate, max_rate),
            "status": random.choice(["Available", "Available", "In Use", "In Use", "In Use", "Under Maintenance"]),
            "current_project": None,
            "last_service_date": random_date(datetime.date(2024, 1, 1), datetime.date(2025, 12, 1)),
            "operator_required": random.choice([True, True, False])
        })
    return pd.DataFrame(equipment)

def generate_projects(clients_df, employees_df):
    projects = []
    project_names = [
        "Al Khail Heights Tower", "Business Bay Plaza", "Marina Residences Phase 2",
        "Downtown Commercial Tower", "Palm Beach Villas", "Creek Harbour Mixed-Use",
        "Dubai South Logistics Hub", "JVC Community Center", "DIFC Office Tower",
        "Meydan One Mall Extension", "Expo City Pavilion", "Al Quoz Warehouse Complex",
        "JLT Office Renovation", "Dubai Hills Villa Cluster", "Deira Waterfront",
        "Bur Dubai Heritage Hotel", "Festival City Expansion", "Silicon Oasis Tech Park",
        "Sports City Stadium", "Motor City Showroom", "Green Community Villas",
        "Academic City Campus", "Healthcare City Hospital Wing", "Media City Studios",
        "Internet City Office Block", "Knowledge Park College", "Investment Park Factory",
        "Logistics City Distribution Center", "Maritime City Terminal", "Textile City Warehouse",
        "Gold Souk Extension", "Spice Souk Renovation", "Fish Market Redevelopment",
        "Metro Route 2020 Station", "Tram Extension Phase 2", "Bus Depot Facility",
        "Water Treatment Plant", "Sewage Pumping Station", "Power Substation",
        "Beach Access Infrastructure", "Cycling Track Network", "Public Park Development",
        "School Building Complex", "Mosque Construction", "Community Hall",
        "Emirates Towers Phase 2", "World Trade Center Annex", "Frame District Mall",
        "La Mer Commercial Hub", "Bluewaters Office Tower", "City Walk Extension",
        "Al Seef Heritage Hotel", "Ras Al Khor Industrial", "DIP Logistics Warehouse",
        "Mudon Villa Community", "Arabian Ranches Phase 3", "Damac Hills Clubhouse",
        "JBR Promenade Extension", "The Walk at JBR Phase 2", "Burj Area Boulevard",
        "Opera District Residences", "Address Hotel Expansion", "Vida Creek Tower",
        "Rove City Walk Hotel", "SE7EN Residences Palm", "One Za'abeel Tower B",
        "ICD Brookfield Tower 2", "Festival Plaza Extension", "Dragon Mall Phase 2",
        "Ibn Battuta Mall Annex", "Mall of Emirates Wing", "Dubai Mall Extension",
        "Outlet Village Phase 2", "Citizens School Campus", "GEMS Academy Building",
        "Dubai Medical College", "Dubai Science Park HQ", "Techno Hub Office",
        "Smart City Data Center", "Solar Park Operations Center", "Wind Tower Research Lab",
        "Al Maktoum Airport Cargo", "Airport Tunnel Project", "Sheikh Mohammed Bridge",
        "Business Bay Canal Walk", "Dubai Water Canal Phase 2", "Al Shindagha Bridge",
        "Floating Bridge Expansion", "Infinity Bridge South", "Tolerance Bridge Link",
        "RTA Bus Station Hub", "Dubai Metro Red Line Extension", "Green Line Station",
        "Palm Monorail Extension", "Cable Car Terminal", "Water Taxi Station",
        "District 2020 Office Block", "Dubai Future Labs", "Al Quoz Creative Hub",
        "Design District Phase 2", "Fashion Avenue Mall", "Gold Tower Commercial"
    ]

    managers = employees_df[employees_df['designation'].str.contains('Manager|Director', na=False)]['employee_id'].tolist()
    if not managers: managers = employees_df['employee_id'].tolist()[:10]

    for i, name in enumerate(project_names[:NUM_PROJECTS], 1):
        start = random_date(START_DATE, datetime.date(2025, 6, 1))
        duration_months = random.randint(12, 48)
        end = start + datetime.timedelta(days=duration_months * 30)
        budget = random.randint(5000000, 500000000)
        
        today = datetime.date.today()
        if end < today:
            status, completion = "Completed", 100
        elif start > today:
            status, completion = "Awarded", 0
        else:
            status = random.choice(["In Progress", "In Progress", "In Progress", "On Hold", "Delayed"])
            completion = random.randint(10, 95)

        projects.append({
            "project_id": f"PRJ-{i:03d}",
            "project_name": name,
            "project_type": random.choice(PROJECT_TYPES),
            "client_id": random.choice(clients_df['client_id'].tolist()),
            "project_manager_id": random.choice(managers),
            "location": random.choice(DUBAI_AREAS),
            "plot_number": f"Plot {random.randint(100, 999)}-{random.randint(1, 99)}",
            "dm_permit_no": f"DM-{start.year}-{random.randint(100000, 999999)}",
            "contract_value_aed": budget,
            "start_date": start,
            "end_date": end,
            "status": status,
            "completion_pct": completion,
            "gfa_sqm": random.randint(5000, 200000),
            "floors": random.randint(1, 80) if "Tower" in name or "High-Rise" in random.choice(PROJECT_TYPES) else random.randint(1, 10),
            "description": f"{random.choice(PROJECT_TYPES)} project in {random.choice(DUBAI_AREAS)}"
        })
    return pd.DataFrame(projects)

def generate_contracts(projects_df, contractors_df):
    contracts = []
    contract_id = 1
    for _, proj in projects_df.iterrows():
        # Main contract
        contracts.append({
            "contract_id": f"MC-{proj['project_id']}-001",
            "project_id": proj['project_id'],
            "contract_type": "Main Contract",
            "contractor_id": None,
            "fidic_type": random.choice(["FIDIC Red Book", "FIDIC Yellow Book", "FIDIC Silver Book"]),
            "contract_value_aed": proj['contract_value_aed'],
            "signed_date": proj['start_date'] - datetime.timedelta(days=random.randint(30, 90)),
            "commencement_date": proj['start_date'],
            "completion_date": proj['end_date'],
            "retention_pct": random.choice([5.0, 10.0]),
            "performance_bond_pct": 10.0,
            "advance_payment_pct": random.choice([0, 10, 15, 20]),
            "lad_per_day_aed": int(proj['contract_value_aed'] * 0.001),
            "defect_liability_months": random.choice([12, 24]),
            "status": "Active" if proj['status'] != "Completed" else "Closed"
        })
        # Subcontracts
        num_subs = random.randint(3, 8)
        for j in range(num_subs):
            sub_value = int(proj['contract_value_aed'] * random.uniform(0.02, 0.15))
            contracts.append({
                "contract_id": f"SC-{proj['project_id']}-{j+1:03d}",
                "project_id": proj['project_id'],
                "contract_type": "Subcontract",
                "contractor_id": random.choice(contractors_df['contractor_id'].tolist()),
                "fidic_type": "Subcontract Agreement",
                "contract_value_aed": sub_value,
                "signed_date": proj['start_date'] + datetime.timedelta(days=random.randint(0, 90)),
                "commencement_date": proj['start_date'] + datetime.timedelta(days=random.randint(30, 120)),
                "completion_date": proj['end_date'] - datetime.timedelta(days=random.randint(30, 90)),
                "retention_pct": 10.0,
                "performance_bond_pct": 10.0,
                "advance_payment_pct": 0,
                "lad_per_day_aed": int(sub_value * 0.002),
                "defect_liability_months": 12,
                "status": random.choice(["Active", "Active", "Completed"])
            })
    return pd.DataFrame(contracts)

def generate_project_phases(projects_df):
    phases = []
    for _, proj in projects_df.iterrows():
        proj_duration = (proj['end_date'] - proj['start_date']).days
        current_start = proj['start_date']
        for phase_name, weight in PHASE_NAMES:
            phase_duration = int(proj_duration * weight)
            phase_end = current_start + datetime.timedelta(days=phase_duration)
            today = datetime.date.today()
            if phase_end < today:
                status, progress = "Completed", 100
            elif current_start > today:
                status, progress = "Not Started", 0
            else:
                status, progress = "In Progress", random.randint(20, 90)
            
            phases.append({
                "phase_id": f"PH-{proj['project_id']}-{PHASE_NAMES.index((phase_name, weight))+1:02d}",
                "project_id": proj['project_id'],
                "phase_name": phase_name,
                "planned_start": current_start,
                "planned_end": phase_end,
                "actual_start": current_start if status != "Not Started" else None,
                "actual_end": phase_end if status == "Completed" else None,
                "status": status,
                "progress_pct": progress,
                "budget_aed": int(proj['contract_value_aed'] * weight)
            })
            current_start = phase_end
    return pd.DataFrame(phases)

def generate_work_packages(projects_df):
    packages = []
    boq_items = [
        ("Excavation", "CUM", 50, 150), ("Backfilling", "CUM", 30, 80),
        ("Lean Concrete", "CUM", 400, 600), ("Reinforcement Steel", "TON", 3500, 5000),
        ("Structural Concrete", "CUM", 800, 1200), ("Formwork", "SQM", 80, 150),
        ("Waterproofing", "SQM", 40, 100), ("Blockwork", "SQM", 120, 200),
        ("Plastering", "SQM", 25, 50), ("Tiling", "SQM", 80, 250),
        ("Painting", "SQM", 15, 40), ("MEP First Fix", "LS", 100000, 500000),
        ("MEP Second Fix", "LS", 80000, 400000), ("Fire Fighting", "LS", 50000, 200000),
        ("HVAC Installation", "LS", 100000, 600000), ("Facade Cladding", "SQM", 400, 1200)
    ]
    pkg_id = 1
    for _, proj in projects_df.iterrows():
        num_items = random.randint(8, 15)
        selected_items = random.sample(boq_items, num_items)
        for item_name, unit, min_rate, max_rate in selected_items:
            qty = random.randint(100, 10000)
            rate = random.randint(min_rate, max_rate)
            packages.append({
                "package_id": f"WP-{pkg_id:05d}",
                "project_id": proj['project_id'],
                "item_description": item_name,
                "unit": unit,
                "quantity": qty,
                "unit_rate_aed": rate,
                "total_amount_aed": qty * rate,
                "executed_qty": int(qty * proj['completion_pct'] / 100),
                "executed_amount_aed": int(qty * rate * proj['completion_pct'] / 100)
            })
            pkg_id += 1
    return pd.DataFrame(packages)

def generate_permits(projects_df):
    permits = []
    for _, proj in projects_df.iterrows():
        num_permits = random.randint(4, 8)
        selected = random.sample(PERMIT_TYPES, num_permits)
        for permit_name, prefix in selected:
            issue_date = random_date(proj['start_date'] - datetime.timedelta(days=180), proj['start_date'])
            permits.append({
                "permit_id": f"{prefix}-{proj['project_id'][-3:]}-{random.randint(1000, 9999)}",
                "project_id": proj['project_id'],
                "permit_type": permit_name,
                "reference_number": f"{prefix}-{random.randint(100000, 999999)}",
                "application_date": issue_date - datetime.timedelta(days=random.randint(30, 90)),
                "issue_date": issue_date,
                "expiry_date": issue_date + datetime.timedelta(days=random.choice([365, 730, 1095])),
                "status": random.choice(["Approved", "Approved", "Approved", "Pending", "Renewal Required"]),
                "fees_aed": random.randint(1000, 50000),
                "remarks": None
            })
    return pd.DataFrame(permits)

def generate_inspections(projects_df, employees_df):
    inspections = []
    inspection_types = ["Concrete Pour", "Rebar Placement", "Formwork", "MEP Rough-In", 
                        "Waterproofing", "Fire Safety", "Structural", "Final Handover"]
    inspectors = employees_df[employees_df['department'].isin(['QA/QC', 'HSE'])]['employee_id'].tolist()
    if not inspectors: inspectors = employees_df['employee_id'].tolist()[:10]
    
    for _, proj in projects_df.iterrows():
        if proj['status'] in ['Awarded', 'Not Started']: continue
        num_inspections = random.randint(20, 60)
        for _ in range(num_inspections):
            insp_date = random_date(proj['start_date'], min(proj['end_date'], datetime.date.today()))
            result = random.choice(["Passed", "Passed", "Passed", "Passed", "Failed", "Conditional"])
            inspections.append({
                "inspection_id": f"INS-{fake.uuid4()[:8].upper()}",
                "project_id": proj['project_id'],
                "inspection_type": random.choice(inspection_types),
                "inspection_date": insp_date,
                "inspector_id": random.choice(inspectors),
                "location": f"Level {random.randint(0, 50)}, Zone {random.choice(['A','B','C','D'])}",
                "result": result,
                "defects_found": random.randint(0, 5) if result != "Passed" else 0,
                "corrective_action": fake.sentence() if result != "Passed" else None,
                "reinspection_date": insp_date + datetime.timedelta(days=random.randint(3, 14)) if result == "Failed" else None,
                "remarks": fake.sentence() if random.random() > 0.7 else None
            })
    return pd.DataFrame(inspections)

def generate_safety_incidents(projects_df, employees_df):
    incidents = []
    incident_types = ["Near Miss", "First Aid", "Medical Treatment", "Lost Time Injury", 
                      "Property Damage", "Environmental Spill", "Fire Incident"]
    
    for _, proj in projects_df.iterrows():
        if proj['status'] in ['Awarded', 'Not Started']: continue
        num_incidents = random.randint(0, 15)
        for _ in range(num_incidents):
            inc_type = random.choices(incident_types, weights=[40, 25, 15, 5, 8, 5, 2])[0]
            inc_date = random_date(proj['start_date'], min(proj['end_date'], datetime.date.today()))
            incidents.append({
                "incident_id": f"INC-{fake.uuid4()[:8].upper()}",
                "project_id": proj['project_id'],
                "incident_date": inc_date,
                "incident_type": inc_type,
                "severity": "High" if inc_type in ["Lost Time Injury", "Fire Incident"] else random.choice(["Low", "Medium"]),
                "description": fake.sentence(),
                "location": f"Level {random.randint(0, 50)}, {random.choice(['North', 'South', 'East', 'West'])} Wing",
                "injured_persons": 1 if inc_type in ["First Aid", "Medical Treatment", "Lost Time Injury"] else 0,
                "lti_days": random.randint(1, 30) if inc_type == "Lost Time Injury" else 0,
                "root_cause": random.choice(["Unsafe Act", "Unsafe Condition", "Lack of Training", "Equipment Failure"]),
                "corrective_action": fake.sentence(),
                "status": random.choice(["Closed", "Closed", "Open"])
            })
    return pd.DataFrame(incidents)

def generate_payment_applications(projects_df):
    payments = []
    for _, proj in projects_df.iterrows():
        if proj['status'] == 'Awarded': continue
        num_ipcs = max(1, proj['completion_pct'] // 10)
        cumulative = 0
        for ipc_num in range(1, num_ipcs + 1):
            gross = int(proj['contract_value_aed'] * 0.1)
            cumulative += gross
            retention = int(gross * 0.10)
            deductions = random.randint(0, int(gross * 0.05))
            
            submit_date = proj['start_date'] + datetime.timedelta(days=30 * ipc_num)
            payments.append({
                "ipc_id": f"IPC-{proj['project_id']}-{ipc_num:03d}",
                "project_id": proj['project_id'],
                "ipc_number": ipc_num,
                "period_from": submit_date - datetime.timedelta(days=30),
                "period_to": submit_date,
                "submission_date": submit_date,
                "gross_value_aed": gross,
                "cumulative_value_aed": cumulative,
                "retention_aed": retention,
                "deductions_aed": deductions,
                "net_certified_aed": gross - retention - deductions,
                "status": random.choice(["Submitted", "Under Review", "Certified", "Certified", "Paid", "Paid"]),
                "payment_date": submit_date + datetime.timedelta(days=random.randint(30, 60)) if random.random() > 0.3 else None
            })
    return pd.DataFrame(payments)

def generate_variation_orders(projects_df):
    variations = []
    vo_reasons = ["Client Request", "Design Change", "Site Condition", "Authority Requirement", 
                  "Scope Addition", "Material Substitution"]
    
    for _, proj in projects_df.iterrows():
        if proj['status'] == 'Awarded': continue
        num_vos = random.randint(0, 10)
        for vo_num in range(1, num_vos + 1):
            vo_value = int(proj['contract_value_aed'] * random.uniform(0.005, 0.05))
            time_impact = random.choice([0, 0, 0, 7, 14, 30, 45])
            variations.append({
                "vo_id": f"VO-{proj['project_id']}-{vo_num:03d}",
                "project_id": proj['project_id'],
                "vo_number": vo_num,
                "description": fake.sentence(),
                "reason": random.choice(vo_reasons),
                "submitted_date": random_date(proj['start_date'], min(proj['end_date'], datetime.date.today())),
                "vo_value_aed": vo_value if random.random() > 0.3 else -vo_value,
                "time_impact_days": time_impact,
                "status": random.choice(["Draft", "Submitted", "Under Negotiation", "Approved", "Approved", "Rejected"]),
                "approved_date": None,
                "approved_by": fake.name() if random.random() > 0.5 else None
            })
    return pd.DataFrame(variations)

def generate_purchase_orders(projects_df, suppliers_df, contractors_df):
    pos = []
    for i, proj in projects_df.iterrows():
        if proj['status'] == 'Awarded': continue
        num_pos = random.randint(15, 40)
        for j in range(num_pos):
            is_material = random.random() > 0.3
            pos.append({
                "po_id": f"LPO-{proj['project_id']}-{j+1:04d}",
                "project_id": proj['project_id'],
                "po_type": "Material" if is_material else "Subcontract",
                "vendor_id": random.choice(suppliers_df['supplier_id'].tolist()) if is_material else random.choice(contractors_df['contractor_id'].tolist()),
                "description": fake.sentence(nb_words=6),
                "issue_date": random_date(proj['start_date'], min(proj['end_date'], datetime.date.today())),
                "delivery_date": random_date(proj['start_date'], proj['end_date']),
                "amount_aed": random.randint(5000, 500000),
                "status": random.choice(["Draft", "Issued", "Delivered", "Delivered", "Invoiced", "Paid"]),
                "approved_by": fake.name()
            })
    return pd.DataFrame(pos)

def generate_daily_reports(projects_df, employees_df):
    reports = []
    weather_conditions = ["Clear", "Clear", "Sunny", "Sunny", "Hot", "Humid", "Windy", "Sandstorm", "Rain"]
    site_engineers = employees_df[employees_df['designation'].str.contains('Site|Engineer', na=False)]['employee_id'].tolist()
    if not site_engineers: site_engineers = employees_df['employee_id'].tolist()[:20]
    
    for _, proj in projects_df.iterrows():
        if proj['status'] in ['Awarded', 'Completed']: continue
        # Generate last 90 days of reports
        end_date = min(datetime.date.today(), proj['end_date'])
        start_date = max(proj['start_date'], end_date - datetime.timedelta(days=90))
        
        current = start_date
        while current <= end_date:
            if current.weekday() != 4:  # Skip Fridays
                reports.append({
                    "report_id": f"DSR-{proj['project_id']}-{current.strftime('%Y%m%d')}",
                    "project_id": proj['project_id'],
                    "report_date": current,
                    "prepared_by": random.choice(site_engineers),
                    "weather": random.choice(weather_conditions),
                    "temperature_c": random.randint(25, 48),
                    "manpower_direct": random.randint(50, 300),
                    "manpower_subcon": random.randint(100, 500),
                    "equipment_on_site": random.randint(10, 50),
                    "work_description": fake.sentence(),
                    "issues": fake.sentence() if random.random() > 0.7 else None,
                    "visitors": random.randint(0, 10)
                })
            current += datetime.timedelta(days=1)
    return pd.DataFrame(reports)

def generate_documents(projects_df, employees_df):
    documents = []
    doc_controllers = employees_df[employees_df['designation'].str.contains('Document|Admin', na=False)]['employee_id'].tolist()
    if not doc_controllers: doc_controllers = employees_df['employee_id'].tolist()[:10]
    
    for _, proj in projects_df.iterrows():
        num_docs = random.randint(80, 200)
        for _ in range(num_docs):
            doc_type = random.choice(DOCUMENT_TYPES)
            doc_date = random_date(proj['start_date'], min(proj['end_date'], datetime.date.today()))
            documents.append({
                "document_id": f"DOC-{fake.uuid4()[:8].upper()}",
                "project_id": proj['project_id'],
                "document_type": doc_type,
                "document_number": f"{doc_type[:3].upper()}-{proj['project_id'][-3:]}-{random.randint(100, 999)}",
                "title": fake.sentence(nb_words=5).replace(".", ""),
                "revision": f"R{random.choice(['0', '0', '0', '1', '1', '2', '3'])}",
                "created_date": doc_date,
                "created_by": random.choice(doc_controllers),
                "status": random.choice(["Draft", "For Review", "Approved", "Approved", "Superseded"]),
                "discipline": random.choice(["Civil", "Structural", "Architectural", "MEP", "HSE"]),
                "file_path": f"//server/projects/{proj['project_id']}/{doc_type}/{fake.lexify('??????')}.pdf"
            })
    return pd.DataFrame(documents)

# === MAIN EXECUTION ===
def main():
    output_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(output_dir)
    
    print("=" * 60)
    print("DUBAI CIVIL ENGINEERING - DATA GENERATOR")
    print("=" * 60)
    
    print("\n[1/17] Generating Clients...")
    clients_df = generate_clients()
    clients_df.to_csv('clients.csv', index=False)
    print(f"       ✓ {len(clients_df)} clients")
    
    print("[2/17] Generating Employees...")
    employees_df = generate_employees()
    employees_df.to_csv('employees.csv', index=False)
    print(f"       ✓ {len(employees_df)} employees")
    
    print("[3/17] Generating Contractors...")
    contractors_df = generate_contractors()
    contractors_df.to_csv('contractors.csv', index=False)
    print(f"       ✓ {len(contractors_df)} contractors")
    
    print("[4/17] Generating Suppliers...")
    suppliers_df = generate_suppliers()
    suppliers_df.to_csv('suppliers.csv', index=False)
    print(f"       ✓ {len(suppliers_df)} suppliers")
    
    print("[5/17] Generating Equipment...")
    equipment_df = generate_equipment()
    equipment_df.to_csv('equipment.csv', index=False)
    print(f"       ✓ {len(equipment_df)} equipment items")
    
    print("[6/17] Generating Projects...")
    projects_df = generate_projects(clients_df, employees_df)
    projects_df.to_csv('projects.csv', index=False)
    print(f"       ✓ {len(projects_df)} projects")
    
    print("[7/17] Generating Contracts...")
    contracts_df = generate_contracts(projects_df, contractors_df)
    contracts_df.to_csv('contracts.csv', index=False)
    print(f"       ✓ {len(contracts_df)} contracts")
    
    print("[8/17] Generating Project Phases...")
    phases_df = generate_project_phases(projects_df)
    phases_df.to_csv('project_phases.csv', index=False)
    print(f"       ✓ {len(phases_df)} phases")
    
    print("[9/17] Generating Work Packages...")
    packages_df = generate_work_packages(projects_df)
    packages_df.to_csv('work_packages.csv', index=False)
    print(f"       ✓ {len(packages_df)} work packages")
    
    print("[10/17] Generating Permits & Approvals...")
    permits_df = generate_permits(projects_df)
    permits_df.to_csv('permits_approvals.csv', index=False)
    print(f"       ✓ {len(permits_df)} permits")
    
    print("[11/17] Generating Inspections...")
    inspections_df = generate_inspections(projects_df, employees_df)
    inspections_df.to_csv('inspections.csv', index=False)
    print(f"       ✓ {len(inspections_df)} inspections")
    
    print("[12/17] Generating Safety Incidents...")
    incidents_df = generate_safety_incidents(projects_df, employees_df)
    incidents_df.to_csv('safety_incidents.csv', index=False)
    print(f"       ✓ {len(incidents_df)} incidents")
    
    print("[13/17] Generating Payment Applications...")
    payments_df = generate_payment_applications(projects_df)
    payments_df.to_csv('payment_applications.csv', index=False)
    print(f"       ✓ {len(payments_df)} IPCs")
    
    print("[14/17] Generating Variation Orders...")
    variations_df = generate_variation_orders(projects_df)
    variations_df.to_csv('variation_orders.csv', index=False)
    print(f"       ✓ {len(variations_df)} VOs")
    
    print("[15/17] Generating Purchase Orders...")
    pos_df = generate_purchase_orders(projects_df, suppliers_df, contractors_df)
    pos_df.to_csv('purchase_orders.csv', index=False)
    print(f"       ✓ {len(pos_df)} POs")
    
    print("[16/17] Generating Daily Site Reports...")
    reports_df = generate_daily_reports(projects_df, employees_df)
    reports_df.to_csv('daily_site_reports.csv', index=False)
    print(f"       ✓ {len(reports_df)} reports")
    
    print("[17/17] Generating Project Documents...")
    documents_df = generate_documents(projects_df, employees_df)
    documents_df.to_csv('project_documents.csv', index=False)
    print(f"       ✓ {len(documents_df)} documents")
    
    print("\n" + "=" * 60)
    print("DATA GENERATION COMPLETE!")
    print(f"Files saved to: {output_dir}")
    print("=" * 60)

if __name__ == "__main__":
    main()
