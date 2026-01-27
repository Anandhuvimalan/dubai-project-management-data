
import pandas as pd
import random
from faker import Faker
import datetime
import os

fake = Faker()
Faker.seed(42)  # For reproducibility

# --- Configuration ---
NUM_PROJECTS = 50
NUM_CLIENTS = 25
NUM_EMPLOYEES = 250
NUM_TASKS_PER_PROJECT = (10, 30)
NUM_VENDORS = 40
NUM_POS = 300
NUM_DOCUMENTS = 600
START_DATE = datetime.date(2023, 1, 1)
END_DATE = datetime.date(2026, 12, 31)

# --- Dubai Specific Context ---
DUBAI_LOCATIONS = [
    "Downtown Dubai", "Business Bay", "Dubai Marina", "DIFC", "Jumeirah Lake Towers (JLT)",
    "Palm Jumeirah", "Deira", "Bur Dubai", "Dubai Hills Estate", "Al Quoz Industrial Area",
    "Dubai Silicon Oasis", "Jebel Ali Free Zone (JAFZA)", "Expo City Dubai", "Sheikh Zayed Road",
    "Meydan", "City Walk", "Dubai Creek Harbour", "Al Barsha"
]

PROJECT_TYPES = [
    "Construction - Residential Tower", "Construction - Commercial Mall", "Infrastructure - Road Network",
    "IT - Government Smart Services Portal", "IT - ERP Implementation", "Events - Global Conference Setup",
    "Energy - Solar Park Installation", "Consulting - Strategic Vision 2030", "Marketing - Tourism Campaign",
    "Healthcare - Hospital Wing Expansion", "Logistics - Warehouse Automation"
]

CLIENT_TYPES = [
    "Real Estate Developer", "Government Authority", "Multinational Corp", "Retail Giant", "Logistics Firm", "Hospitality Group"
]

DEPARTMENTS = [
    "Project Management Office (PMO)", "Engineering", "Procurement", "Information Technology",
    "Human Resources", "Finance", "Legal", "Sales & Marketing", "Operations", "Quality Assurance"
]

NATIONALITIES = [
    "UAE", "India", "Philippines", "UK", "Pakistan", "Egypt", "Lebanon", "USA", "Canada", "Jordan", 
    "Syria", "Australia", "Germany", "South Africa", "France"
]

MILESTONE_STAGES = [
    "Initiation & Feasibility", "Concept Design", "Detailed Design", "Tendering & Procurement", 
    "NOCs & Approvals", "Construction / Implementation", "Testing & Commissioning", "Handover & Closeout"
]

DOCUMENT_TYPES = [
    "MOU", "NDA", "Contract Agreement", "NOC (No Objection Certificate)", "Blueprint / Schematic", 
    "Variation Order", "Safety Report", "Inspection Request", "Invoice", "Meeting Minutes"
]

VENDOR_CATEGORIES = [
    "Construction Materials", "Heavy Machinery", "IT Software & Licensing", "IT Hardware", 
    "Manpower Supply", "Catering Services", "Logistics & Transport", "Consultancy Services"
]

# --- Helper Functions ---

def random_date(start, end):
    if start >= end: return start
    return fake.date_between(start_date=start, end_date=end)

# --- Data Generation Functions ---

def generate_clients():
    clients = []
    prefixes = ["Al", "Dubai", "Emirates", "Gulf", "Middle East", "Royal", "Future", "Arabian", "Desert"]
    suffixes = ["Holdings", "Group", "Investments", "Properties", "Solutions", "Enterprises", "Authority", "Development"]

    for i in range(1, NUM_CLIENTS + 1):
        name = f"{random.choice(prefixes)} {fake.word().capitalize()} {random.choice(suffixes)}"
        clients.append({
            "client_id": f"CL-{i:03d}",
            "client_name": name,
            "industry": random.choice(CLIENT_TYPES),
            "contact_person": fake.name(),
            "email": fake.company_email(),
            "phone": "+971-5" + str(random.randint(0, 9)) + "-" + str(random.randint(1000000, 9999999)),
            "location": random.choice(DUBAI_LOCATIONS),
            "payment_terms": random.choice(["Net 30", "Net 60", "Net 90"]),
            "is_vip": random.choice([True, False, False])
        })
    return pd.DataFrame(clients)

def generate_employees():
    employees = []
    roles = {
        "Project Management Office (PMO)": ["Project Manager", "Program Manager", "Project Coordinator", "Scheduler"],
        "Engineering": ["Civil Engineer", "MEP Engineer", "Architect", "Site Engineer"],
        "Information Technology": ["Software Developer", "System Analyst", "IT Support", "Data Scientist"],
        "Finance": ["Accountant", "Financial Analyst", "Payroll Officer"],
        "Human Resources": ["HR Manager", "Recruiter", "PRO (Public Relations Officer)"],
        "Procurement": ["Procurement Specialist", "Buyer", "Vendor Manager"],
        "Legal": ["Legal Counsel", "Contract Specialist"],
        "Sales & Marketing": ["Sales Executive", "Marketing Manager", "Digital Specialist"]
    }
    
    for i in range(1, NUM_EMPLOYEES + 1):
        dept = random.choice(DEPARTMENTS)
        role = random.choice(roles.get(dept, ["Specialist", "Officer", "Manager"]))
        employees.append({
            "employee_id": f"EMP-{i:04d}",
            "full_name": fake.name(),
            "department": dept,
            "role": role,
            "nationality": random.choice(NATIONALITIES),
            "joining_date": random_date(datetime.date(2020, 1, 1), START_DATE),
            "salary_aed": random.randint(4000, 65000),
            "visa_status": random.choice(["Employment Visa", "Golden Visa", "Mission Visa", "Partner Visa"]),
            "labor_card_number": fake.numerify("##########")
        })
    return pd.DataFrame(employees)

def generate_projects(clients_df, employees_df):
    projects = []
    managers = employees_df[employees_df['department'] == 'Project Management Office (PMO)']['employee_id'].tolist()
    if not managers: managers = employees_df['employee_id'].tolist()

    for i in range(1, NUM_PROJECTS + 1):
        start = random_date(START_DATE, datetime.date(2025, 12, 31))
        duration = random.randint(6, 48) # months
        end = start + datetime.timedelta(days=duration*30)
        
        today = datetime.date.today()
        if end < today:
            status = "Completed"
            completion_pct = 100
        elif start > today:
            status = "Planned"
            completion_pct = 0
        else:
            status = random.choice(["In Progress", "In Progress", "In Progress", "On Hold", "Delayed"])
            completion_pct = random.randint(10, 95)

        budget = random.randint(500000, 150000000)

        projects.append({
            "project_id": f"PRJ-{i:03d}",
            "project_name": fake.bs().title(), # More professional sounding titles
            "type": random.choice(PROJECT_TYPES),
            "client_id": random.choice(clients_df['client_id']),
            "manager_id": random.choice(managers),
            "start_date": start,
            "end_date": end,
            "status": status,
            "budget_aed": budget,
            "location": random.choice(DUBAI_LOCATIONS),
            "priority": random.choice(["High", "Medium", "Low", "Critical"]),
            "completion_percentage": completion_pct,
            "description": fake.catch_phrase()
        })
    return pd.DataFrame(projects)

def generate_milestones(projects_df):
    milestones = []
    for _, project in projects_df.iterrows():
        start = project['start_date']
        end = project['end_date']
        total_days = (end - start).days
        
        current_date_pointer = start
        for stage in MILESTONE_STAGES:
            duration = int(total_days / len(MILESTONE_STAGES)) # Simple distribution
            m_end = current_date_pointer + datetime.timedelta(days=duration)
            
            status = "Pending"
            today = datetime.date.today()
            if m_end < today: status = "Approved"
            elif current_date_pointer < today < m_end: status = "In Progress"
            
            milestones.append({
                "milestone_id": fake.uuid4()[:8],
                "project_id": project['project_id'],
                "milestone_name": stage,
                "planned_start": current_date_pointer,
                "planned_end": m_end,
                "status": status,
                "sign_off_by": fake.name() if status == "Approved" else None
            })
            current_date_pointer = m_end
    return pd.DataFrame(milestones)

def generate_tasks(projects_df, employees_df, milestones_df):
    tasks = []
    for _, project in projects_df.iterrows():
        # Get milestones for this project to link tasks
        proj_milestones = milestones_df[milestones_df['project_id'] == project['project_id']]
        
        num_tasks = random.randint(*NUM_TASKS_PER_PROJECT)
        for j in range(num_tasks):
            # Link to a random milestone
            if not proj_milestones.empty:
                milestone = proj_milestones.sample(1).iloc[0]
                m_start = milestone['planned_start']
                m_end = milestone['planned_end']
            else:
                m_start, m_end = project['start_date'], project['end_date']

            start = random_date(m_start, m_end - datetime.timedelta(days=1))
            end = random_date(start, m_end)
            
            status = "Completed" if end < datetime.date.today() else random.choice(["Not Started", "In Progress", "Blocked", "Review"])
            
            tasks.append({
                "task_id": f"TSK-{project['project_id']}-{j+1:03d}",
                "project_id": project['project_id'],
                "milestone_id": milestone['milestone_id'] if not proj_milestones.empty else None,
                "task_name": fake.sentence(nb_words=5).replace(".", ""),
                "description": fake.text(max_nb_chars=50),
                "assigned_to": random.choice(employees_df['employee_id']),
                "start_date": start,
                "end_date": end,
                "status": status,
                "priority": random.choice(["High", "Medium", "Low"]),
                "estimated_hours": random.choice([4, 8, 16, 24, 40, 80])
            })
    return pd.DataFrame(tasks)

def generate_vendors():
    vendors = []
    for i in range(1, NUM_VENDORS + 1):
        vendors.append({
            "vendor_id": f"VEN-{i:03d}",
            "vendor_name": fake.company(),
            "category": random.choice(VENDOR_CATEGORIES),
            "contact_person": fake.name(),
            "email": fake.company_email(),
            "phone": "+971-4-" + str(random.randint(1000000, 9999999)),
            "location": random.choice(["Al Quoz", "Jebel Ali", "Ras Al Khor", "Dubai Investment Park", "International City"]),
            "trn_number": "100" + fake.numerify("##########") + "9" # Tax Registration Number format
        })
    return pd.DataFrame(vendors)

def generate_purchase_orders(projects_df, vendors_df):
    pos = []
    for i in range(1, NUM_POS + 1):
        project = projects_df.sample(1).iloc[0]
        vendor = vendors_df.sample(1).iloc[0]
        date = random_date(project['start_date'], min(project['end_date'], datetime.date.today()))
        
        amount = random.uniform(5000, 250000)
        
        pos.append({
            "po_id": f"PO-{i:05d}",
            "project_id": project['project_id'],
            "vendor_id": vendor['vendor_id'],
            "issue_date": date,
            "status": random.choice(["Draft", "Issued", "Delivered", "Invoiced", "Paid"]),
            "amount_aed": round(amount, 2),
            "items_description": fake.sentence(nb_words=6),
            "approved_by": fake.name()
        })
    return pd.DataFrame(pos)

def generate_documents(projects_df):
    docs = []
    for _ in range(NUM_DOCUMENTS):
        project = projects_df.sample(1).iloc[0]
        date = random_date(project['start_date'], min(project['end_date'], datetime.date.today()))
        
        docs.append({
            "document_id": fake.uuid4()[:8],
            "project_id": project['project_id'],
            "type": random.choice(DOCUMENT_TYPES),
            "title": fake.sentence(nb_words=4).replace(".", ""),
            "created_date": date,
            "version": f"v{random.randint(1, 5)}.0",
            "status": random.choice(["Draft", "Under Review", "Approved", "Obsolete"]),
            "link": f"https://sharepoint.company.com/docs/{fake.lexify(text='??????')}.pdf"
        })
    return pd.DataFrame(docs)

def generate_timesheets(tasks_df):
    timesheets = []
    # Generate timesheets for a subset of tasks to avoid massive file size
    active_tasks = tasks_df.sample(frac=0.3) 
    
    for _, task in active_tasks.iterrows():
        # Generate 1-5 entries per task
        num_entries = random.randint(1, 5)
        for _ in range(num_entries):
            date = random_date(task['start_date'], task['end_date'])
            hours = random.choice([1, 2, 4, 8])
            
            timesheets.append({
                "entry_id": fake.uuid4()[:8],
                "task_id": task['task_id'],
                "employee_id": task['assigned_to'], 
                "date": date,
                "hours_logged": hours,
                "description": fake.sentence(nb_words=5),
                "is_billable": random.choice([True, True, False]),
                "status": "Approved"
            })
    return pd.DataFrame(timesheets)

def generate_assignments(tasks_df, employees_df):
    assignments = []
    
    for _, task in tasks_df.iterrows():
        # 1. Primary Assignee (Lead)
        assignments.append({
            "assignment_id": fake.uuid4()[:8],
            "task_id": task['task_id'],
            "employee_id": task['assigned_to'],
            "role": "Lead",
            "start_date": task['start_date'],
            "end_date": task['end_date'],
            "allocated_hours": task['estimated_hours']
        })
        
        # 2. Additional Resources (Support/Reviewer)
        # Randomly assign 0-3 extra people
        num_additional = random.randint(0, 3)
        potential_assignees = employees_df[employees_df['employee_id'] != task['assigned_to']]
        
        if not potential_assignees.empty and num_additional > 0:
            additional_peeps = potential_assignees.sample(min(num_additional, len(potential_assignees)))
            for _, emp in additional_peeps.iterrows():
                role = random.choice(["Support", "Reviewer", "QA"])
                # Allocate partial hours for support
                hours = int(task['estimated_hours'] * random.uniform(0.1, 0.5))
                if hours < 1: hours = 1
                
                assignments.append({
                    "assignment_id": fake.uuid4()[:8],
                    "task_id": task['task_id'],
                    "employee_id": emp['employee_id'],
                    "role": role,
                    "start_date": task['start_date'],
                    "end_date": task['end_date'],
                    "allocated_hours": hours
                })
                
    return pd.DataFrame(assignments)

# --- Execution ---

def main():
    print("Generating Clients...")
    clients_df = generate_clients()
    clients_df.to_csv('clients.csv', index=False)
    
    print("Generating Employees...")
    employees_df = generate_employees()
    employees_df.to_csv('employees.csv', index=False)
    
    print("Generating Vendors...")
    vendors_df = generate_vendors()
    vendors_df.to_csv('vendors.csv', index=False)
    
    print("Generating Projects...")
    projects_df = generate_projects(clients_df, employees_df)
    projects_df.to_csv('projects.csv', index=False)
    
    print("Generating Milestones...")
    milestones_df = generate_milestones(projects_df)
    milestones_df.to_csv('project_milestones.csv', index=False)
    
    print("Generating Tasks...")
    tasks_df = generate_tasks(projects_df, employees_df, milestones_df)
    tasks_df.to_csv('tasks.csv', index=False)
    
    print("Generating Purchase Orders...")
    pos_df = generate_purchase_orders(projects_df, vendors_df)
    pos_df.to_csv('purchase_orders.csv', index=False)
    
    print("Generating Documents...")
    docs_df = generate_documents(projects_df)
    docs_df.to_csv('project_documents.csv', index=False)
    
    print("Generating Timesheets...")
    timesheets_df = generate_timesheets(tasks_df)
    timesheets_df.to_csv('timesheets.csv', index=False)

    print("Generating Assignments...")
    assignments_df = generate_assignments(tasks_df, employees_df)
    assignments_df.to_csv('assignments.csv', index=False)
    
    # Re-generate Risks and Expenses with new project context (Simplified inline)
    print("Generating Risks & Expenses...")
    # ... logic similar to previous but using new project list if needed. 
    # For brevity, reusing the simple structure logic inline here would be good but I'll stick to the structure
    # Let's add them back quickly to keep the set complete
    
    expenses = []
    for _ in range(800):
        project = projects_df.sample(1).iloc[0]
        expenses.append({
            "expense_id": fake.uuid4()[:8],
            "project_id": project['project_id'],
            "date": random_date(project['start_date'], min(project['end_date'], datetime.date.today())),
            "category": random.choice(["Travel", "Materials", "Software", "Entertainment", "Site Office"]),
            "amount_aed": round(random.uniform(50, 5000), 2)
        })
    pd.DataFrame(expenses).to_csv('expenses.csv', index=False)
    
    risks = []
    for _ in range(200):
        project = projects_df.sample(1).iloc[0]
        risks.append({
            "risk_id": fake.uuid4()[:8],
            "project_id": project['project_id'],
            "description": fake.sentence(),
            "impact": random.choice(["Critical", "High", "Medium", "Low"]),
            "status": "Active"
        })
    pd.DataFrame(risks).to_csv('risks.csv', index=False)

    print(f"Data generation complete! Files saved to {os.getcwd()}")

if __name__ == "__main__":
    main()
