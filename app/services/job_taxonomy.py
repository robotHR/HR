"""
Taxonomie joburi NEXAS HR.

Scop: motor de matching HR stabil, fara scoruri absurde intre familii diferite.
Regula: intai familia si subfamilia jobului, apoi scorul AI.
Suporta titluri RO si EN.
"""

from __future__ import annotations

JOB_FAMILIES = {
    "IT_SOFTWARE": {
        "label": "IT Software",
        "subfamilies": {
            "software_development": ["software developer", "software engineer", "programator", "dezvoltator software", "dezvoltator it", "developer", "application developer"],
            "backend": ["backend", "back end", "api developer", "python developer", "java developer", "php developer", "node developer", "c# developer", ".net developer"],
            "frontend": ["frontend", "front end", "react developer", "vue developer", "angular developer", "javascript developer", "typescript developer"],
            "fullstack": ["full stack", "fullstack", "full-stack", "full stack developer"],
            "mobile": ["mobile developer", "android developer", "ios developer", "flutter developer", "react native developer"],
            "qa_automation": ["qa automation", "automation tester", "test automation", "sdet"],
            "web_development": ["web developer", "web programmer", "dezvoltator web", "wordpress developer"],
            "database_development": ["sql developer", "database developer", "pl sql developer", "oracle developer"],
            "embedded": ["embedded developer", "embedded software", "firmware developer"],
            "game_development": ["game developer", "unity developer", "unreal developer"]
        },
        "hard_signals": ["python", "javascript", "typescript", "java", "c#", "php", "react", "angular", "vue", "node", "django", "flask", "spring", "sql", "git", "api", "backend", "frontend", "full stack", "programare", "software", "cod", "aplicatii"]
    },
    "IT_INFRASTRUCTURE": {
        "label": "IT Infrastructure",
        "subfamilies": {
            "system_admin": ["system administrator", "sysadmin", "administrator sistem", "administrator servere"],
            "network": ["network engineer", "administrator retea", "network administrator", "retelistica"],
            "devops": ["devops", "devops engineer", "site reliability", "sre"],
            "cloud": ["cloud engineer", "aws engineer", "azure engineer", "gcp engineer"],
            "linux": ["linux administrator", "linux engineer"],
            "windows": ["windows administrator", "active directory administrator"],
            "database_admin": ["dba", "database administrator", "administrator baze de date"],
            "noc": ["noc engineer", "noc technician"],
            "virtualization": ["vmware", "virtualization engineer"],
            "telecom": ["telecom engineer", "telecommunications engineer"]
        },
        "hard_signals": ["server", "linux", "windows server", "network", "retelistica", "aws", "azure", "docker", "kubernetes", "active directory", "vmware"]
    },
    "IT_SUPPORT": {
        "label": "IT Support",
        "subfamilies": {
            "helpdesk": ["helpdesk", "help desk", "it helpdesk"],
            "service_desk": ["service desk", "it service desk"],
            "technical_support": ["technical support", "suport tehnic", "support engineer"],
            "desktop_support": ["desktop support", "it support specialist"],
            "field_it": ["field it technician", "tehnician it teren"],
            "hardware": ["hardware technician", "tehnician hardware"],
            "application_support": ["application support", "suport aplicatii"],
            "erp_support": ["erp support", "sap support"],
            "pos_support": ["pos support", "suport pos"],
            "it_operator": ["operator it", "it operator"]
        },
        "hard_signals": ["suport tehnic", "ticket", "helpdesk", "service desk", "hardware", "it support", "windows", "troubleshooting"]
    },
    "DATA_BI": {
        "label": "Data & BI",
        "subfamilies": {
            "data_analyst": ["data analyst", "analist date", "analyst data"],
            "bi_developer": ["bi developer", "business intelligence", "power bi developer"],
            "data_engineer": ["data engineer", "inginer date"],
            "reporting": ["reporting analyst", "analist raportare"],
            "etl": ["etl developer", "etl specialist"],
            "database_reporting": ["sql analyst", "sql reporting"],
            "data_science": ["data scientist", "machine learning engineer"],
            "analytics": ["analytics specialist", "web analytics"],
            "data_entry": ["operator introducere date", "data entry", "operator validare date"],
            "master_data": ["master data specialist", "master data analyst"]
        },
        "hard_signals": ["excel", "sql", "power bi", "tableau", "raportare", "analiza date", "etl", "python", "dashboard"]
    },
    "CYBERSECURITY": {
        "label": "Cybersecurity",
        "subfamilies": {
            "soc": ["soc analyst", "security operations"],
            "security_engineer": ["security engineer", "cybersecurity engineer"],
            "penetration_testing": ["penetration tester", "ethical hacker"],
            "grc": ["grc analyst", "information security officer"],
            "incident_response": ["incident responder", "incident response"],
            "iam": ["iam specialist", "identity access"],
            "network_security": ["network security"],
            "cloud_security": ["cloud security"],
            "appsec": ["application security", "appsec"],
            "security_admin": ["security administrator"]
        },
        "hard_signals": ["cybersecurity", "security", "soc", "siem", "vulnerability", "penetration", "iso 27001"]
    },
    "FINANCE_ACCOUNTING": {
        "label": "Finance & Accounting",
        "subfamilies": {
            "accounting": ["contabil", "accountant", "contabilitate"],
            "chief_accountant": ["contabil sef", "chief accountant"],
            "financial_analysis": ["analist financiar", "financial analyst"],
            "economist": ["economist"],
            "audit": ["auditor", "audit financiar"],
            "tax": ["tax specialist", "specialist taxe"],
            "payroll": ["payroll", "salarizare"],
            "accounts_payable": ["accounts payable", "ap accountant"],
            "accounts_receivable": ["accounts receivable", "ar accountant"],
            "controller": ["financial controller", "controller financiar"]
        },
        "hard_signals": ["contabilitate", "balanta", "bilant", "saga", "ceccar", "facturi", "taxe", "raportare financiara", "excel"]
    },
    "BANKING_INSURANCE": {
        "label": "Banking & Insurance",
        "subfamilies": {
            "banking_officer": ["ofiter bancar", "banking officer"],
            "credit": ["analist credite", "credit analyst"],
            "insurance_agent": ["agent asigurari", "insurance agent"],
            "claims": ["claims specialist", "specialist daune"],
            "underwriting": ["underwriter", "subscriere"],
            "risk": ["risk analyst", "analist risc"],
            "branch": ["branch officer", "consilier bancar"],
            "leasing": ["leasing officer", "consultant leasing"],
            "compliance_banking": ["aml analyst", "compliance banking"],
            "brokerage": ["broker asigurari", "broker credite"]
        },
        "hard_signals": ["banca", "credit", "asigurari", "daune", "underwriting", "aml", "risk"]
    },
    "HR_RECRUITMENT": {
        "label": "HR & Recruitment",
        "subfamilies": {
            "recruitment": ["recruiter", "specialist recrutare", "recruitment specialist"],
            "hr_generalist": ["hr generalist", "specialist resurse umane", "resurse umane"],
            "hrbp": ["hr business partner", "hrbp"],
            "payroll_hr": ["inspector resurse umane", "inspector hr", "revisal"],
            "training": ["training specialist", "learning development", "l&d"],
            "comp_ben": ["compensation benefits", "comp & ben"],
            "talent": ["talent acquisition", "talent management"],
            "employee_relations": ["employee relations"],
            "hr_admin": ["hr admin", "administrator hr"],
            "organizational_development": ["organizational development", "od specialist"]
        },
        "hard_signals": ["recrutare", "interviuri", "resurse umane", "hr", "revisal", "talent", "employee"]
    },
    "LEGAL": {
        "label": "Legal",
        "subfamilies": {
            "legal_advisor": ["jurist", "legal advisor", "consilier juridic"],
            "lawyer": ["avocat", "lawyer", "attorney"],
            "paralegal": ["paralegal", "asistent juridic"],
            "compliance": ["compliance officer", "specialist conformitate"],
            "contracts": ["contract manager", "specialist contracte"],
            "notary": ["notar", "notary"],
            "gdpr": ["gdpr specialist", "data protection officer", "dpo"],
            "litigation": ["litigation", "contencios"],
            "public_procurement": ["achizitii publice juridic"],
            "legal_admin": ["legal assistant", "secretar juridic"]
        },
        "hard_signals": ["juridic", "contracte", "lege", "avocat", "litigii", "gdpr", "compliance"]
    },
    "MARKETING": {
        "label": "Marketing",
        "subfamilies": {
            "marketing_generalist": ["marketing specialist", "specialist marketing", "marketing coordinator"],
            "brand": ["brand manager", "brand specialist"],
            "trade_marketing": ["trade marketing"],
            "product_marketing": ["product marketing"],
            "market_research": ["market research", "cercetare piata"],
            "communications": ["communications specialist", "specialist comunicare"],
            "pr": ["pr specialist", "public relations"],
            "events": ["event marketing", "event specialist"],
            "copywriting": ["copywriter", "content writer"],
            "marketing_manager": ["marketing manager", "director marketing"]
        },
        "hard_signals": ["marketing", "brand", "campanii", "comunicare", "content", "pr", "evenimente"]
    },
    "DIGITAL_MARKETING": {
        "label": "Digital Marketing",
        "subfamilies": {
            "seo": ["seo specialist", "seo"],
            "ppc": ["ppc specialist", "google ads", "paid search"],
            "social_media": ["social media", "social media manager"],
            "performance": ["performance marketing", "growth marketing"],
            "email_marketing": ["email marketing", "crm marketing"],
            "ecommerce_marketing": ["ecommerce marketing", "e-commerce marketing"],
            "content_marketing": ["content marketing", "content creator"],
            "affiliate": ["affiliate marketing"],
            "marketing_automation": ["marketing automation"],
            "digital_manager": ["digital marketing manager"]
        },
        "hard_signals": ["seo", "ppc", "google ads", "facebook ads", "social media", "analytics", "campanii digitale", "content"]
    },
    "SALES_B2B": {
        "label": "Sales B2B",
        "subfamilies": {
            "sales_representative": ["agent vanzari", "reprezentant vanzari", "sales representative", "sales agent", "agent comercial"],
            "field_sales": ["field sales", "vanzari teren", "reprezentant teren"],
            "account_manager": ["account manager", "manager cont", "key account"],
            "business_development": ["business development", "bdm"],
            "inside_sales": ["inside sales", "telesales b2b"],
            "sales_engineer": ["sales engineer", "inginer vanzari"],
            "area_sales": ["area sales manager", "regional sales"],
            "sales_consultant": ["consultant vanzari", "sales consultant"],
            "export_sales": ["export sales", "international sales"],
            "sales_manager": ["sales manager", "director vanzari"]
        },
        "hard_signals": ["vanzari", "clienti", "portofoliu", "negociere", "ofertare", "b2b", "agent comercial", "sales"]
    },
    "SALES_RETAIL": {
        "label": "Sales Retail",
        "subfamilies": {
            "shop_assistant": ["lucrator comercial", "vanzator", "sales assistant", "shop assistant"],
            "cashier_sales": ["casier vanzator", "cashier sales"],
            "store_manager": ["store manager", "sef magazin", "manager magazin"],
            "department_manager": ["sef raion", "department manager"],
            "merchandiser": ["merchandiser", "mercantizor"],
            "promoter": ["promoter", "brand promoter"],
            "retail_consultant": ["consultant retail", "consultant magazin"],
            "visual_merchandising": ["visual merchandiser"],
            "retail_supervisor": ["retail supervisor"],
            "store_keeper": ["magazioner retail"]
        },
        "hard_signals": ["magazin", "retail", "clienti", "vanzare", "casa de marcat", "raft", "marfa"]
    },
    "CUSTOMER_SUPPORT": {
        "label": "Customer Support",
        "subfamilies": {
            "call_center": ["operator call center", "call center", "agent call center"],
            "customer_service": ["customer service", "relatii clienti", "suport clienti"],
            "customer_success": ["customer success", "customer success manager"],
            "backoffice": ["back office", "backoffice"],
            "technical_customer_support": ["technical customer support"],
            "chat_support": ["chat support", "live chat"],
            "complaints": ["complaints specialist", "reclamatii"],
            "dispatch_support": ["dispecer", "dispatcher"],
            "order_support": ["order support", "procesare comenzi"],
            "contact_center_supervisor": ["team leader call center", "contact center supervisor"]
        },
        "hard_signals": ["call center", "clienti", "suport", "relatii clienti", "telefon", "ticket", "reclamatii", "dispecer"]
    },
    "ADMINISTRATIVE": {
        "label": "Administrative",
        "subfamilies": {
            "office_assistant": ["office assistant", "asistent birou"],
            "secretary": ["secretara", "secretar", "secretary"],
            "reception_office": ["receptionist office", "receptioner birou"],
            "data_entry_admin": ["operator date", "operator introducere date", "data entry"],
            "document_controller": ["document controller", "arhivar"],
            "personal_assistant": ["personal assistant", "asistent manager"],
            "administrative_clerk": ["referent", "referent administrativ"],
            "registry": ["registratura"],
            "office_manager": ["office manager"],
            "front_desk": ["front desk office"]
        },
        "hard_signals": ["administrativ", "documente", "birou", "secretariat", "arhivare", "excel", "date"]
    },
    "PROCUREMENT": {
        "label": "Procurement & Purchasing",
        "subfamilies": {
            "buyer": ["buyer", "achizitor", "specialist achizitii"],
            "procurement_specialist": ["procurement specialist", "specialist procurement"],
            "purchasing": ["purchasing specialist", "aprovizionare"],
            "strategic_sourcing": ["strategic sourcing"],
            "vendor_management": ["vendor manager", "supplier manager"],
            "category_buyer": ["category buyer", "category manager procurement"],
            "import_purchasing": ["import specialist", "achizitii import"],
            "contract_procurement": ["contract procurement"],
            "materials_planner": ["materials planner"],
            "procurement_manager": ["procurement manager", "manager achizitii"]
        },
        "hard_signals": ["achizitii", "furnizori", "negociere", "aprovizionare", "purchase", "procurement"]
    },
    "SUPPLY_CHAIN": {
        "label": "Supply Chain",
        "subfamilies": {
            "planner": ["supply planner", "demand planner", "planner"],
            "inventory": ["inventory specialist", "control stocuri"],
            "warehouse_planning": ["warehouse planner"],
            "logistics_planning": ["logistics planner"],
            "import_export": ["import export specialist", "export import"],
            "customs": ["customs specialist", "specialist vamal"],
            "transport_planning": ["transport planner"],
            "supply_chain_analyst": ["supply chain analyst"],
            "operations_planning": ["operations planner"],
            "supply_chain_manager": ["supply chain manager"]
        },
        "hard_signals": ["supply chain", "stocuri", "planificare", "transport", "vama", "import", "export", "logistica"]
    },
    "LOGISTICS_WAREHOUSE": {
        "label": "Logistics & Warehouse",
        "subfamilies": {
            "warehouse_worker": ["manipulant marfa", "warehouse worker", "lucrator depozit", "muncitor depozit"],
            "picker": ["picker", "order picker", "pregatitor comenzi"],
            "packer": ["packer", "ambalator", "ambalare"],
            "forklift": ["stivuitorist", "forklift operator", "operator stivuitor"],
            "warehouse_clerk": ["gestionar depozit", "depozitar", "warehouse clerk", "magazioner"],
            "inventory_warehouse": ["inventarist", "inventory clerk"],
            "loading": ["incarcare descarcare", "loader", "unloader"],
            "warehouse_operator": ["operator depozit", "warehouse operator"],
            "warehouse_supervisor": ["sef depozit", "warehouse supervisor"],
            "logistics_operator": ["operator logistica", "logistics operator"]
        },
        "hard_signals": ["depozit", "marfa", "stivuitor", "picking", "packing", "gestionar", "incarcare", "descarcare", "wms", "iscir"]
    },
    "TRANSPORTATION": {
        "label": "Transportation",
        "subfamilies": {
            "driver_b": ["sofer categoria b", "sofer b", "driver b", "sofer distributie"],
            "driver_c": ["sofer categoria c", "sofer c", "truck driver c"],
            "driver_ce": ["sofer tir", "sofer ce", "tir", "truck driver", "hgv driver"],
            "courier": ["curier", "courier", "livrator", "delivery driver"],
            "bus_driver": ["sofer autobuz", "bus driver"],
            "taxi": ["taxi driver", "sofer taxi", "rideshare"],
            "dispatcher_transport": ["dispecer transport", "transport dispatcher"],
            "fleet": ["fleet manager", "administrator flota"],
            "transport_coordinator": ["transport coordinator", "coordonator transport"],
            "route_planner": ["route planner", "planificator rute"]
        },
        "hard_signals": ["sofer", "permis", "categoria b", "categoria c", "categoria ce", "tir", "livrari", "distributie", "curier"]
    },
    "PRODUCTION_MANUFACTURING": {
        "label": "Production & Manufacturing",
        "subfamilies": {
            "production_operator": ["operator productie", "production operator"],
            "machine_operator": ["operator masina", "machine operator"],
            "assembly": ["asamblor", "assembly operator"],
            "cnc": ["operator cnc", "cnc operator"],
            "quality_control": ["control calitate", "quality control", "qc inspector"],
            "line_leader": ["line leader", "sef linie"],
            "production_planner": ["planificator productie", "production planner"],
            "process_operator": ["operator proces", "process operator"],
            "manufacturing_engineer": ["inginer productie", "manufacturing engineer", "process engineer"],
            "production_manager": ["production manager", "manager productie"]
        },
        "hard_signals": ["productie", "linie", "masina", "asamblare", "cnc", "calitate", "proces", "fabricatie"]
    },
    "MAINTENANCE_TECHNICAL": {
        "label": "Maintenance & Technical Service",
        "subfamilies": {
            "maintenance_technician": ["tehnician mentenanta", "maintenance technician"],
            "mechanical_maintenance": ["mecanic mentenanta", "mechanical maintenance"],
            "electrical_maintenance": ["electrician mentenanta", "electrical maintenance"],
            "industrial_mechanic": ["lacatus mecanic", "mecanic industrial"],
            "automation": ["automatist", "automation technician"],
            "hvac": ["frigotehnist", "hvac technician"],
            "service_technician": ["tehnician service", "service technician"],
            "maintenance_engineer": ["maintenance engineer", "inginer mentenanta"],
            "facility_maintenance": ["facility technician", "tehnician facility"],
            "maintenance_manager": ["maintenance manager", "manager mentenanta"]
        },
        "hard_signals": ["mentenanta", "service", "reparatii", "preventiva", "echipamente", "mecanic", "electric", "automatizari"]
    },
    "MECHANICAL_ENGINEERING": {
        "label": "Mechanical Engineering",
        "subfamilies": {
            "mechanical_engineer": ["inginer mecanic", "mechanical engineer"],
            "design_engineer": ["inginer proiectant", "design engineer", "cad engineer"],
            "process_engineer": ["process engineer", "inginer proces"],
            "product_engineer": ["product engineer", "inginer produs"],
            "r_and_d": ["r&d engineer", "research development engineer"],
            "industrial_engineer": ["industrial engineer", "inginer industrial"],
            "cad_cam": ["cad cam", "catia", "solidworks engineer"],
            "tooling": ["tooling engineer", "inginer scule"],
            "technical_project": ["technical project engineer"],
            "mechanical_manager": ["mechanical engineering manager"]
        },
        "hard_signals": ["inginer mecanic", "cad", "solidworks", "catia", "proiectare", "mecanica", "industrial"]
    },
    "AUTOMOTIVE_SERVICE": {
        "label": "Automotive Service",
        "subfamilies": {
            "auto_mechanic": ["mecanic auto", "auto mechanic", "tehnician mecanic auto"],
            "assistant_mechanic": ["ajutor mecanic", "ucenic mecanic"],
            "auto_electrician": ["electrician auto", "auto electrician"],
            "diagnostics": ["diagnoza auto", "diagnostic auto", "diagnostic technician"],
            "body_repair": ["tinichigiu auto", "body repair", "car body repair"],
            "auto_painter": ["vopsitor auto", "auto painter"],
            "service_advisor": ["consilier service auto", "service advisor"],
            "itp": ["inspector itp", "itp"],
            "tire_service": ["vulcanizator", "tire technician"],
            "auto_service_manager": ["sef service auto", "service manager auto"]
        },
        "hard_signals": ["service auto", "mecanic auto", "diagnoza", "revizii", "frane", "suspensie", "distributii", "benzina", "diesel", "tinichigerie", "vopsitorie auto"]
    },
    "ELECTRICAL_ENGINEERING": {
        "label": "Electrical Engineering",
        "subfamilies": {
            "electrician": ["electrician", "electrician constructii", "electrician industrial"],
            "electrical_engineer": ["inginer electric", "electrical engineer"],
            "automation_engineer": ["automation engineer", "inginer automatizari"],
            "electronics": ["electronics engineer", "inginer electronist"],
            "panel_builder": ["tablotier", "electric panel builder"],
            "low_voltage": ["low voltage", "curenti slabi"],
            "high_voltage": ["high voltage", "inalta tensiune"],
            "plc": ["plc programmer", "programator plc"],
            "electrical_design": ["electrical design engineer"],
            "energy": ["energy engineer", "inginer energetician"]
        },
        "hard_signals": ["electrician", "electrice", "automatizari", "plc", "tablou electric", "curenti slabi", "tensiune"]
    },
    "CONSTRUCTION": {
        "label": "Construction",
        "subfamilies": {
            "civil_engineer": ["inginer constructii", "civil engineer"],
            "site_manager": ["sef santier", "site manager"],
            "construction_worker": ["muncitor constructii", "construction worker"],
            "mason": ["zidar", "mason"],
            "painter": ["zugrav", "zugrav vopsitor"],
            "tiler": ["faiantar", "tiler"],
            "plumber": ["instalator", "plumber"],
            "welder": ["sudor", "welder"],
            "carpenter": ["tamplar", "carpenter"],
            "quantity_surveyor": ["devizier", "quantity surveyor"]
        },
        "hard_signals": ["constructii", "santier", "zidarie", "zugrav", "faianta", "instalatii", "sudura", "devize"]
    },
    "HOSPITALITY": {
        "label": "Hospitality",
        "subfamilies": {
            "hotel_reception": ["receptioner hotel", "hotel receptionist", "front desk hotel"],
            "hotel_manager": ["hotel manager", "manager hotel"],
            "housekeeping": ["camerista", "housekeeping"],
            "reservation": ["agent rezervari", "reservation agent"],
            "concierge": ["concierge"],
            "guest_relations": ["guest relations"],
            "night_auditor": ["night auditor"],
            "spa_reception": ["spa receptionist"],
            "tourism_agent": ["agent turism", "travel agent"],
            "hotel_supervisor": ["hotel supervisor"]
        },
        "hard_signals": ["hotel", "receptie", "receptioner", "rezervari", "turism", "oaspeti", "guest"]
    },
    "FOOD_SERVICE": {
        "label": "Food Service",
        "subfamilies": {
            "cook": ["bucatar", "cook", "chef", "line cook"],
            "assistant_cook": ["ajutor bucatar", "kitchen helper"],
            "waiter": ["ospatar", "chelner", "waiter", "server"],
            "bartender": ["barman", "bartender"],
            "barista": ["barista"],
            "pastry": ["patiser", "pastry chef"],
            "confectioner": ["cofetar", "confectioner"],
            "kitchen_manager": ["sef bucatar", "head chef", "sous chef"],
            "dishwasher": ["spalator vase", "dishwasher"],
            "restaurant_manager": ["restaurant manager", "manager restaurant"]
        },
        "hard_signals": ["restaurant", "bucatarie", "bucatar", "ospatar", "servire", "bar", "meniuri", "horeca", "preparate"]
    },
    "HEALTHCARE": {
        "label": "Healthcare",
        "subfamilies": {
            "nurse": ["asistent medical", "nurse", "asistent medical generalist"],
            "doctor": ["medic", "doctor"],
            "pharmacist": ["farmacist", "pharmacist"],
            "kinetotherapy": ["kinetoterapeut", "physiotherapist"],
            "dentistry": ["asistent stomatologie", "dentist", "dental assistant"],
            "medical_reception": ["receptioner clinica", "medical receptionist"],
            "lab": ["laborant", "lab technician"],
            "caregiver": ["infirmier", "caregiver"],
            "radiology": ["radiology technician"],
            "medical_manager": ["manager clinica"]
        },
        "hard_signals": ["medical", "pacienti", "clinica", "spital", "asistent medical", "farmacie", "medic", "kineto"]
    },
    "EDUCATION": {
        "label": "Education",
        "subfamilies": {
            "teacher": ["profesor", "teacher"],
            "educator": ["educator", "invatator"],
            "trainer": ["trainer", "formator"],
            "university": ["lector universitar", "profesor universitar"],
            "tutor": ["tutor", "meditator"],
            "instructional_design": ["instructional designer"],
            "school_admin": ["secretar scoala"],
            "kindergarten": ["educatoare", "gradinita"],
            "language_teacher": ["profesor limbi straine"],
            "education_manager": ["director scoala", "manager educational"]
        },
        "hard_signals": ["profesor", "educatie", "predare", "cursuri", "training", "formator", "elevi", "studenti"]
    },
    "MANAGEMENT_LEADERSHIP": {
        "label": "Management & Leadership",
        "subfamilies": {
            "general_manager": ["general manager", "director general"],
            "operations_manager": ["operations manager", "manager operational"],
            "project_manager": ["project manager", "manager proiect"],
            "team_leader": ["team leader", "coordonator echipa"],
            "department_manager": ["department manager", "sef departament"],
            "branch_manager": ["branch manager", "manager filiala"],
            "plant_manager": ["plant manager", "manager fabrica"],
            "program_manager": ["program manager"],
            "business_manager": ["business manager"],
            "executive": ["ceo", "coo", "director executiv"]
        },
        "hard_signals": ["manager", "director", "coordonare", "echipa", "buget", "strategie", "operational", "project management"]
    }
}

# Familii apropiate, dar NU intra in top strict decat daca nu se cere filtru strict.
ADJACENT_FAMILIES = {
    "IT_SOFTWARE": {"DATA_BI", "IT_INFRASTRUCTURE", "CYBERSECURITY"},
    "IT_INFRASTRUCTURE": {"IT_SUPPORT", "CYBERSECURITY", "IT_SOFTWARE"},
    "DATA_BI": {"IT_SOFTWARE", "FINANCE_ACCOUNTING", "DIGITAL_MARKETING"},
    "DIGITAL_MARKETING": {"MARKETING"},
    "MARKETING": {"DIGITAL_MARKETING", "SALES_B2B"},
    "SALES_B2B": {"SALES_RETAIL", "MARKETING"},
    "SALES_RETAIL": {"SALES_B2B", "RETAIL_CASHIER"},
    "LOGISTICS_WAREHOUSE": {"TRANSPORTATION", "SUPPLY_CHAIN"},
    "TRANSPORTATION": {"LOGISTICS_WAREHOUSE"},
    "AUTOMOTIVE_SERVICE": {"MAINTENANCE_TECHNICAL", "MECHANICAL_ENGINEERING"},
    "MAINTENANCE_TECHNICAL": {"AUTOMOTIVE_SERVICE", "ELECTRICAL_ENGINEERING", "PRODUCTION_MANUFACTURING"},
    "HOSPITALITY": {"FOOD_SERVICE", "CUSTOMER_SUPPORT"},
    "FOOD_SERVICE": {"HOSPITALITY"},
    "FINANCE_ACCOUNTING": {"BANKING_INSURANCE"},
    "HR_RECRUITMENT": {"ADMINISTRATIVE"},
    "ADMINISTRATIVE": {"CUSTOMER_SUPPORT", "HR_RECRUITMENT"},
}

# Adaugam familie separata pentru casier, ca sa nu amestecam retail vanzari cu retail cash handling.
JOB_FAMILIES["RETAIL_CASHIER"] = {
    "label": "Retail Cashier",
    "subfamilies": {
        "cashier": ["casier", "cashier", "operator casa", "lucrator casierie", "casa de marcat"],
        "pos": ["operator pos", "pos cashier"],
        "money_handling": ["incasari", "numerar", "cash handling"],
        "retail_cashier": ["casier retail", "casier magazin"],
        "supermarket_cashier": ["casier supermarket", "casier hypermarket"],
        "front_cashier": ["front cashier"],
        "ticketing": ["ticketing cashier"],
        "billing_cashier": ["billing cashier"],
        "cash_office": ["cash office"],
        "head_cashier": ["sef casieri", "head cashier"]
    },
    "hard_signals": ["casier", "casa de marcat", "pos", "incasari", "numerar", "bon fiscal", "retail"]
}
ADJACENT_FAMILIES.setdefault("RETAIL_CASHIER", {"SALES_RETAIL", "CUSTOMER_SUPPORT"})
