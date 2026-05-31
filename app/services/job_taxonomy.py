"""
Taxonomie joburi NEXAS HR — 61 familii (31 originale + 30 noi).
Structura: Familie > Subfamilie > Specializari
Suporta romana si engleza.
"""

from __future__ import annotations

JOB_FAMILIES: dict = {

    # ─── IT & TECH ────────────────────────────────────────────────────────────

    "IT_SOFTWARE": {
        "label": "IT Software",
        "subfamilies": {
            "software_development": ["software developer", "software engineer", "programator", "dezvoltator software", "dezvoltator it", "developer", "application developer", "inginer software"],
            "backend": ["backend", "back end", "api developer", "python developer", "java developer", "php developer", "node developer", "c# developer", ".net developer", "ruby developer", "go developer"],
            "frontend": ["frontend", "front end", "react developer", "vue developer", "angular developer", "javascript developer", "typescript developer", "ui developer"],
            "fullstack": ["full stack", "fullstack", "full-stack", "full stack developer", "full stack engineer"],
            "mobile": ["mobile developer", "android developer", "ios developer", "flutter developer", "react native developer", "swift developer", "kotlin developer"],
            "qa_automation": ["qa automation", "automation tester", "test automation", "sdet", "qa engineer", "quality assurance engineer", "tester software"],
            "web_development": ["web developer", "web programmer", "dezvoltator web", "wordpress developer", "cms developer"],
            "database_development": ["sql developer", "database developer", "pl sql developer", "oracle developer", "postgresql developer"],
            "embedded": ["embedded developer", "embedded software", "firmware developer", "embedded systems engineer"],
            "game_development": ["game developer", "unity developer", "unreal developer", "game programmer"]
        },
        "specializations": {
            "backend": ["python", "java", ".net", "php", "node.js", "c#", "ruby", "go", "django", "flask", "spring", "laravel"],
            "frontend": ["react", "angular", "vue", "javascript", "typescript", "html", "css", "sass"],
            "mobile": ["android", "ios", "flutter", "react native", "swift", "kotlin"],
            "qa_automation": ["selenium", "cypress", "playwright", "jira", "postman", "testng"],
            "database_development": ["sql", "postgresql", "oracle", "mysql", "pl/sql"]
        },
        "hard_signals": ["python", "javascript", "typescript", "java", "c#", "php", "react", "angular", "vue", "node", "django", "flask", "spring", "sql", "git", "api", "backend", "frontend", "full stack", "programare", "software", "cod", "aplicatii"]
    },

    "IT_INFRASTRUCTURE": {
        "label": "IT Infrastructure",
        "subfamilies": {
            "system_admin": ["system administrator", "sysadmin", "administrator sistem", "administrator servere", "infrastructure engineer"],
            "network": ["network engineer", "administrator retea", "network administrator", "retelistica", "network specialist"],
            "devops": ["devops", "devops engineer", "site reliability", "sre", "platform engineer"],
            "cloud": ["cloud engineer", "aws engineer", "azure engineer", "gcp engineer", "cloud architect"],
            "linux": ["linux administrator", "linux engineer", "linux system admin"],
            "windows": ["windows administrator", "active directory administrator", "exchange administrator"],
            "database_admin": ["dba", "database administrator", "administrator baze de date"],
            "noc": ["noc engineer", "noc technician", "noc analyst"],
            "virtualization": ["vmware", "virtualization engineer", "hyper-v administrator"],
            "storage": ["storage engineer", "san administrator", "backup engineer"]
        },
        "specializations": {
            "cloud": ["aws", "azure", "gcp", "terraform", "kubernetes", "docker"],
            "devops": ["jenkins", "gitlab ci", "ansible", "terraform", "kubernetes", "docker"],
            "system_admin": ["linux", "windows server", "active directory", "vmware"]
        },
        "hard_signals": ["server", "linux", "windows server", "network", "retelistica", "aws", "azure", "docker", "kubernetes", "active directory", "vmware"]
    },

    "IT_SUPPORT": {
        "label": "IT Support",
        "subfamilies": {
            "helpdesk": ["helpdesk", "help desk", "it helpdesk", "l1 support"],
            "service_desk": ["service desk", "it service desk", "l2 support"],
            "technical_support": ["technical support", "suport tehnic", "support engineer", "it support specialist"],
            "desktop_support": ["desktop support", "field it technician", "tehnician it teren"],
            "hardware": ["hardware technician", "tehnician hardware", "it hardware"],
            "application_support": ["application support", "suport aplicatii", "erp support", "sap support"],
            "it_operator": ["operator it", "it operator", "computer operator"],
            "network_support": ["network support", "suport retea"],
            "mobile_support": ["mobile device support"],
            "av_support": ["av technician", "audio video support"]
        },
        "hard_signals": ["suport tehnic", "ticket", "helpdesk", "service desk", "hardware", "it support", "windows", "troubleshooting"]
    },

    "DATA_BI": {
        "label": "Data & BI",
        "subfamilies": {
            "data_analyst": ["data analyst", "analist date", "analyst data", "business analyst data"],
            "bi_developer": ["bi developer", "business intelligence", "power bi developer", "tableau developer", "qlik developer"],
            "data_engineer": ["data engineer", "inginer date", "big data engineer"],
            "data_science": ["data scientist", "machine learning engineer", "ml engineer", "ai engineer"],
            "reporting": ["reporting analyst", "analist raportare", "reporting specialist"],
            "etl": ["etl developer", "etl specialist", "data integration"],
            "master_data": ["master data specialist", "master data analyst", "mdm specialist"],
            "analytics": ["analytics specialist", "web analytics", "google analytics specialist"],
            "data_entry": ["operator introducere date", "data entry", "operator validare date"],
            "database_reporting": ["sql analyst", "sql reporting", "database analyst"]
        },
        "specializations": {
            "bi_developer": ["power bi", "tableau", "qlik", "looker", "excel"],
            "data_science": ["python", "r", "tensorflow", "pytorch", "scikit-learn"],
            "data_engineer": ["spark", "hadoop", "kafka", "airflow", "sql", "python"]
        },
        "hard_signals": ["excel", "sql", "power bi", "tableau", "raportare", "analiza date", "etl", "python", "dashboard"]
    },

    "CYBERSECURITY": {
        "label": "Cybersecurity",
        "subfamilies": {
            "soc": ["soc analyst", "security operations", "security analyst", "analyst soc"],
            "security_engineer": ["security engineer", "cybersecurity engineer", "information security engineer"],
            "penetration_testing": ["penetration tester", "ethical hacker", "pentest", "red team"],
            "grc": ["grc analyst", "information security officer", "iso 27001", "compliance security"],
            "incident_response": ["incident responder", "incident response", "dfir analyst"],
            "iam": ["iam specialist", "identity access", "pam specialist"],
            "network_security": ["network security engineer", "firewall engineer"],
            "cloud_security": ["cloud security engineer", "devsecops"],
            "appsec": ["application security", "appsec engineer"],
            "security_admin": ["security administrator", "security operations manager"]
        },
        "hard_signals": ["cybersecurity", "security", "soc", "siem", "vulnerability", "penetration", "iso 27001", "firewall"]
    },

    # ─── FINANCE & BANKING ────────────────────────────────────────────────────

    "FINANCE_ACCOUNTING": {
        "label": "Finance & Accounting",
        "subfamilies": {
            "accounting": ["contabil", "accountant", "contabilitate", "contabil primar"],
            "chief_accountant": ["contabil sef", "chief accountant", "director financiar contabil"],
            "financial_analysis": ["analist financiar", "financial analyst", "fp&a analyst"],
            "economist": ["economist", "referent economic"],
            "audit": ["auditor", "audit financiar", "auditor intern", "auditor extern"],
            "tax": ["tax specialist", "specialist taxe", "consultant fiscal"],
            "payroll": ["payroll", "salarizare", "specialist salarizare"],
            "accounts_payable": ["accounts payable", "ap accountant", "furnizori contabilitate"],
            "accounts_receivable": ["accounts receivable", "ar accountant", "clienti contabilitate"],
            "controller": ["financial controller", "controller financiar", "business controller"]
        },
        "specializations": {
            "accounting": ["saga", "ciel", "winmentor", "sap fi", "oracle financials"],
            "audit": ["ifrs", "us gaap", "isa", "ceccar", "cafr"]
        },
        "hard_signals": ["contabilitate", "balanta", "bilant", "saga", "ceccar", "facturi", "taxe", "raportare financiara", "excel"]
    },

    "BANKING_INSURANCE": {
        "label": "Banking & Insurance",
        "subfamilies": {
            "banking_officer": ["ofiter bancar", "banking officer", "consilier bancar", "ofiter credite retail"],
            "credit": ["analist credite", "credit analyst", "ofiter credite"],
            "insurance_agent": ["agent asigurari", "insurance agent", "consultant asigurari"],
            "claims": ["claims specialist", "specialist daune", "inspector daune"],
            "underwriting": ["underwriter", "subscriere", "specialist subscriere"],
            "risk": ["risk analyst", "analist risc", "risk manager"],
            "leasing": ["leasing officer", "consultant leasing", "analist leasing"],
            "compliance_banking": ["aml analyst", "compliance banking", "ofiter conformitate"],
            "brokerage": ["broker asigurari", "broker credite", "financial advisor"],
            "treasury": ["trezorier", "treasury analyst", "dealer valutar"]
        },
        "hard_signals": ["banca", "credit", "asigurari", "daune", "underwriting", "aml", "risk", "leasing"]
    },

    # ─── HR, LEGAL, ADMIN ─────────────────────────────────────────────────────

    "HR_RECRUITMENT": {
        "label": "HR & Recruitment",
        "subfamilies": {
            "recruitment": ["recruiter", "specialist recrutare", "recruitment specialist", "talent acquisition specialist", "headhunter"],
            "hr_generalist": ["hr generalist", "specialist resurse umane", "resurse umane", "hr specialist"],
            "hrbp": ["hr business partner", "hrbp", "people partner"],
            "payroll_hr": ["inspector resurse umane", "inspector hr", "revisal", "specialist salarizare hr"],
            "training": ["training specialist", "learning development", "l&d specialist", "formator"],
            "comp_ben": ["compensation benefits", "comp & ben", "total rewards specialist"],
            "talent": ["talent management", "talent development"],
            "employee_relations": ["employee relations", "relatii angajati"],
            "hr_admin": ["hr admin", "administrator hr", "asistent hr"],
            "organizational_development": ["organizational development", "od specialist", "change management"]
        },
        "hard_signals": ["recrutare", "interviuri", "resurse umane", "hr", "revisal", "talent", "employee", "onboarding"]
    },

    "LEGAL": {
        "label": "Legal",
        "subfamilies": {
            "legal_advisor": ["jurist", "legal advisor", "consilier juridic", "specialist juridic"],
            "lawyer": ["avocat", "lawyer", "attorney", "avocat colaborator"],
            "paralegal": ["paralegal", "asistent juridic", "legal assistant"],
            "compliance": ["compliance officer", "specialist conformitate", "risk compliance"],
            "contracts": ["contract manager", "specialist contracte", "contract administrator"],
            "notary": ["notar", "notary", "asistent notar"],
            "gdpr": ["gdpr specialist", "data protection officer", "dpo"],
            "litigation": ["litigation", "contencios", "avocat litigii"],
            "public_procurement": ["achizitii publice", "specialist achizitii publice"],
            "legal_admin": ["secretar juridic", "grefa", "legal secretary"]
        },
        "hard_signals": ["juridic", "contracte", "lege", "avocat", "litigii", "gdpr", "compliance", "codul muncii"]
    },

    "ADMINISTRATIVE": {
        "label": "Administrative",
        "subfamilies": {
            "office_assistant": ["office assistant", "asistent birou", "asistent administrativ"],
            "secretary": ["secretara", "secretar", "secretary", "secretar executiv"],
            "reception_office": ["receptionist office", "receptioner birou", "front desk"],
            "data_entry_admin": ["operator date", "operator introducere date", "data entry"],
            "document_controller": ["document controller", "arhivar", "gestionar documente"],
            "personal_assistant": ["personal assistant", "asistent manager", "executive assistant"],
            "administrative_clerk": ["referent", "referent administrativ", "functionar administrativ"],
            "registry": ["registratura", "operator registratura"],
            "office_manager": ["office manager", "administrator birou"],
            "facility_admin": ["facility administrator", "administrator cladire"]
        },
        "hard_signals": ["administrativ", "documente", "birou", "secretariat", "arhivare", "excel", "date"]
    },

    # ─── MARKETING & SALES ────────────────────────────────────────────────────

    "MARKETING": {
        "label": "Marketing",
        "subfamilies": {
            "marketing_generalist": ["marketing specialist", "specialist marketing", "marketing coordinator", "marketing executive"],
            "brand": ["brand manager", "brand specialist", "brand coordinator"],
            "trade_marketing": ["trade marketing", "trade marketing specialist"],
            "product_marketing": ["product marketing", "product marketing manager"],
            "market_research": ["market research", "cercetare piata", "research analyst"],
            "communications": ["communications specialist", "specialist comunicare", "corporate communications"],
            "pr": ["pr specialist", "public relations", "specialist relatii publice"],
            "events": ["event marketing", "event specialist", "event manager", "organizator evenimente"],
            "copywriting": ["copywriter", "content writer", "redactor"],
            "marketing_manager": ["marketing manager", "director marketing", "head of marketing"]
        },
        "hard_signals": ["marketing", "brand", "campanii", "comunicare", "content", "pr", "evenimente", "strategie marketing"]
    },

    "DIGITAL_MARKETING": {
        "label": "Digital Marketing",
        "subfamilies": {
            "seo": ["seo specialist", "seo expert", "search engine optimization"],
            "ppc": ["ppc specialist", "google ads specialist", "paid search", "sem specialist", "facebook ads specialist"],
            "social_media": ["social media", "social media manager", "community manager"],
            "performance": ["performance marketing", "growth marketing", "growth hacker"],
            "email_marketing": ["email marketing", "crm marketing", "marketing automation specialist"],
            "ecommerce_marketing": ["ecommerce marketing", "e-commerce marketing"],
            "content_marketing": ["content marketing", "content creator", "content strategist"],
            "affiliate": ["affiliate marketing", "affiliate manager"],
            "analytics_digital": ["digital analytics", "google analytics", "web analytics specialist"],
            "digital_manager": ["digital marketing manager", "head of digital"]
        },
        "specializations": {
            "ppc": ["google ads", "facebook ads", "tiktok ads", "linkedin ads"],
            "seo": ["semrush", "ahrefs", "screaming frog", "google search console"]
        },
        "hard_signals": ["seo", "ppc", "google ads", "facebook ads", "social media", "analytics", "campanii digitale", "content"]
    },

    "SALES_B2B": {
        "label": "Sales B2B",
        "subfamilies": {
            "sales_representative": ["agent vanzari", "reprezentant vanzari", "sales representative", "sales agent", "agent comercial", "consultant vanzari b2b"],
            "field_sales": ["field sales", "vanzari teren", "reprezentant teren", "area sales representative"],
            "account_manager": ["account manager", "manager cont", "key account manager", "key account"],
            "business_development": ["business development", "bdm", "business development manager"],
            "inside_sales": ["inside sales", "telesales b2b", "sales development representative"],
            "sales_engineer": ["sales engineer", "inginer vanzari", "technical sales"],
            "area_sales": ["area sales manager", "regional sales manager"],
            "export_sales": ["export sales", "international sales"],
            "presales": ["presales consultant", "solutions consultant"],
            "sales_manager": ["sales manager", "director vanzari", "national sales manager"]
        },
        "hard_signals": ["vanzari", "clienti", "portofoliu", "negociere", "ofertare", "b2b", "agent comercial", "sales", "crm", "pipeline"]
    },

    "SALES_RETAIL": {
        "label": "Sales Retail",
        "subfamilies": {
            "shop_assistant": ["lucrator comercial", "vanzator", "sales assistant", "shop assistant", "consilier vanzari"],
            "cashier_sales": ["casier vanzator", "cashier sales"],
            "store_manager": ["store manager", "sef magazin", "manager magazin"],
            "department_manager": ["sef raion", "department manager", "manager raion"],
            "merchandiser": ["merchandiser", "mercantizor", "visual merchandiser"],
            "promoter": ["promoter", "brand promoter", "demonstrator produse"],
            "retail_consultant": ["consultant retail", "consultant showroom"],
            "store_supervisor": ["retail supervisor", "supervisor vanzari"],
            "loss_prevention": ["agent securitate retail", "loss prevention"],
            "outlet_manager": ["outlet manager", "manager punct de lucru"]
        },
        "hard_signals": ["magazin", "retail", "clienti", "vanzare", "casa de marcat", "raft", "marfa", "showroom"]
    },

    "RETAIL_CASHIER": {
        "label": "Retail Cashier",
        "subfamilies": {
            "cashier": ["casier", "cashier", "operator casa", "lucrator casierie", "casa de marcat"],
            "pos": ["operator pos", "pos cashier"],
            "supermarket_cashier": ["casier supermarket", "casier hypermarket", "casier magazin alimentar"],
            "head_cashier": ["sef casieri", "head cashier", "coordonator casieri"],
            "ticketing": ["ticketing cashier", "agent bilete"],
            "banking_cashier": ["casier bancar", "ghiseu bancar"],
            "fuel_cashier": ["casier benzinarie", "operator pompe"],
            "parking_cashier": ["casier parcare"],
            "event_cashier": ["casier evenimente"],
            "cash_office": ["cash office", "administrator casa"]
        },
        "hard_signals": ["casier", "casa de marcat", "pos", "incasari", "numerar", "bon fiscal"]
    },

    # ─── CUSTOMER SUPPORT & PROCUREMENT ──────────────────────────────────────

    "CUSTOMER_SUPPORT": {
        "label": "Customer Support",
        "subfamilies": {
            "call_center": ["operator call center", "call center", "agent call center", "telefonie clienti"],
            "customer_service": ["customer service", "relatii clienti", "suport clienti", "customer care"],
            "customer_success": ["customer success", "customer success manager"],
            "backoffice": ["back office", "backoffice", "operator backoffice"],
            "technical_customer_support": ["technical customer support", "suport tehnic clienti"],
            "chat_support": ["chat support", "live chat", "online support"],
            "complaints": ["complaints specialist", "specialist reclamatii"],
            "dispatch_support": ["dispecer", "dispatcher", "dispecer urgente"],
            "order_support": ["order support", "procesare comenzi", "operator comenzi"],
            "contact_center_supervisor": ["team leader call center", "contact center supervisor"]
        },
        "hard_signals": ["call center", "clienti", "suport", "relatii clienti", "telefon", "ticket", "reclamatii", "dispecer"]
    },

    "PROCUREMENT": {
        "label": "Procurement & Purchasing",
        "subfamilies": {
            "buyer": ["buyer", "achizitor", "specialist achizitii", "purchasing officer"],
            "procurement_specialist": ["procurement specialist", "specialist procurement"],
            "purchasing": ["purchasing specialist", "aprovizionare", "specialist aprovizionare"],
            "strategic_sourcing": ["strategic sourcing", "category management"],
            "vendor_management": ["vendor manager", "supplier manager"],
            "category_buyer": ["category buyer", "category manager procurement"],
            "import_purchasing": ["import specialist", "achizitii import"],
            "contract_procurement": ["contract procurement", "specialist contracte achizitii"],
            "materials_planner": ["materials planner", "planificator materiale"],
            "procurement_manager": ["procurement manager", "manager achizitii", "director achizitii"]
        },
        "hard_signals": ["achizitii", "furnizori", "negociere", "aprovizionare", "purchase", "procurement", "rfq", "tender"]
    },

    "SUPPLY_CHAIN": {
        "label": "Supply Chain",
        "subfamilies": {
            "planner": ["supply planner", "demand planner", "planner supply chain"],
            "inventory": ["inventory specialist", "control stocuri", "inventory analyst"],
            "import_export": ["import export specialist", "export import coordinator"],
            "customs": ["customs specialist", "specialist vamal", "declarant vamal"],
            "transport_planning": ["transport planner", "transport coordinator sc"],
            "supply_chain_analyst": ["supply chain analyst", "sc analyst"],
            "operations_planning": ["operations planner", "operations analyst"],
            "s_and_op": ["s&op analyst", "sales and operations planning"],
            "warehouse_planning": ["warehouse planner", "depot planner"],
            "supply_chain_manager": ["supply chain manager", "director supply chain"]
        },
        "hard_signals": ["supply chain", "stocuri", "planificare", "transport", "vama", "import", "export", "logistica", "s&op"]
    },

    # ─── LOGISTICA & TRANSPORT ────────────────────────────────────────────────

    "LOGISTICS_WAREHOUSE": {
        "label": "Logistics & Warehouse",
        "subfamilies": {
            "warehouse_worker": ["manipulant marfa", "warehouse worker", "lucrator depozit", "muncitor depozit", "operator depozit"],
            "picker": ["picker", "order picker", "pregatitor comenzi", "picking operator"],
            "packer": ["packer", "ambalator", "ambalare", "operator ambalare"],
            "forklift": ["stivuitorist", "forklift operator", "operator stivuitor", "operator motostivuitor"],
            "warehouse_clerk": ["gestionar depozit", "depozitar", "warehouse clerk", "magazioner"],
            "inventory_warehouse": ["inventarist", "inventory clerk", "operator inventar"],
            "loading": ["incarcare descarcare", "loader", "unloader", "dockar"],
            "warehouse_supervisor": ["sef depozit", "warehouse supervisor", "team leader depozit"],
            "logistics_operator": ["operator logistica", "logistics operator"],
            "wms_operator": ["operator wms", "wms specialist"]
        },
        "hard_signals": ["depozit", "marfa", "stivuitor", "picking", "packing", "gestionar", "incarcare", "descarcare", "wms", "iscir", "manipulant"]
    },

    "TRANSPORTATION": {
        "label": "Transportation",
        "subfamilies": {
            "driver_b": ["sofer categoria b", "sofer b", "driver b", "sofer distributie", "sofer marfa mica"],
            "driver_c": ["sofer categoria c", "sofer c", "truck driver c", "sofer camion"],
            "driver_ce": ["sofer tir", "sofer ce", "tir", "truck driver", "hgv driver", "sofer trailer", "sofer cap tractor"],
            "courier": ["curier", "courier", "livrator", "delivery driver", "agent livrare"],
            "bus_driver": ["sofer autobuz", "bus driver", "sofer transport persoane", "sofer microbuz"],
            "taxi": ["taxi driver", "sofer taxi", "rideshare driver"],
            "dispatcher_transport": ["dispecer transport", "transport dispatcher", "dispecer flota"],
            "fleet": ["fleet manager", "administrator flota", "fleet coordinator"],
            "transport_coordinator": ["transport coordinator", "coordonator transport", "expeditor transport"],
            "route_planner": ["route planner", "planificator rute"]
        },
        "specializations": {
            "driver_ce": ["adr", "transport frigorific", "transport agabaritic", "tahograf"],
            "driver_c": ["gps", "tahograf", "livrare marfa"]
        },
        "hard_signals": ["sofer", "permis", "categoria b", "categoria c", "categoria ce", "tir", "livrari", "distributie", "curier", "tahograf"]
    },

    "HEAVY_EQUIPMENT": {
        "label": "Utilaje Grele",
        "subfamilies": {
            "crane_operator": ["macaragiu", "crane operator", "operator macara", "operator pod rulant"],
            "excavator_operator": ["excavatorist", "excavator operator", "operator excavator"],
            "bulldozer_operator": ["buldozerist", "bulldozer operator", "operator buldozer"],
            "forklift_heavy": ["operator stivuitor greu", "reach truck operator"],
            "roller_operator": ["operator compactor", "roller operator", "sofer compactor"],
            "paver_operator": ["operator asfaltor", "asphalt paver operator"],
            "grader_operator": ["operator greder", "grader operator"],
            "loader_operator": ["operator incarcator frontal", "wheel loader operator", "buldoexcavatorist"],
            "telehandler": ["operator telehandler", "operator nacela", "operator platforma elevatoare"],
            "heavy_equipment_mechanic": ["mecanic utilaje grele", "service utilaje grele"]
        },
        "hard_signals": ["macaragiu", "excavator", "buldozer", "utilaje grele", "pod rulant", "nacela", "iscir", "permis iscir", "compactor"]
    },

    "FREIGHT_FORWARDING": {
        "label": "Expeditii & Freight Forwarding",
        "subfamilies": {
            "freight_forwarder": ["freight forwarder", "expeditor marfa", "agent expeditii", "freight coordinator"],
            "air_freight": ["air freight agent", "cargo aerian", "agent cargo aerian"],
            "sea_freight": ["sea freight agent", "cargo maritim", "agent maritim"],
            "road_freight": ["road freight", "transport rutier international", "expeditor rutier"],
            "multimodal": ["multimodal transport", "intermodal coordinator"],
            "customs_agent": ["agent vamal", "broker vamal", "declarant vamal"],
            "import_agent": ["agent import", "import coordinator"],
            "export_agent": ["agent export", "export coordinator"],
            "tracking": ["tracking specialist", "shipment tracker"],
            "forwarding_manager": ["forwarding manager", "manager expeditii"]
        },
        "hard_signals": ["expeditii", "freight", "cargo", "awb", "bill of lading", "incoterms", "transport international", "vamal", "vama"]
    },

    "DANGEROUS_GOODS_ADR": {
        "label": "Transport ADR & Marfuri Periculoase",
        "subfamilies": {
            "adr_driver": ["sofer adr", "adr driver", "sofer transport marfuri periculoase"],
            "adr_specialist": ["specialist adr", "adr specialist", "consilier adr"],
            "hazmat_handler": ["hazmat handler", "operator marfuri periculoase"],
            "chemical_transport": ["transport chimicale", "transport substante periculoase"],
            "fuel_transport": ["sofer cisterna", "transport combustibil", "sofer combustibil"],
            "gas_transport": ["transport gaze", "sofer gaze"],
            "adr_instructor": ["instructor adr", "formator adr"],
            "dangerous_goods_doc": ["documentatii adr", "documente marfuri periculoase"],
            "port_hazmat": ["hazmat port operator"],
            "adr_coordinator": ["coordonator transport adr"]
        },
        "hard_signals": ["adr", "marfuri periculoase", "cisterna", "transport periculos", "hazmat", "substante periculoase"]
    },

    "COLD_CHAIN": {
        "label": "Logistica Frigorifica",
        "subfamilies": {
            "cold_storage_worker": ["lucrator depozit frigorific", "cold storage worker", "operator depozit frigorific"],
            "refrigerated_driver": ["sofer transport refrigerat", "sofer frigorific", "sofer marfa perisabila"],
            "cold_chain_coordinator": ["coordonator lant frig", "cold chain coordinator"],
            "temperature_controller": ["specialist temperatura", "monitor temperatura"],
            "frozen_operator": ["operator produse congelate", "frozen warehouse operator"],
            "pharma_cold": ["logistica farmaceutica frigorifica", "pharma cold chain"],
            "cold_chain_manager": ["manager lant frig", "cold chain manager"],
            "reefer_technician": ["tehnician trailer frigorific", "reefer technician"],
            "quality_cold": ["inspector calitate frigorific"],
            "cold_dispatch": ["dispecer frigorific"]
        },
        "hard_signals": ["frigorific", "lant frig", "temperatura controlata", "refrigerat", "congelat", "perisabile", "cold chain"]
    },

    "POSTAL_SERVICES": {
        "label": "Servicii Postale & Curierat",
        "subfamilies": {
            "postal_worker": ["lucrator postal", "postal worker", "agent posta"],
            "sorter": ["sortator postal", "sorter", "operator sortare"],
            "mail_carrier": ["factor postal", "mail carrier", "distribuitor posta"],
            "parcel_handler": ["manipulant colete", "parcel handler"],
            "post_office_agent": ["agent oficiu postal", "ghiseu posta"],
            "courier_operator": ["operator curierat", "curier express"],
            "hub_operator": ["operator hub curierat", "hub sortare"],
            "postal_supervisor": ["sef tura sortare", "supervisor postal"],
            "postal_admin": ["administrator oficiu postal"],
            "last_mile": ["last mile delivery", "livrare ultima mila"]
        },
        "hard_signals": ["posta", "sortare", "colete", "curierat", "livrare", "oficiu postal", "factor postal", "dpd", "gls", "fan courier"]
    },

    "RAIL_TRANSPORT": {
        "label": "Transport Feroviar",
        "subfamilies": {
            "train_driver": ["mecanic locomotiva", "train driver", "conducator tren"],
            "conductor": ["conductor tren", "revizor bilete tren", "agent tren"],
            "rail_technician": ["tehnician cai ferate", "revizor vagoane", "electrician cfr"],
            "rail_dispatcher": ["dispecer cfr", "dispecer feroviar"],
            "rail_maintenance": ["mecanic intretinere cai ferate", "lucrator linii cfr"],
            "station_agent": ["agent gara", "casier gara", "agent statie"],
            "freight_rail": ["agent marfuri cfr", "wagon coordinator"],
            "signal_technician": ["tehnician semnalizare feroviara"],
            "rail_infrastructure": ["inginer infrastructura feroviara"],
            "rail_manager": ["manager regional cfr"]
        },
        "hard_signals": ["cfr", "feroviar", "locomotiva", "tren", "cale ferata", "gara", "revizor", "mecanic cfr"]
    },

    "MARITIME_TRANSPORT": {
        "label": "Transport Naval & Maritim",
        "subfamilies": {
            "sailor": ["marinar", "matelot", "sailor", "seafarer"],
            "captain": ["capitan nava", "captain", "comandant nava"],
            "navigation_officer": ["ofiter navigatie", "navigation officer", "pilot fluvial"],
            "ship_mechanic": ["mecanic naval", "ship engineer", "ofiter mecanic nava"],
            "port_worker": ["lucrator port", "stevedore", "docker port"],
            "port_agent": ["agent port", "port agent"],
            "maritime_dispatcher": ["dispecer port", "port dispatcher"],
            "ship_captain_river": ["capitan nava fluviala", "pilot dunare"],
            "vessel_operator": ["operator nava", "vessel coordinator"],
            "maritime_manager": ["manager maritim", "ship manager"]
        },
        "hard_signals": ["maritim", "naval", "nava", "port", "marinar", "capitan nava", "ofiter navigatie", "dunare", "constanta", "stcw"]
    },

    "AVIATION_GROUND": {
        "label": "Handling Aviatic & Aeroporturi",
        "subfamilies": {
            "ground_handler": ["handler aviatic", "ground handler", "agent handling"],
            "check_in_agent": ["agent check-in", "check in agent", "airport agent"],
            "baggage_handler": ["operator bagaje", "baggage handler"],
            "cargo_agent": ["agent cargo aerian", "cargo handler"],
            "ramp_agent": ["agent pista", "ramp agent", "ramp operator"],
            "security_airport": ["agent securitate aeroport", "airport security officer"],
            "dispatcher_aviation": ["dispecer aviatic", "operations dispatcher"],
            "fueling_technician": ["operator alimentare aeronave", "fueling agent"],
            "customer_service_airport": ["customer service airport", "agent asistenta pasageri"],
            "aviation_manager": ["manager operatiuni aeroportuare"]
        },
        "hard_signals": ["aeroport", "aviatie", "handling", "aeronave", "cargo aerian", "check-in", "pista", "terminal", "iata"]
    },

    # ─── PRODUCTIE & INDUSTRIE ────────────────────────────────────────────────

    "PRODUCTION_MANUFACTURING": {
        "label": "Production & Manufacturing",
        "subfamilies": {
            "production_operator": ["operator productie", "production operator", "muncitor productie"],
            "machine_operator": ["operator masina", "machine operator", "operator utilaj productie"],
            "assembly": ["asamblor", "assembly operator", "operator asamblare", "montator productie"],
            "cnc": ["operator cnc", "cnc operator", "frezor cnc", "strungist cnc"],
            "quality_control": ["control calitate", "quality control", "qc inspector", "inspector calitate productie"],
            "line_leader": ["line leader", "sef linie productie", "team leader productie"],
            "production_planner": ["planificator productie", "production planner"],
            "process_operator": ["operator proces", "process operator"],
            "manufacturing_engineer": ["inginer productie", "manufacturing engineer", "process engineer productie"],
            "production_manager": ["production manager", "manager productie", "director productie"]
        },
        "specializations": {
            "cnc": ["fanuc", "siemens cnc", "heidenhain", "g-code"],
            "quality_control": ["iso 9001", "6 sigma", "spc", "fmea", "ppap"]
        },
        "hard_signals": ["productie", "linie", "masina", "asamblare", "cnc", "calitate", "proces", "fabricatie", "schimburi", "norme productie"]
    },

    "FOOD_PRODUCTION": {
        "label": "Productie Alimentara",
        "subfamilies": {
            "food_operator": ["operator productie alimentara", "muncitor productie alimentara"],
            "slaughterhouse": ["muncitor abator", "operator abator", "slaughterhouse worker"],
            "meat_processing": ["preparator produse din carne", "operator procesare carne"],
            "dairy_operator": ["operator produse lactate", "dairy operator"],
            "bakery_production": ["brutar industrial", "baker industrial", "operator brutarie"],
            "confectionery_production": ["operator cofetarie industriala"],
            "beverage_operator": ["operator productie bauturi", "beverage production operator"],
            "food_packer": ["ambalator produse alimentare", "food packer"],
            "food_quality": ["inspector calitate alimente", "food quality inspector"],
            "food_production_manager": ["sef sectie productie alimentara", "food production manager"]
        },
        "hard_signals": ["productie alimentara", "abator", "procesare carne", "lactate", "brutarie", "alimente", "haccp", "igiena alimentara"]
    },

    "TEXTILE_GARMENT": {
        "label": "Textile & Confectii",
        "subfamilies": {
            "sewing_operator": ["operator masini cusut", "sewing operator", "cosatoreasa", "cusatoreasa"],
            "tailor_industrial": ["croitor industrial", "industrial tailor", "confectioner textile"],
            "pattern_maker": ["tipar", "pattern maker", "modeliera"],
            "textile_operator": ["operator textile", "textile machine operator"],
            "knitting_operator": ["operator tricotaje", "knitting machine operator"],
            "dyeing_operator": ["operator vopsitorie textile"],
            "quality_textile": ["inspector calitate textile"],
            "cutting_operator": ["operator croire", "cutting operator", "stantator"],
            "embroidery": ["operator broderie", "embroidery operator"],
            "textile_manager": ["sef atelier textile", "manager productie textile"]
        },
        "hard_signals": ["textile", "confectii", "cusut", "croitorie", "masina cusut", "ata", "stofa", "tricotaje"]
    },

    "WOOD_FURNITURE_PRODUCTION": {
        "label": "Prelucrare Lemn & Mobila",
        "subfamilies": {
            "carpenter_production": ["tamplar productie", "carpenter industrial", "tamplarie productie"],
            "cnc_wood": ["operator cnc lemn", "cnc wood operator"],
            "furniture_assembler": ["montator mobila", "furniture assembler", "asamblor mobila"],
            "lacquer_varnish": ["lacuitor", "vopsitor mobila", "finisator lemn"],
            "parquet_flooring": ["parchetar", "instalator parchet", "montator pardoseli lemn"],
            "sawmill_operator": ["operator fierastrau", "sawmill operator"],
            "furniture_upholstery": ["tapiter", "upholsterer", "tapiterie mobila"],
            "wood_quality": ["inspector calitate lemn"],
            "wood_production_manager": ["sef atelier tamplar", "manager productie lemn"],
            "wood_dryer": ["operator uscatorie lemn"]
        },
        "hard_signals": ["tamplarie", "mobila", "lemn", "parchet", "cnc lemn", "lacuire", "tapiterie", "furniruit"]
    },

    "PRINTING_PACKAGING": {
        "label": "Tipar & Ambalaje",
        "subfamilies": {
            "offset_operator": ["tipograf", "offset operator", "operator tipar offset"],
            "digital_print": ["operator tipar digital", "digital print operator"],
            "flexo_operator": ["operator flexografie", "flexo printer"],
            "packaging_operator": ["operator ambalaje", "packaging operator", "ambalator industrial"],
            "prepress": ["operator prepress", "prepress specialist", "dtp operator"],
            "bookbinding": ["legator carte", "bookbinder"],
            "label_operator": ["operator etichete", "label printer operator"],
            "carton_operator": ["operator carton", "cartonaj operator"],
            "print_quality": ["inspector calitate tipar"],
            "print_manager": ["sef productie tipar", "print production manager"]
        },
        "hard_signals": ["tipar", "imprimare", "offset", "flexografie", "ambalaje", "dtp", "prepress", "print", "tipografie"]
    },

    "CHEMICAL_PROCESS": {
        "label": "Industrie Chimica & Procesare",
        "subfamilies": {
            "chemical_operator": ["operator instalatii chimice", "chemical process operator"],
            "lab_chemist": ["chimist laborator", "laborant chimie", "lab technician chemistry"],
            "refinery_operator": ["operator rafinarie", "refinery operator"],
            "pharmaceutical_operator": ["operator productie farmaceutica", "pharma operator"],
            "paint_operator": ["operator productie vopsele"],
            "environmental_chemistry": ["specialist mediu chimie", "environmental chemist"],
            "quality_chemistry": ["inspector calitate chimica", "quality chemist"],
            "safety_chemical": ["specialist securitate chimica", "hse chemical"],
            "chemical_manager": ["manager productie chimica"],
            "plastics_chemical": ["operator mase plastice chimice"]
        },
        "hard_signals": ["chimic", "reactori", "distilare", "procese chimice", "laborator chimic", "substante", "industrie chimica", "farmaceutic productie"]
    },

    "PLASTICS_RUBBER": {
        "label": "Mase Plastice & Cauciuc",
        "subfamilies": {
            "injection_operator": ["operator injectie plastic", "injection molding operator", "presator plastic"],
            "extrusion_operator": ["operator extrudare", "extrusion operator"],
            "blow_molding": ["operator suflare plastic", "blow molding operator"],
            "rubber_operator": ["operator cauciuc", "rubber production operator", "vulcanizator industrial"],
            "mold_setter": ["reglor injectie", "mold setter"],
            "thermoforming": ["operator termoformare", "thermoforming operator"],
            "quality_plastics": ["inspector calitate plastic"],
            "plastic_assembler": ["montator plastic", "plastic assembler"],
            "recycling_operator": ["operator reciclare plastic"],
            "plastics_manager": ["manager productie plastic"]
        },
        "hard_signals": ["plastic", "injectie", "extrudare", "cauciuc", "forme injectie", "termoformare", "mase plastice"]
    },

    "AGRICULTURE_FARMING": {
        "label": "Agricultura & Zootehnie",
        "subfamilies": {
            "tractor_operator": ["tractorist", "tractor operator", "operator utilaje agricole"],
            "combine_operator": ["operator combina", "combine harvester operator"],
            "agricultural_worker": ["muncitor agricol", "agricultural worker", "lucrator agricol"],
            "agronomist": ["agronom", "agronomist", "inginer agronom"],
            "livestock_worker": ["ingrijitor animale", "livestock worker", "crescator animale"],
            "greenhouse_operator": ["operator sera", "greenhouse worker", "lucrator sera"],
            "irrigation_operator": ["operator irigatii", "irrigation technician"],
            "pesticide_applicator": ["operator fitosanitar", "pesticide applicator"],
            "farm_manager": ["manager ferma", "farm manager", "administrator ferma"],
            "veterinarian": ["veterinar", "veterinary doctor", "asistent veterinar"]
        },
        "hard_signals": ["agricultura", "ferma", "tractoare", "utilaje agricole", "recoltat", "plantat", "agronom", "zootehnie", "animale ferma"]
    },

    # ─── CONSTRUCTII & INSTALATII ─────────────────────────────────────────────

    "MAINTENANCE_TECHNICAL": {
        "label": "Maintenance & Technical Service",
        "subfamilies": {
            "maintenance_technician": ["tehnician mentenanta", "maintenance technician", "specialist mentenanta"],
            "mechanical_maintenance": ["mecanic mentenanta", "mechanical maintenance", "mecanic intretinere"],
            "electrical_maintenance": ["electrician mentenanta", "electrical maintenance", "electrician intretinere"],
            "industrial_mechanic": ["lacatus mecanic", "mecanic industrial", "fitter mechanic"],
            "automation": ["automatist", "automation technician", "specialist automatizari"],
            "hvac_maintenance": ["frigotehnist mentenanta", "hvac technician", "tehnician climatizare"],
            "service_technician": ["tehnician service", "service technician", "field service engineer"],
            "maintenance_engineer": ["maintenance engineer", "inginer mentenanta", "reliability engineer"],
            "facility_maintenance": ["facility technician", "tehnician facility"],
            "maintenance_manager": ["maintenance manager", "manager mentenanta", "director tehnic mentenanta"]
        },
        "hard_signals": ["mentenanta", "service", "reparatii", "preventiva", "echipamente", "mecanic", "electric", "automatizari", "revizii", "intretinere"]
    },

    "MECHANICAL_ENGINEERING": {
        "label": "Mechanical Engineering",
        "subfamilies": {
            "mechanical_engineer": ["inginer mecanic", "mechanical engineer", "inginer constructii mecanice"],
            "design_engineer": ["inginer proiectant", "design engineer", "cad engineer"],
            "process_engineer": ["process engineer", "inginer proces", "inginer procese industriale"],
            "product_engineer": ["product engineer", "inginer produs"],
            "r_and_d": ["r&d engineer", "research development engineer", "inginer cercetare"],
            "industrial_engineer": ["industrial engineer", "inginer industrial"],
            "cad_cam": ["cad cam specialist", "catia specialist", "solidworks engineer"],
            "tooling": ["tooling engineer", "inginer scule", "SDV engineer"],
            "quality_engineer": ["quality engineer", "inginer calitate", "inginer qms"],
            "mechanical_manager": ["mechanical engineering manager", "director tehnic mecanic"]
        },
        "specializations": {
            "design_engineer": ["solidworks", "catia", "autocad", "inventor", "creo"],
            "cad_cam": ["catia v5", "solidworks", "nxcam", "powermill"]
        },
        "hard_signals": ["inginer mecanic", "cad", "solidworks", "catia", "proiectare", "mecanica", "industrial", "autocad", "fmea"]
    },

    "AUTOMOTIVE_SERVICE": {
        "label": "Automotive Service",
        "subfamilies": {
            "auto_mechanic": ["mecanic auto", "auto mechanic", "tehnician mecanic auto", "mecanic autovehicule"],
            "assistant_mechanic": ["ajutor mecanic", "ucenic mecanic", "mecanic stagiar"],
            "auto_electrician": ["electrician auto", "auto electrician", "tehnician electric auto"],
            "diagnostics": ["diagnoza auto", "diagnostic auto", "specialist diagnoza"],
            "body_repair": ["tinichigiu auto", "body repair", "tinichigerie"],
            "auto_painter": ["vopsitor auto", "auto painter", "vopsitor caroserie"],
            "service_advisor": ["consilier service auto", "service advisor", "receptioner service auto"],
            "itp": ["inspector itp", "itp technician", "inspector tehnic"],
            "tire_service": ["vulcanizator", "tire technician", "specialist anvelope"],
            "auto_service_manager": ["sef service auto", "service manager auto"]
        },
        "hard_signals": ["service auto", "mecanic auto", "diagnoza", "revizii", "frane", "suspensie", "distributii", "benzina", "diesel", "tinichigerie", "vopsitorie auto", "itp"]
    },

    "ELECTRICAL_ENGINEERING": {
        "label": "Electrical Engineering",
        "subfamilies": {
            "electrician": ["electrician", "electrician constructii", "electrician industrial", "electrician autorizat"],
            "electrical_engineer": ["inginer electric", "electrical engineer", "inginer electrotehnic"],
            "automation_engineer": ["automation engineer", "inginer automatizari", "specialist scada"],
            "electronics": ["electronics engineer", "inginer electronist"],
            "panel_builder": ["tablotier", "electric panel builder", "montator tablouri electrice"],
            "low_voltage": ["low voltage specialist", "curenti slabi", "specialist curenti slabi"],
            "high_voltage": ["high voltage specialist", "inalta tensiune"],
            "plc": ["plc programmer", "programator plc", "specialist plc"],
            "electrical_design": ["electrical design engineer", "proiectant electric"],
            "energy_engineer": ["energy engineer", "inginer energetician", "auditor energetic"]
        },
        "hard_signals": ["electrician", "electrice", "automatizari", "plc", "tablou electric", "curenti slabi", "tensiune", "scada", "hmi"]
    },

    "CONSTRUCTION": {
        "label": "Construction",
        "subfamilies": {
            "civil_engineer": ["inginer constructii", "civil engineer", "inginer civil", "inginer drumuri poduri"],
            "site_manager": ["sef santier", "site manager", "manager proiect constructii"],
            "construction_worker": ["muncitor constructii", "construction worker", "muncitor calificat constructii"],
            "mason": ["zidar", "mason", "zidar armar"],
            "concrete_worker": ["betonist", "concrete worker", "operator beton"],
            "welder_construction": ["sudor constructii", "structural welder"],
            "scaffolding": ["montator schele", "scaffolder", "lucrator la inaltime"],
            "quantity_surveyor": ["devizier", "quantity surveyor", "inginer devize"],
            "building_inspector": ["inspector santier", "diriginte santier"],
            "construction_manager": ["director proiect constructii", "project manager constructii"]
        },
        "hard_signals": ["constructii", "santier", "zidarie", "beton", "armaturi", "cofraje", "devize", "sudura constructii"]
    },

    "PLUMBING_SANITARY": {
        "label": "Instalatii Sanitare & Termice",
        "subfamilies": {
            "plumber": ["instalator sanitar", "plumber", "instalator apa canal", "montator instalatii sanitare"],
            "heating_installer": ["instalator termic", "heating installer", "montator centrale termice", "instalator incalzire"],
            "gas_installer": ["instalator gaze", "gas installer", "montator gaze naturale"],
            "pipe_fitter": ["montator tevi", "pipe fitter", "instalator tevi industriale"],
            "bathroom_installer": ["montator baie", "bathroom fitter", "instalator obiecte sanitare"],
            "hvac_installer": ["montator hvac", "hvac installer"],
            "solar_installer": ["montator panouri solare termice", "solar thermal installer"],
            "underfloor_heating": ["montator incalzire pardoseala", "underfloor heating installer"],
            "water_treatment": ["operator statie apa", "water treatment technician"],
            "plumbing_supervisor": ["sef echipa instalatii", "supervizor instalatii sanitare"]
        },
        "hard_signals": ["instalator", "sanitar", "tevi", "centrale termice", "gaze", "apa canal", "instalatii termice", "coloane", "racorduri"]
    },

    "HVAC_CLIMATE": {
        "label": "Climatizare & Ventilatie",
        "subfamilies": {
            "hvac_technician": ["frigotehnist", "hvac technician", "tehnician climatizare", "specialist climatizare"],
            "ac_installer": ["montator aer conditionat", "ac installer", "montator split"],
            "ventilation": ["montator ventilatie", "ventilation installer", "specialist ventilatie"],
            "refrigeration": ["frigotehnist industrial", "refrigeration engineer", "specialist frig industrial"],
            "chiller_technician": ["tehnician chiller", "chiller specialist"],
            "heat_pump": ["montator pompe caldura", "heat pump installer"],
            "bms_technician": ["operator bms", "building management system technician"],
            "ahu_technician": ["tehnician ahu", "air handling unit specialist"],
            "commissioning": ["specialist punere in functiune hvac", "hvac commissioning engineer"],
            "hvac_manager": ["manager hvac", "sef echipa climatizare"]
        },
        "hard_signals": ["frigotehnist", "hvac", "climatizare", "aer conditionat", "ventilatie", "frig", "split", "chiller", "pompa caldura", "f-gaze"]
    },

    "WELDING": {
        "label": "Sudura & Lacatuserie",
        "subfamilies": {
            "mig_mag_welder": ["sudor mig mag", "mig mag welder", "sudor co2"],
            "tig_welder": ["sudor tig", "tig welder", "sudor wig"],
            "electric_welder": ["sudor electric", "electric arc welder", "sudor mma"],
            "gas_welder": ["sudor autogen", "gas welder", "sudor oxiacetilenic"],
            "pipe_welder": ["sudor tevi", "pipe welder", "sudor conducte"],
            "structural_welder": ["sudor constructii metalice", "structural steel welder"],
            "locksmith_industrial": ["lacatus mecanic", "locksmith industrial", "lacatus constructii metalice"],
            "metal_fabricator": ["confectioner metalic", "metal fabricator"],
            "welding_inspector": ["inspector sudura", "welding inspector", "specialist ndt"],
            "welding_engineer": ["inginer sudura", "welding engineer"]
        },
        "specializations": {
            "tig_welder": ["inox", "aluminiu", "titan"],
            "pipe_welder": ["5g", "6g", "presiune", "api"]
        },
        "hard_signals": ["sudor", "sudura", "mig", "mag", "tig", "wig", "lacatus", "confectii metalice", "autogen", "electrozi", "wps"]
    },

    "PAINTING_FINISHING": {
        "label": "Zugravit, Vopsit & Finisaje",
        "subfamilies": {
            "painter_interior": ["zugrav", "painter interior", "zugrav vopsitor", "decorator interior"],
            "industrial_painter": ["vopsitor industrial", "industrial painter"],
            "spray_painter": ["vopsitor pistol", "spray painter", "operator vopsire"],
            "floor_painter": ["vopsitor pardoseli", "floor coating applicator"],
            "powder_coating": ["operator vopsire pulbere", "powder coating operator"],
            "plaster_worker": ["tencuitor", "plasterer", "aplicator tencuiala"],
            "putty_worker": ["gletuit", "putty applicator", "aplicator glet"],
            "wallpaper": ["tapetator", "wallpaper hanger"],
            "epoxy_applicator": ["aplicator rasini epoxidice", "epoxy floor applicator"],
            "finishing_supervisor": ["sef echipa finisaje", "finishing supervisor"]
        },
        "hard_signals": ["zugrav", "vopsitor", "finisaje", "tencuiala", "glet", "vopsire", "tapete", "zugravituri"]
    },

    "FLOORING": {
        "label": "Pardoseli & Placari",
        "subfamilies": {
            "tiler": ["faiantar", "tiler", "montator faianta gresie", "placator"],
            "marble_granite": ["montator marmura granit", "marble installer"],
            "parquet_installer": ["parchetar", "parquet installer", "montator parchet"],
            "laminate_installer": ["montator stratificat", "laminate floor installer"],
            "carpet_installer": ["montator mocheta", "carpet installer"],
            "epoxy_flooring": ["montator pardoseli epoxidice", "industrial floor applicator"],
            "concrete_leveling": ["operator sapa", "concrete floor leveler", "aplicator sapa"],
            "linoleum": ["montator linoleum", "linoleum installer"],
            "outdoor_tiling": ["montator pavaje", "paving installer", "pavator"],
            "flooring_supervisor": ["sef echipa pardoseli"]
        },
        "hard_signals": ["faiantar", "gresie", "faianta", "parchetar", "pardoseli", "placari", "marmura", "pavaje", "sapa"]
    },

    "ROOFING_INSULATION": {
        "label": "Acoperisuri & Hidroizolatii",
        "subfamilies": {
            "roofer": ["acoperitor", "roofer", "montator invelitori", "specialist acoperisuri"],
            "tile_roofer": ["acoperitor tigla", "tile roofer"],
            "metal_roofer": ["acoperitor tabla", "metal roofer", "tablagar"],
            "waterproofing": ["hidroizolator", "waterproofing specialist", "montator membrane"],
            "thermal_insulation": ["montator termoizolatii", "thermal insulation installer", "izolator termic"],
            "acoustic_insulation": ["montator izolatii fonice", "acoustic insulation installer"],
            "green_roof": ["montator acoperis verde", "green roof installer"],
            "skylight_installer": ["montator luminatoare", "skylight installer"],
            "mansard": ["mansardar", "mansard specialist"],
            "roofing_supervisor": ["sef echipa acoperisuri"]
        },
        "hard_signals": ["acoperis", "invelitoare", "tigla", "hidroizolatie", "termoizolatie", "membrane", "izolatii", "mansarda"]
    },

    "FIRE_SAFETY_SYSTEMS": {
        "label": "Sisteme PSI & Securitate",
        "subfamilies": {
            "fire_alarm_technician": ["tehnician detectie incendiu", "fire alarm technician", "montator sisteme alarmare"],
            "sprinkler_technician": ["montator sprinklere", "sprinkler fitter"],
            "fire_suppression": ["montator sisteme stingere incendiu", "fire suppression technician"],
            "security_systems": ["montator sisteme securitate", "security systems installer", "montator alarma"],
            "cctv_technician": ["montator cctv", "cctv technician", "operator supraveghere video"],
            "access_control": ["montator control acces", "access control technician"],
            "psi_inspector": ["inspector psi", "fire safety inspector"],
            "intercom": ["montator interfonie", "intercom installer"],
            "bms_installer": ["montator bms", "building automation installer"],
            "security_systems_manager": ["manager sisteme securitate"]
        },
        "hard_signals": ["psi", "incendiu", "detectie incendiu", "sprinkler", "alarma", "securitate", "cctv", "supraveghere", "control acces"]
    },

    "ELEVATOR_ESCALATOR": {
        "label": "Lifturi & Escalatoare",
        "subfamilies": {
            "lift_technician": ["tehnician lifturi", "lift technician", "specialist lifturi", "electrician lifturi"],
            "lift_installer": ["montator lifturi", "lift installer", "elevator installer"],
            "escalator_technician": ["tehnician escalatoare", "escalator technician"],
            "lift_inspector": ["inspector lifturi", "lift inspector", "iscir inspector lifturi"],
            "lift_maintenance": ["mentenanta lifturi", "lift maintenance technician"],
            "freight_lift": ["tehnician lift marfa", "freight elevator specialist"],
            "hydraulic_lift": ["specialist lift hidraulic", "hydraulic elevator technician"],
            "lift_programmer": ["programator controller lift"],
            "lift_dispatcher": ["dispecer avarii lifturi"],
            "lift_manager": ["manager service lifturi"]
        },
        "hard_signals": ["lift", "ascensor", "escalator", "iscir", "montaj lift", "mentenanta lift"]
    },

    "TELECOM_NETWORKS": {
        "label": "Retele Telecom & Fibra Optica",
        "subfamilies": {
            "fiber_technician": ["tehnician fibra optica", "fiber technician", "montator fibra optica", "splicer fibra"],
            "telecom_technician": ["tehnician telecom", "telecoms technician", "specialist telecom"],
            "cable_installer": ["montator cabluri", "cable installer", "cablar"],
            "antenna_technician": ["tehnician antene", "antenna installer", "montator antene"],
            "5g_technician": ["tehnician 5g", "5g network technician"],
            "isp_technician": ["tehnician isp", "field technician internet", "tehnician instalare internet"],
            "structured_cabling": ["montator cablare structurata", "structured cabling specialist"],
            "radio_technician": ["tehnician radiocomunicatii", "radio technician"],
            "telecom_engineer": ["inginer telecom", "telecommunications engineer", "rf engineer"],
            "network_operations": ["operator noc telecom", "network operations telecom"]
        },
        "hard_signals": ["fibra optica", "telecom", "retele", "cablu", "antena", "5g", "internet", "cablare structurata", "splicer", "otdr"]
    },

    # ─── SANATATE, EDUCATIE, SERVICII ─────────────────────────────────────────

    "HOSPITALITY": {
        "label": "Hospitality",
        "subfamilies": {
            "hotel_reception": ["receptioner hotel", "hotel receptionist", "front desk hotel", "agent receptie hotel"],
            "hotel_manager": ["hotel manager", "manager hotel", "director hotel"],
            "housekeeping": ["camerista", "housekeeping", "guvernanta", "room attendant"],
            "reservation": ["agent rezervari", "reservation agent", "rezervari hotel"],
            "concierge": ["concierge", "portar hotel", "bell boy"],
            "guest_relations": ["guest relations", "relatii clienti hotel"],
            "night_auditor": ["night auditor", "auditor noapte hotel"],
            "spa_reception": ["spa receptionist", "receptioner spa"],
            "tourism_agent": ["agent turism", "travel agent", "consultant turism"],
            "hotel_supervisor": ["hotel supervisor", "sef receptie"]
        },
        "hard_signals": ["hotel", "receptie", "receptioner", "rezervari", "turism", "oaspeti", "guest", "hospitality", "accommodation"]
    },

    "FOOD_SERVICE": {
        "label": "Food Service",
        "subfamilies": {
            "cook": ["bucatar", "cook", "chef", "line cook", "bucatar profesionist"],
            "assistant_cook": ["ajutor bucatar", "kitchen helper", "kitchen porter"],
            "waiter": ["ospatar", "chelner", "waiter", "server", "sef de rang"],
            "bartender": ["barman", "bartender", "bar back"],
            "barista": ["barista", "specialist cafea"],
            "pastry": ["patiser", "pastry chef"],
            "confectioner": ["cofetar", "confectioner", "decorator torturi"],
            "kitchen_manager": ["sef bucatar", "head chef", "sous chef", "executive chef"],
            "dishwasher": ["spalator vase", "dishwasher", "kitchen cleaner"],
            "restaurant_manager": ["restaurant manager", "manager restaurant", "f&b manager"]
        },
        "hard_signals": ["restaurant", "bucatarie", "bucatar", "ospatar", "servire", "bar", "meniuri", "horeca", "preparate", "barman"]
    },

    "HEALTHCARE": {
        "label": "Healthcare",
        "subfamilies": {
            "nurse": ["asistent medical", "nurse", "asistent medical generalist", "asistent medical principal"],
            "doctor": ["medic", "doctor", "medic specialist", "medic rezident"],
            "pharmacist": ["farmacist", "pharmacist", "asistent farmacie"],
            "kinetotherapy": ["kinetoterapeut", "physiotherapist", "fizioterapeut"],
            "dentistry": ["asistent stomatologie", "dentist", "dental assistant", "medic dentist"],
            "medical_reception": ["receptioner clinica", "medical receptionist", "secretar medical"],
            "lab": ["laborant", "lab technician", "asistent laborator medical"],
            "caregiver": ["infirmier", "caregiver", "ingrijitor pacienti", "insotitor bolnavi"],
            "radiology": ["radiology technician", "radiolog", "operator radiologie"],
            "medical_manager": ["manager clinica", "director medical"]
        },
        "hard_signals": ["medical", "pacienti", "clinica", "spital", "asistent medical", "farmacie", "medic", "kineto", "tratament"]
    },

    "EDUCATION": {
        "label": "Education",
        "subfamilies": {
            "teacher": ["profesor", "teacher", "invatator"],
            "educator": ["educator", "educator puericultura"],
            "trainer": ["trainer", "formator", "instructor curs"],
            "university": ["lector universitar", "profesor universitar", "conferentiar"],
            "tutor": ["tutor", "meditator", "profesor particular"],
            "instructional_design": ["instructional designer", "e-learning developer"],
            "school_admin": ["secretar scoala", "administrator scoala"],
            "kindergarten": ["educatoare", "gradinita", "educator prescolar"],
            "language_teacher": ["profesor limbi straine", "teacher engleza"],
            "education_manager": ["director scoala", "manager educational", "inspector scolar"]
        },
        "hard_signals": ["profesor", "educatie", "predare", "cursuri", "training", "formator", "elevi", "studenti"]
    },

    "MANAGEMENT_LEADERSHIP": {
        "label": "Management & Leadership",
        "subfamilies": {
            "general_manager": ["general manager", "director general", "ceo", "managing director"],
            "operations_manager": ["operations manager", "manager operational", "director operational"],
            "project_manager": ["project manager", "manager proiect", "pmp", "scrum master"],
            "team_leader": ["team leader", "coordonator echipa", "lider echipa"],
            "department_manager": ["department manager", "sef departament", "manager departament"],
            "branch_manager": ["branch manager", "manager filiala", "manager regional"],
            "plant_manager": ["plant manager", "manager fabrica"],
            "program_manager": ["program manager", "program director"],
            "business_manager": ["business manager", "manager afaceri"],
            "executive": ["coo", "director executiv", "vicepresident", "chief officer"]
        },
        "hard_signals": ["manager", "director", "coordonare", "echipa", "buget", "strategie", "operational", "project management", "kpi", "p&l"]
    },

    # ─── SERVICII DIVERSE ─────────────────────────────────────────────────────

    "REAL_ESTATE": {
        "label": "Imobiliare",
        "subfamilies": {
            "real_estate_agent": ["agent imobiliar", "real estate agent", "broker imobiliar", "consultant imobiliar"],
            "property_manager": ["administrator imobile", "property manager", "facility manager imobiliar"],
            "real_estate_evaluator": ["evaluator imobiliar", "appraiser", "inspector imobiliar"],
            "rental_agent": ["agent inchirieri", "leasing agent imobiliar"],
            "developer_sales": ["agent vanzari dezvoltator", "consultant vanzari apartamente"],
            "property_admin": ["administrator bloc", "administrator proprietati"],
            "land_registry": ["cadastru", "inginer topograf", "specialist carte funciara"],
            "commercial_property": ["agent spatii comerciale", "commercial real estate"],
            "real_estate_manager": ["manager agentie imobiliara", "director vanzari imobiliare"],
            "real_estate_lawyer": ["jurist imobiliar"]
        },
        "hard_signals": ["imobiliar", "imobile", "apartament", "casa", "teren", "chirie", "vanzare imobil", "agentie imobiliara"]
    },

    "BEAUTY_WELLNESS": {
        "label": "Beauty & Wellness",
        "subfamilies": {
            "hairdresser": ["coafor", "hairdresser", "frizerie", "frizer", "stilist par"],
            "beautician": ["cosmeticiana", "beautician", "estetician", "beauty therapist"],
            "nail_tech": ["nail technician", "manichiurista", "nail artist", "unghii"],
            "massage_therapist": ["maseur", "massage therapist", "terapeut masaj"],
            "makeup_artist": ["makeup artist", "artist machiaj", "make-up artist"],
            "waxing_specialist": ["specialist epilare", "waxing specialist"],
            "eyelash_tech": ["specialist gene", "lash technician", "extensii gene"],
            "spa_therapist": ["terapeut spa", "spa therapist"],
            "tattoo_artist": ["tattoo artist", "tatuator"],
            "salon_manager": ["manager salon", "sef salon beauty"]
        },
        "hard_signals": ["coafor", "cosmetica", "beauty", "spa", "unghii", "masaj", "estetician", "salon", "frizerie", "epilare"]
    },

    "SECURITY_GUARD": {
        "label": "Securitate & Paza",
        "subfamilies": {
            "security_guard": ["agent securitate", "security guard", "paznic", "agent paza"],
            "bodyguard": ["bodyguard", "personal protection officer"],
            "security_supervisor": ["sef paza", "security supervisor", "dispecer securitate"],
            "cash_transport": ["agent transport valori", "cash in transit officer"],
            "event_security": ["agent securitate evenimente", "event security"],
            "mall_security": ["agent securitate mall", "mall security officer"],
            "airport_security": ["agent securitate aeroport", "airport security"],
            "private_investigator": ["investigator privat", "private investigator", "detectiv"],
            "security_manager": ["manager securitate", "director securitate"],
            "fire_warden": ["pompier privat", "agent psi", "responsabil evacuare"]
        },
        "hard_signals": ["securitate", "paza", "agent", "paznic", "atestat paza", "dispecer", "ronda", "bodyguard", "transport valori"]
    },

    "SOCIAL_SERVICES": {
        "label": "Asistenta Sociala & Psihologie",
        "subfamilies": {
            "social_worker": ["asistent social", "social worker", "lucrator social"],
            "psychologist": ["psiholog", "psychologist", "consilier psihologic"],
            "counselor": ["consilier", "counselor", "terapeut"],
            "child_protection": ["specialist protectia copilului", "child protection worker"],
            "community_worker": ["mediator comunitar", "community worker"],
            "disability_support": ["insotitor persoane cu dizabilitati", "disability support worker"],
            "elderly_care": ["ingrijitor batrani", "elderly care worker", "insotitor varstnici"],
            "probation_officer": ["consilier de probatiune", "probation officer"],
            "ngo_coordinator": ["coordonator proiecte sociale", "ngo program coordinator"],
            "social_manager": ["manager servicii sociale", "director centru social"]
        },
        "hard_signals": ["asistenta sociala", "psiholog", "consiliere", "social", "vulnerabil", "dizabilitati", "batrani", "copii", "dgaspc"]
    },

    "CLEANING_FACILITY": {
        "label": "Curatenie & Facility Management",
        "subfamilies": {
            "cleaning_worker": ["femeie serviciu", "om de serviciu", "cleaning worker", "lucrator curatenie", "muncitor curatenie"],
            "industrial_cleaner": ["agent curatenie industriala", "industrial cleaner", "operator curatenie"],
            "window_cleaner": ["spalator geamuri", "window cleaner"],
            "janitorial": ["ingrijitor", "janitor", "portar curatenie"],
            "facility_manager": ["facility manager", "manager facility", "administrator cladire"],
            "cleaning_supervisor": ["sef echipa curatenie", "cleaning supervisor"],
            "laundry": ["spalatoreasa", "laundry worker", "operator spalatorie"],
            "pest_control": ["dezinsectie deratizare", "pest control technician", "operator DDD"],
            "waste_management": ["gestionar deseuri", "waste management operator"],
            "cleaning_manager": ["manager servicii curatenie"]
        },
        "hard_signals": ["curatenie", "facility", "ingrijire", "menaj", "curatare", "dezinfectie", "DDD"]
    },

}

# ─── FAMILII ADIACENTE ────────────────────────────────────────────────────────
ADJACENT_FAMILIES: dict = {
    "IT_SOFTWARE": {"DATA_BI", "IT_INFRASTRUCTURE", "CYBERSECURITY", "IT_SUPPORT"},
    "IT_INFRASTRUCTURE": {"IT_SUPPORT", "CYBERSECURITY", "IT_SOFTWARE", "TELECOM_NETWORKS"},
    "IT_SUPPORT": {"IT_INFRASTRUCTURE", "ADMINISTRATIVE"},
    "DATA_BI": {"IT_SOFTWARE", "FINANCE_ACCOUNTING", "DIGITAL_MARKETING"},
    "CYBERSECURITY": {"IT_SOFTWARE", "IT_INFRASTRUCTURE"},
    "DIGITAL_MARKETING": {"MARKETING", "IT_SOFTWARE"},
    "MARKETING": {"DIGITAL_MARKETING", "SALES_B2B", "HR_RECRUITMENT"},
    "SALES_B2B": {"SALES_RETAIL", "MARKETING", "CUSTOMER_SUPPORT"},
    "SALES_RETAIL": {"SALES_B2B", "RETAIL_CASHIER", "CUSTOMER_SUPPORT"},
    "RETAIL_CASHIER": {"SALES_RETAIL", "CUSTOMER_SUPPORT"},
    "CUSTOMER_SUPPORT": {"SALES_B2B", "ADMINISTRATIVE"},
    "ADMINISTRATIVE": {"CUSTOMER_SUPPORT", "HR_RECRUITMENT"},
    "HR_RECRUITMENT": {"ADMINISTRATIVE", "MANAGEMENT_LEADERSHIP"},
    "FINANCE_ACCOUNTING": {"BANKING_INSURANCE", "ADMINISTRATIVE"},
    "BANKING_INSURANCE": {"FINANCE_ACCOUNTING", "LEGAL"},
    "LEGAL": {"ADMINISTRATIVE", "HR_RECRUITMENT"},
    "PROCUREMENT": {"SUPPLY_CHAIN", "LOGISTICS_WAREHOUSE"},
    "SUPPLY_CHAIN": {"PROCUREMENT", "LOGISTICS_WAREHOUSE", "TRANSPORTATION"},
    "LOGISTICS_WAREHOUSE": {"TRANSPORTATION", "SUPPLY_CHAIN", "PROCUREMENT"},
    "TRANSPORTATION": {"LOGISTICS_WAREHOUSE", "DANGEROUS_GOODS_ADR", "COLD_CHAIN"},
    "DANGEROUS_GOODS_ADR": {"TRANSPORTATION", "COLD_CHAIN"},
    "COLD_CHAIN": {"TRANSPORTATION", "LOGISTICS_WAREHOUSE"},
    "FREIGHT_FORWARDING": {"SUPPLY_CHAIN", "CUSTOMS_BROKERAGE"},
    "CUSTOMS_BROKERAGE": {"FREIGHT_FORWARDING", "SUPPLY_CHAIN"},
    "POSTAL_SERVICES": {"LOGISTICS_WAREHOUSE", "TRANSPORTATION"},
    "HEAVY_EQUIPMENT": {"CONSTRUCTION", "AGRICULTURE_FARMING"},
    "RAIL_TRANSPORT": {"TRANSPORTATION"},
    "MARITIME_TRANSPORT": {"TRANSPORTATION", "FREIGHT_FORWARDING"},
    "AVIATION_GROUND": {"TRANSPORTATION", "FREIGHT_FORWARDING"},
    "PRODUCTION_MANUFACTURING": {"MAINTENANCE_TECHNICAL", "MECHANICAL_ENGINEERING"},
    "FOOD_PRODUCTION": {"PRODUCTION_MANUFACTURING", "FOOD_SERVICE"},
    "TEXTILE_GARMENT": {"PRODUCTION_MANUFACTURING"},
    "WOOD_FURNITURE_PRODUCTION": {"PRODUCTION_MANUFACTURING", "CONSTRUCTION"},
    "PRINTING_PACKAGING": {"PRODUCTION_MANUFACTURING"},
    "CHEMICAL_PROCESS": {"PRODUCTION_MANUFACTURING"},
    "PLASTICS_RUBBER": {"PRODUCTION_MANUFACTURING", "CHEMICAL_PROCESS"},
    "AGRICULTURE_FARMING": {"HEAVY_EQUIPMENT", "FOOD_PRODUCTION"},
    "MAINTENANCE_TECHNICAL": {"AUTOMOTIVE_SERVICE", "ELECTRICAL_ENGINEERING", "PRODUCTION_MANUFACTURING", "MECHANICAL_ENGINEERING"},
    "MECHANICAL_ENGINEERING": {"MAINTENANCE_TECHNICAL", "AUTOMOTIVE_SERVICE", "PRODUCTION_MANUFACTURING"},
    "AUTOMOTIVE_SERVICE": {"MAINTENANCE_TECHNICAL", "MECHANICAL_ENGINEERING"},
    "ELECTRICAL_ENGINEERING": {"MAINTENANCE_TECHNICAL", "CONSTRUCTION", "TELECOM_NETWORKS"},
    "CONSTRUCTION": {"PLUMBING_SANITARY", "ELECTRICAL_ENGINEERING", "WELDING", "HEAVY_EQUIPMENT"},
    "PLUMBING_SANITARY": {"CONSTRUCTION", "HVAC_CLIMATE"},
    "HVAC_CLIMATE": {"PLUMBING_SANITARY", "ELECTRICAL_ENGINEERING", "MAINTENANCE_TECHNICAL"},
    "WELDING": {"CONSTRUCTION", "MECHANICAL_ENGINEERING", "MAINTENANCE_TECHNICAL"},
    "PAINTING_FINISHING": {"CONSTRUCTION", "FLOORING"},
    "FLOORING": {"CONSTRUCTION", "PAINTING_FINISHING"},
    "ROOFING_INSULATION": {"CONSTRUCTION"},
    "FIRE_SAFETY_SYSTEMS": {"ELECTRICAL_ENGINEERING", "SECURITY_GUARD"},
    "ELEVATOR_ESCALATOR": {"ELECTRICAL_ENGINEERING", "MAINTENANCE_TECHNICAL"},
    "TELECOM_NETWORKS": {"IT_INFRASTRUCTURE", "ELECTRICAL_ENGINEERING"},
    "HOSPITALITY": {"FOOD_SERVICE", "CUSTOMER_SUPPORT", "REAL_ESTATE"},
    "FOOD_SERVICE": {"HOSPITALITY", "FOOD_PRODUCTION"},
    "HEALTHCARE": {"SOCIAL_SERVICES"},
    "EDUCATION": {"SOCIAL_SERVICES", "HR_RECRUITMENT"},
    "REAL_ESTATE": {"ADMINISTRATIVE", "LEGAL", "CONSTRUCTION"},
    "BEAUTY_WELLNESS": {"HOSPITALITY", "HEALTHCARE"},
    "SECURITY_GUARD": {"FIRE_SAFETY_SYSTEMS", "TRANSPORTATION"},
    "SOCIAL_SERVICES": {"HEALTHCARE", "EDUCATION"},
    "CLEANING_FACILITY": {"ADMINISTRATIVE", "HOSPITALITY"},
    "MANAGEMENT_LEADERSHIP": {"HR_RECRUITMENT", "ADMINISTRATIVE"},
}

# Alias pentru backward compatibility
CUSTOMS_BROKERAGE = "FREIGHT_FORWARDING"
