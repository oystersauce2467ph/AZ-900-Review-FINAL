# -*- coding: utf-8 -*-
"""
Build an AZ-900 Kahoot-importable .xlsx from the 3 reviewer PDFs.

Sources (all content STRICTLY from these files, no invented facts/answers):
  P1 = "AZ-900 Fundamentals Reviewer.pdf"            (dense fundamentals notes)
  P2 = "AZ900_Practice_Exam_With_Answers.pdf"        (Q1-Q36 with answers)
  P3 = "RADOVAN_AZ-900_ Microsoft Azure Fundamentals Reviewer.pdf" (domain + practice Qs)
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = "AZ-900_Kahoot_Import.xlsx"

# ---------- styling helpers ----------
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, color="1F4E78", size=14)
NOTE_FONT = Font(italic=True, color="555555", size=10)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = BORDER

def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


wb = Workbook()

# =====================================================================
# SHEET 1: READ ME / INSTRUCTIONS
# =====================================================================
ws = wb.active
ws.title = "READ ME"
set_widths(ws, [4, 120])
ws["B1"] = "AZ-900 Kahoot Question Bank - Built from your 3 reviewer PDFs"
ws["B1"].font = TITLE_FONT
lines = [
    "",
    "SOURCES (every question & answer comes strictly from these files - nothing invented):",
    "  - AZ-900 Fundamentals Reviewer.pdf            (referred to below as P1)",
    "  - AZ900_Practice_Exam_With_Answers.pdf        (referred to below as P2)",
    "  - RADOVAN_AZ-900_ Microsoft Azure Fundamentals Reviewer.pdf  (referred to below as P3)",
    "",
    "WORKBOOK TABS:",
    "  1. READ ME                - this sheet.",
    "  2. Multiple Choice        - questions that have 4 (or 2) lettered options. Kahoot import format.",
    "  3. True or False          - statements answered True/False. Kahoot import format (2 options).",
    "  4. Use Cases (Type Answer)- scenario / fill-in / definition questions. Player TYPES the answer.",
    "  5. Redundant Questions    - duplicate / overlapping questions to AVOID repeating.",
    "  6. Source Map             - quick legend of which PDF each block came from.",
    "",
    "HOW TO IMPORT INTO KAHOOT:",
    "  - 'Multiple Choice' and 'True or False' tabs match Kahoot's spreadsheet template",
    "    (Question | Answer 1..4 | Time limit | Correct answer(s)). In Kahoot, create a new Kahoot,",
    "    choose 'Add question' > 'Import questions from spreadsheet', and upload the relevant tab",
    "    (copy the single tab into Kahoot's official template if the importer asks for it).",
    "  - 'Correct answer(s)' uses option NUMBERS: 1=Answer 1, 2=Answer 2, etc.",
    "    For True/False: Answer 1 = True, Answer 2 = False.",
    "  - 'Use Cases (Type Answer)' must be added as Kahoot 'Type Answer' questions. The bulk",
    "    spreadsheet importer only supports quiz/true-false, so add these manually (or via the",
    "    Type Answer template). The correct answer from the PDF is provided so YOU can type it in.",
    "",
    "KAHOOT LIMITS TO REMEMBER:",
    "  - Question text: max 120 characters.   - Quiz answer option: max 75 characters.",
    "  - Type Answer accepted answer: max ~20 characters (add short accepted forms, e.g. 'RBAC').",
    "    A 'Suggested Short Answer' column is provided where the full answer is long.",
    "  - Allowed time limits (sec): 5, 10, 20, 30, 60, 90, 120, 240. Default used here: 30.",
    "",
    "INTEGRITY NOTE:",
    "  - True/False items come VERBATIM from the Yes/No statements in P2 (Yes=True, No=False).",
    "  - Multiple-choice items use ONLY the options that already exist in the PDFs (no invented distractors).",
    "  - Use-case / definition items are grounded in a specific PDF statement (see the 'Source' column).",
    "  - See the 'Redundant Questions' tab before building your game to avoid asking the same thing twice.",
]
r = 2
for ln in lines:
    ws.cell(row=r, column=2, value=ln)
    if ln.endswith(":") or ln.startswith(("SOURCES", "WORKBOOK", "HOW TO", "KAHOOT LIMITS", "INTEGRITY")):
        ws.cell(row=r, column=2).font = Font(bold=True, size=11, color="1F4E78")
    r += 1
ws.sheet_view.showGridLines = False

# =====================================================================
# SHEET 2: MULTIPLE CHOICE  (Kahoot import format)
# =====================================================================
# Each row: question, a1, a2, a3, a4, time, correct(number), source, note
mc = [
    # P2 Q2
    ("Each of your 10 departments needs a different payment option for Azure. What should you create for each department?",
     "A resource group", "A subscription", "A container instance", "A reservation", 30, "2", "P2 Q2", ""),
    # P2 Q4
    ("What is used to grant permission to Azure Virtual Desktop resources?",
     "Role-based access control (RBAC) roles", "Tags", "Resource groups", "Application security groups", 30, "1", "P2 Q4", ""),
    # P2 Q6
    ("You plan to create a virtual machine in Azure. Where will the virtual machine be placed?",
     "In a storage account", "In an administrative unit", "In a resource group", "In an application group", 30, "3", "P2 Q6", ""),
    # P2 Q10
    ("How many copies of data are kept by an Azure Storage account that uses locally redundant storage (LRS)?",
     "3", "4", "6", "9", 30, "1", "P2 Q10", ""),
    # P3 practice test Q14 (2 options only - given in source)
    ("Which pairing counts as TRUE multi-factor authentication (MFA)?",
     "Password + fingerprint", "Fingerprint + retina scan", None, None, 30, "1", "P3 Practice Test Q14",
     "Only 2 options exist in the source (two of the same factor type do not qualify as MFA)."),
]

ws = wb.create_sheet("Multiple Choice")
headers = ["Question", "Answer 1", "Answer 2", "Answer 3", "Answer 4",
           "Time limit (sec)", "Correct answer(s)", "Source (PDF)", "Notes"]
ws.append(headers)
style_header(ws, 1, len(headers))
for row in mc:
    ws.append(list(row))
set_widths(ws, [60, 28, 28, 26, 26, 12, 14, 16, 45])
for r_ in range(2, ws.max_row + 1):
    for c_ in range(1, len(headers) + 1):
        ws.cell(row=r_, column=c_).alignment = WRAP
        ws.cell(row=r_, column=c_).border = BORDER
    ws.cell(row=r_, column=6).alignment = CENTER
    ws.cell(row=r_, column=7).alignment = CENTER
ws.freeze_panes = "A2"

# =====================================================================
# SHEET 3: TRUE OR FALSE  (Kahoot import format; A1=True, A2=False)
# =====================================================================
# (statement, correct: "1"=True / "2"=False, source)
tf = [
    # P2 Q1 (Yes/Yes/Yes)
    ("You can create custom Azure roles to control access to resources.", "1", "P2 Q1"),
    ("A user account can be assigned to multiple Azure roles.", "1", "P2 Q1"),
    ("A resource group can have the Owner role assigned to multiple users.", "1", "P2 Q1"),
    # P2 Q3 (Yes/No/No)
    ("An Azure subscription can have multiple account administrators.", "1", "P2 Q3"),
    ("An Azure subscription can be managed by using a Microsoft account only.", "2", "P2 Q3"),
    ("An Azure resource group can contain multiple Azure subscriptions.", "2", "P2 Q3"),
    # P2 Q5 (No/No/Yes)
    ("Single sign-on (SSO) requires that all users sign in by using the Microsoft Authenticator app.", "2", "P2 Q5"),
    ("Authentication is the process of establishing which level of access an authenticated user has.", "2", "P2 Q5"),
    ("Conditional Access uses signals collected during sign-in to allow or deny access.", "1", "P2 Q5"),
    # P2 Q8 (Yes/No/No)
    ("Data stored in an Azure Storage account automatically has at least three copies.", "1", "P2 Q8"),
    ("All data copied to an Azure Storage account is backed up automatically to another Azure datacenter.", "2", "P2 Q8"),
    ("An Azure Storage account can contain only up to 2 TB of data and one million files.", "2", "P2 Q8"),
    # P2 Q17 (Yes/No/Yes)
    ("Azure Reservations cost less than pay-as-you-go pricing.", "1", "P2 Q17"),
    ("Two Azure virtual machines of the same size always have the same monthly cost.", "2", "P2 Q17"),
    ("When an Azure virtual machine is stopped, storage costs still apply.", "1", "P2 Q17"),
    # P2 Q20 (Yes/No/Yes)
    ("Azure Files is an example of Infrastructure as a Service (IaaS).", "1", "P2 Q20"),
    ("A DNS server running on an Azure virtual machine is an example of Platform as a Service (PaaS).", "2", "P2 Q20"),
    ("Microsoft Intune is an example of Software as a Service (SaaS).", "1", "P2 Q20"),
    # P2 Q23 (No/Yes/Yes)
    ("Creating and configuring a virtual network is part of the PaaS model.", "2", "P2 Q23"),
    ("Updating application code in Azure App Service is the customer's responsibility.", "1", "P2 Q23"),
    ("Configuring user access in PaaS is the customer's responsibility.", "1", "P2 Q23"),
    # P2 Q24 (No/Yes/Yes)
    ("A resource lock can be applied to a Microsoft Entra user.", "2", "P2 Q24"),
    ("Multiple resource locks can be applied to the same virtual machine.", "1", "P2 Q24"),
    ("A delete lock allows modification of a resource but prevents deletion.", "1", "P2 Q24"),
    # P2 Q26 (Yes/Yes/No)
    ("Azure PowerShell can be installed on macOS.", "1", "P2 Q26"),
    ("Azure Cloud Shell can be accessed from a Linux computer using a browser.", "1", "P2 Q26"),
    ("The Azure portal can be accessed only from Windows devices.", "2", "P2 Q26"),
    # P2 Q27 (Yes/No/Yes)
    ("Inbound traffic to Azure using ExpressRoute is always free.", "1", "P2 Q27"),
    ("Outbound traffic from Azure to on-premises networks is always free.", "2", "P2 Q27"),
    ("Data transfer between Azure services in the same region is always free.", "1", "P2 Q27"),
    # P2 Q29 (Yes/Yes/Yes)
    ("Azure Monitor can monitor the performance of on-premises computers.", "1", "P2 Q29"),
    ("Azure Monitor can send alerts to Microsoft Entra security groups.", "1", "P2 Q29"),
    ("Azure Monitor can trigger alerts based on Log Analytics data.", "1", "P2 Q29"),
    # P2 Q30 (No/Yes/No)
    ("Only one tag can be assigned to an Azure resource.", "2", "P2 Q30"),
    ("Tags can be assigned using ARM templates.", "1", "P2 Q30"),
    ("Tags can be used to enforce naming standards.", "2", "P2 Q30"),
]

ws = wb.create_sheet("True or False")
headers = ["Question", "Answer 1", "Answer 2", "Answer 3", "Answer 4",
           "Time limit (sec)", "Correct answer(s)", "Source (PDF)"]
ws.append(headers)
style_header(ws, 1, len(headers))
for stmt, correct, src in tf:
    ws.append([stmt, "True", "False", None, None, 20, correct, src])
set_widths(ws, [78, 12, 12, 10, 10, 12, 14, 12])
for r_ in range(2, ws.max_row + 1):
    for c_ in range(1, len(headers) + 1):
        ws.cell(row=r_, column=c_).alignment = WRAP
        ws.cell(row=r_, column=c_).border = BORDER
    ws.cell(row=r_, column=6).alignment = CENTER
    ws.cell(row=r_, column=7).alignment = CENTER
ws.freeze_panes = "A2"

# =====================================================================
# SHEET 4: USE CASES (TYPE ANSWER)
# =====================================================================
# (Topic, Question, Full answer from PDF, Suggested short answer (<=20 chars or ""), Source)
ta = []

# ---- P2 scenario / fill-in / matching / short-answer ----
ta += [
 ("Networking", "Which connection gives a private, dedicated link from on-premises to Azure with NO public internet?", "ExpressRoute", "ExpressRoute", "P2 Q7"),
 ("Networking", "Which service connects two Azure VNets with low-latency, high-bandwidth?", "Virtual network peering", "VNet peering", "P2 Q7"),
 ("Networking", "Which service provides an encrypted tunnel over the public internet to Azure?", "VPN gateway", "VPN gateway", "P2 Q7"),
 ("Networking", "When extending your on-premises network to Azure, which resource defines the on-premises VPN device?", "Local Network Gateway", "Local Net Gateway", "P2 Q9"),
 ("Security", "Requiring a password AND a one-time passcode at sign-in is an example of which security feature?", "Multi-Factor Authentication (MFA)", "MFA", "P2 Q11"),
 ("Shared Responsibility", "In IaaS, name the two components that are the responsibility of the cloud provider.", "Physical hardware (servers/network) and physical datacenter security", "", "P2 Q12"),
 ("Compute", "Which service provides a full Windows desktop experience from the cloud?", "Azure Virtual Desktop", "AVD", "P2 Q13"),
 ("Cloud Benefits", "____ ensures access to cloud resources in the event of a service failure.", "High availability", "High availability", "P2 Q14"),
 ("Shared Responsibility", "What does a customer provide in a SaaS model?", "Data and user access (accounts/identities)", "Data and accounts", "P2 Q15"),
 ("Shared Responsibility", "In SaaS, which responsibility is shared between Microsoft and the customer?", "Identity and access management", "Identity/access mgmt", "P2 Q16"),
 ("Service Models", "You must migrate an on-premises server using a lift-and-shift approach. Which cloud service model?", "IaaS (Infrastructure as a Service)", "IaaS", "P2 Q18"),
 ("Cloud Benefits", "Which cloud benefit lets you recover from a catastrophic failure?", "Disaster recovery", "Disaster recovery", "P2 Q19"),
 ("Cloud Benefits", "Which cloud benefit deploys resources close to users globally?", "Geo-distribution", "Geo-distribution", "P2 Q19"),
 ("Cloud Benefits", "Which cloud benefit adjusts resources to meet changing demand?", "Scalability", "Scalability", "P2 Q19"),
 ("Cloud Benefits", "____ enables Azure resources to be deployed close to users.", "Geo-distribution (Azure Regions)", "Geo-distribution", "P2 Q21"),
 ("Cost", "When migrating a public website to Azure, you must plan to ____.", "Pay monthly usage costs (OpEx / consumption-based pricing)", "Pay monthly (OpEx)", "P2 Q22"),
 ("Governance", "You need to ensure resources can be created only in specific Azure regions. What should you use?", "Azure Policy", "Azure Policy", "P2 Q25"),
 ("Management", "Azure Resource Manager (ARM) templates use which format?", "JSON", "JSON", "P2 Q28"),
 ("Governance", "You can assign ____ to every Azure resource.", "Tags", "Tags", "P2 Q31"),
 ("Management", "What can you use to identify underutilized or unused Azure virtual machines?", "Azure Advisor", "Azure Advisor", "P2 Q32"),
 ("Management", "Which tool do you use to monitor the health of Azure services?", "Azure Service Health", "Service Health", "P2 Q33"),
 ("Security", "Which tool shows security recommendations?", "Microsoft Defender for Cloud", "Defender for Cloud", "P2 Q33"),
 ("Governance", "You need to prevent accidental deletion of resources in a resource group. Which setting should you use?", "Delete lock (CanNotDelete resource lock)", "Delete lock", "P2 Q34"),
 ("Hybrid", "____ extends Azure compliance and monitoring to hybrid and multicloud environments.", "Azure Arc", "Azure Arc", "P2 Q35"),
 ("Management", "Put in order the steps to deploy 5 identical VMs from an existing VM.", "1) Generalize the VM  2) Create a VM image  3) Create an ARM template referencing the image  4) Deploy 5 VMs", "", "P2 Q36"),
]

# ---- P3 Domain practice questions + practice test ----
ta += [
 ("Cloud Models", "Which cloud model is considered the most flexible?", "Hybrid", "Hybrid", "P3 Domain 1 Q1"),
 ("Cost", "Which cloud expenditure model means paying for resources only as they are used?", "OpEx / Consumption-based", "OpEx", "P3 Domain 1 Q2"),
 ("Cloud Benefits", "What term describes the automatic addition of resources based on demand?", "Elasticity", "Elasticity", "P3 Domain 1 Q3"),
 ("Service Models", "Microsoft 365 Online falls under which cloud service category?", "SaaS (Software as a Service)", "SaaS", "P3 Domain 1 Q4"),
 ("Cloud Models", "Which cloud model gives organizations total control over security and resources?", "Private", "Private", "P3 Domain 1 Q5"),
 ("Architecture", "What is the minimum number of Availability Zones in a region that supports them?", "3", "3", "P3 Domain 2 Q1"),
 ("Storage", "Which storage tier is for data stored at least 180 days and rarely accessed?", "Archive", "Archive", "P3 Domain 2 Q2"),
 ("Networking", "Which Azure service is a Layer 7 load balancer specifically for web applications?", "Azure Application Gateway", "App Gateway", "P3 Domain 2 Q3"),
 ("Architecture", "If a resource group is deleted, what happens to the resources inside it?", "They are deleted", "Deleted", "P3 Domain 2 Q4"),
 ("Storage", "What physical tool transfers massive amounts of data (e.g., 40 TB) to Azure?", "Azure Data Box", "Data Box", "P3 Domain 2 Q5"),
 ("Cost", "Which pricing model gives the largest discount for VMs that can be decommissioned anytime?", "Spot Pricing / Spot Instances", "Spot", "P3 Domain 3 Q1"),
 ("Compliance", "Which website provides Microsoft's audit reports and compliance info (GDPR/ISO)?", "Service Trust Portal", "Service Trust Portal", "P3 Domain 3 Q2"),
 ("Governance", "What is the name of a collection of Azure policies grouped together?", "An Initiative", "Initiative", "P3 Domain 3 Q3"),
 ("Shared Responsibility", "Under the Shared Responsibility Model, what is the cloud provider always responsible for?", "The physical network / physical datacenters", "Physical network", "P3 Domain 3 Q4"),
 ("Management", "Which Azure tool analyzes infrastructure and recommends for cost, security, reliability, performance?", "Azure Advisor", "Azure Advisor", "P3 Domain 3 Q5"),
 ("Cloud Models", "An admin keeps AD on-premises but hosts a database in the cloud (on-prem can't handle it). Which model?", "Hybrid cloud", "Hybrid", "P3 Practice Test Q1"),
 ("Service Models", "Which cloud service type provides a development environment without managing the OS?", "PaaS (Platform as a Service)", "PaaS", "P3 Practice Test Q2"),
 ("Networking", "Which networking service carries data to Microsoft's datacenters WITHOUT the public internet?", "ExpressRoute", "ExpressRoute", "P3 Practice Test Q3"),
 ("Governance", "An admin needs to restrict which regions people can deploy resources into. What should they set?", "Azure Policy", "Azure Policy", "P3 Practice Test Q4"),
 ("Storage", "Which storage redundancy option offers 16 nines durability by copying to a secondary geographic region?", "Geo-Redundant Storage (GRS)", "GRS", "P3 Practice Test Q5"),
 ("Identity", "A user logs into Azure with a username and password. Which identity process just occurred?", "Authentication", "Authentication", "P3 Practice Test Q6"),
 ("Architecture", "What happens to a resource's permissions when it is moved into a different resource group?", "It inherits the new resource group's permissions", "Inherits new RG perms", "P3 Practice Test Q7"),
 ("Management", "What file format is used to build Azure Resource Manager (ARM) templates?", "JSON (Bicep also supported)", "JSON", "P3 Practice Test Q8"),
 ("Cost", "A developer wants the cheapest VM for an experiment that can be shut down anytime. Which pricing option?", "Spot pricing / Spot Instances", "Spot", "P3 Practice Test Q9"),
 ("Storage", "Which two factors does Azure use to price storage disks?", "Quality (SSD vs HDD) and size", "Quality and size", "P3 Practice Test Q10"),
 ("Cloud Models", "A business wants total control of its own hardware/security for a single-org app. Which model fits best?", "Private cloud", "Private", "P3 Practice Test Q11"),
 ("Management", "Which Azure tool should you check FIRST when investigating a suspected Azure-wide outage?", "Azure Service Health", "Service Health", "P3 Practice Test Q12"),
 ("Networking", "A school must isolate network traffic for each classroom within one VNet. What should be created?", "Virtual subnets", "Subnets", "P3 Practice Test Q13"),
 ("Security", "What is the minimum subscription tier required to use Microsoft Defender for Cloud?", "Free", "Free", "P3 Practice Test Q15"),
]

# ---- P1 fundamentals (definitions/use cases grounded in the notes) ----
ta += [
 ("Cloud Concepts", "Define cloud computing (per the reviewer).", "Use of compute, network, and storage resources over the Internet under a consumption-based model", "", "P1 p1"),
 ("Service Models", "Which use cases are listed for IaaS?", "Lift-and-shift migration; testing and deployment", "", "P1 p1"),
 ("Service Models", "Which use cases are listed for PaaS?", "Development framework; analytics and business intelligence", "", "P1 p1"),
 ("Service Models", "Which use cases are listed for SaaS?", "Messaging; productivity; finance/expense tracking", "", "P1 p1"),
 ("Shared Responsibility", "Name the 3 items ALWAYS the responsibility of the company.", "Information/data; devices (mobile/PC); accounts and identities", "", "P1 p1"),
 ("Shared Responsibility", "Name the 3 items ALWAYS the responsibility of the cloud provider.", "Physical hosts; physical network; physical data center", "", "P1 p1"),
 ("Cost", "What is Capital Expenditure (CapEx)?", "One-time upfront costs", "Upfront costs", "P1 p1"),
 ("Cost", "What is Operational Expenditure (OpEx)?", "Spending money over time (Pay-As-You-Go)", "Pay-as-you-go", "P1 p1"),
 ("SLA", "How is an Azure SLA expressed?", "As a percentage (%) of uptime", "% of uptime", "P1 p1"),
 ("SLA", "Per the reviewer, 99% uptime equals how much downtime per month?", "7.2 hours per month (1.68 hours per week)", "7.2 hrs/month", "P1 p1"),
 ("Cloud Benefits", "Which benefit means resources are accessible during failures and updates?", "High Availability", "High Availability", "P1 p1"),
 ("Cloud Benefits", "Which benefit increases/decreases resource capacity based on workload?", "Scalability", "Scalability", "P1 p1"),
 ("Scaling", "What is vertical scaling (scale up/down)?", "Increasing/decreasing resource capabilities (e.g., VM CPU/RAM)", "", "P1 p1"),
 ("Scaling", "What is horizontal scaling (scale in/out)?", "Adding/removing resource instances (more VMs/containers)", "", "P1 p1"),
 ("Cloud Benefits", "Which benefit is automatic scalability?", "Elasticity", "Elasticity", "P1 p1"),
 ("Cloud Benefits", "Which benefit means recovering after system failures?", "Reliability", "Reliability", "P1 p1"),
 ("Cloud Benefits", "Which benefit covers forecasting performance or costs?", "Predictability", "Predictability", "P1 p1"),
 ("Cloud Benefits", "Which benefit covers data encryption, identity & access management, and threat monitoring?", "Security", "Security", "P1 p1"),
 ("Cloud Benefits", "Which benefit involves auditing and flagging of non-compliant resources?", "Compliance", "Compliance", "P1 p1"),
 ("Cloud Benefits", "Which benefit uses predefined templates for corporate/government standards?", "Governance", "Governance", "P1 p1"),
 ("Cloud Benefits", "Which benefit covers managing/deploying resources, tracking, and alerts?", "Manageability", "Manageability", "P1 p1"),
 ("Cloud Benefits", "Which benefit reduces an organization's carbon footprint via economies of scale?", "Sustainability", "Sustainability", "P1 p1"),
 ("Geography", "What are Azure availability regions?", "Geographical areas with data centers (60+ regions in 140+ countries)", "", "P1 p2"),
 ("Geography", "Which services are NOT region-locked?", "Microsoft Entra ID, Azure Traffic Manager, Azure DNS", "", "P1 p2"),
 ("Geography", "What are availability zones?", "Physically separated datacenters in a region with independent power, cooling and networking", "", "P1 p2"),
 ("Geography", "Name the 3 categories of services regarding availability zones.", "Zonal; Zone-redundant; Non-regional", "", "P1 p2"),
 ("Geography", "What is the separation distance for region pairs?", "482 km / 300 miles", "300 miles", "P1 p2"),
 ("Geography", "What are sovereign regions used for?", "Isolation from the main instance for compliance or legal purposes", "Compliance/legal", "P1 p2"),
 ("Compute", "What does a VM Scale Set do?", "Creates/manages a group of identical, load-balanced VMs that scale with demand", "", "P1 p2"),
 ("Compute", "What do VM availability sets group VMs by?", "Update domains and fault domains", "", "P1 p2"),
 ("Compute", "What is an update domain?", "VMs that can be rebooted together during planned maintenance", "", "P1 p2"),
 ("Compute", "What is a fault domain?", "VMs that share a potential power or network failure point", "", "P1 p2"),
 ("Compute", "What does Azure Load Balancer do?", "Distributes web traffic across VMs for improved performance", "", "P1 p2"),
 ("Architecture", "What is the basic building block of Azure?", "A resource", "A resource", "P1 p2"),
 ("Architecture", "Every resource must belong to how many resource groups?", "Exactly one", "One", "P1 p2"),
 ("Architecture", "Can resource groups be nested?", "No", "No", "P1 p2"),
 ("Architecture", "Can a resource group be renamed after creation?", "No", "No", "P1 p2"),
 ("Architecture", "How many management groups can a single directory support, and to what depth?", "10,000 management groups, up to six levels of depth", "10000 / 6 levels", "P1 p2"),
 ("Architecture", "Name the two subscription boundary types.", "Billing and Access control", "Billing, Access", "P1 p2"),
 ("Compute", "Which service is a container orchestration service for many containers?", "Azure Kubernetes Service (AKS)", "AKS", "P1 p3"),
 ("Compute", "What is Azure Functions?", "Event-driven serverless code that runs only when triggered", "", "P1 p3"),
 ("Compute", "Azure Functions runs in response to which triggers?", "HTTP requests, Timers, and Azure Service Messages", "", "P1 p3"),
 ("Compute", "Name the 4 Azure App Service types.", "Web Apps, API Apps, WebJobs, Mobile Apps", "", "P1 p3"),
 ("IoT", "Which IoT service enables secure bi-directional communication between cloud and devices?", "Azure IoT Hub", "IoT Hub", "P1 p3"),
 ("IoT", "Which IoT service is a SaaS IoT platform for solution builders?", "Azure IoT Central", "IoT Central", "P1 p3"),
 ("IoT", "Which IoT service extends cloud capabilities to edge devices?", "Azure IoT Edge", "IoT Edge", "P1 p3"),
 ("AI", "Which Azure service supports generative AI like chat and content generation?", "Azure OpenAI Service", "Azure OpenAI", "P1 p3"),
 ("Networking", "What are Network Security Groups (NSGs)?", "Azure resources with inbound/outbound rules to allow/block traffic by IP, port, protocol", "", "P1 p3"),
 ("Networking", "What is a network virtual appliance?", "A specialized VM that performs a network function such as a firewall or WAN optimization", "", "P1 p3"),
 ("Networking", "Which ExpressRoute connection type goes straight to the Microsoft backbone at 100 Gbps?", "ExpressRoute Direct", "ExpressRoute Direct", "P1 p4"),
 ("Networking", "What does the Azure VPN Gateway use to connect networks?", "The public Internet", "Public Internet", "P1 p4"),
 ("Networking", "Point-to-Site vs Site-to-Site VPN - what is the difference?", "Point-to-Site = single device; Site-to-Site = entire networks", "", "P1 p3"),
 ("Networking", "What is the DEFAULT VPN high-availability option?", "Active/Standby", "Active/Standby", "P1 p4"),
 ("Networking", "What are alias records?", "Records that automatically update DNS records when linked Azure resources change", "", "P1 p4"),
 ("Storage", "What are the Azure storage account naming rules?", "3-24 characters, lowercase letters and numbers only, globally unique", "", "P1 p4"),
 ("Storage", "Which storage account type is the default for most scenarios?", "Standard general-purpose v2", "Standard GPv2", "P1 p4"),
 ("Storage", "LRS keeps how many copies and where?", "3 copies in 1 datacenter (at least 11 nines durability)", "3 copies, 1 DC", "P1 p4"),
 ("Storage", "ZRS keeps how many copies and where?", "3 copies across multiple availability zones (at least 12 nines)", "3 copies, zones", "P1 p4"),
 ("Storage", "GRS keeps how many copies and where?", "3 copies locally + 3 in a paired region (at least 16 nines)", "3 local + 3 paired", "P1 p4"),
 ("Storage", "What does RA-GRS add?", "Read access to the secondary region", "Read access", "P1 p4"),
 ("Storage", "Which storage service is scalable storage for unstructured data?", "Azure Blob", "Azure Blob", "P1 p4"),
 ("Storage", "Which storage service is NoSQL storage for petabytes of data?", "Azure Tables", "Azure Tables", "P1 p4"),
 ("Storage", "Which storage service provides asynchronous messaging queues?", "Azure Queues", "Azure Queues", "P1 p4"),
 ("Storage", "Which storage service provides block-level virtual disks for VMs?", "Azure Disks", "Azure Disks", "P1 p4"),
 ("Storage", "Which storage service manages cloud file shares using SMB and NFS?", "Azure Files", "Azure Files", "P1 p4"),
 ("Storage", "The Hot blob tier is for what kind of access?", "Frequent, daily use", "Frequent/daily", "P1 p4"),
 ("Storage", "The Cool blob tier is for what kind of access?", "Infrequent, stored at least 30 days", "Infrequent 30d", "P1 p4"),
 ("Storage", "The Cold blob tier is for what kind of access?", "Rarely accessed, at least 90 days", "Rare 90d", "P1 p4"),
 ("Storage", "The Archive blob tier is for what kind of access?", "Long-term storage, at least 180 days", "Long-term 180d", "P1 p4"),
 ("Storage", "What is the maximum usable capacity of an Azure Data Box device?", "80 terabytes", "80 TB", "P1 p5"),
 ("Storage", "Which standard is used to wipe Data Box disks after upload?", "NIST 800-88r1", "NIST 800-88r1", "P1 p5"),
 ("Storage", "What is AzCopy?", "A command-line utility to upload, download, copy and one-way sync files between storage accounts", "", "P1 p5"),
 ("Storage", "What is Azure Storage Explorer?", "A standalone GUI app to manage files and blobs (Windows/macOS/Linux); uses AzCopy on the backend", "", "P1 p5"),
 ("Storage", "What two features does Azure File Sync provide?", "Bidirectional data sync and cloud tiering", "", "P1 p5"),
 ("Identity", "What is Microsoft Entra ID?", "Microsoft's cloud-based identity and access management service", "", "P1 p5"),
 ("Identity", "What bridges on-premises Active Directory with Microsoft Entra ID?", "Microsoft Entra Connect", "Entra Connect", "P1 p5"),
 ("Identity", "What is the synchronization direction for an Entra Domain Services managed domain?", "One-way, from Microsoft Entra ID to Entra Domain Services", "One-way", "P1 p5"),
 ("Identity", "Name the 4 Azure authentication methods.", "Passwords, SSO, MFA, Passwordless", "", "P1 p6"),
 ("Identity", "Name the 3 Azure passwordless options.", "Windows Hello, Microsoft Authenticator, FIDO2 security keys", "", "P1 p6"),
 ("Identity", "MFA factors are based on which three things?", "Something you know, something you are, something you own", "", "P1 p6"),
 ("Security", "What three needs does Microsoft Defender for Cloud fill?", "Continuously assess, Secure, and Defend", "", "P1 p6"),
 ("Security", "Name the 7 layers of the Defense in Depth model.", "Physical; Identity & Access; Perimeter; Network; Compute; Application; Data", "", "P1 p6"),
 ("Governance", "What is Azure Policy?", "A governance tool that enforces and audits resource compliance with organizational standards", "", "P1 p6"),
 ("Governance", "Name the 6 effects in Azure Policy.", "Append, Audit, AuditIfNotExists, DeployIfNotExists, Deny, Disabled", "", "P1 p6"),
 ("Governance", "What is an Azure Initiative?", "A collection of Azure Policy conditions", "Collection of policies", "P1 p7"),
 ("Governance", "What is an Azure Blueprint?", "A set of standards, patterns, and requirements for implementing Azure cloud services", "", "P1 p7"),
 ("Governance", "What does a resource tag consist of?", "A name and a value", "Name and value", "P1 p7"),
 ("Governance", "What does a DoNotDelete lock allow?", "Read and modify, but NOT delete", "Read/modify only", "P1 p7"),
 ("Governance", "What does a ReadOnly lock allow?", "Read only - no modify or delete", "Read only", "P1 p7"),
 ("Management", "What syntax style does the Azure CLI use?", "Bash-style syntax", "Bash", "P1 p7"),
 ("Management", "What does Azure PowerShell use?", "PowerShell cmdlets and REST APIs", "", "P1 p7"),
 ("Management", "What is Azure Resource Manager (ARM)?", "The deployment/management service through which ALL actions on Azure resources go", "", "P1 p7"),
 ("Management", "Which declarative languages are used for ARM templates?", "JSON and Bicep", "JSON and Bicep", "P1 p7"),
 ("Security", "What is Azure Key Vault used for?", "Securely storing and controlling secrets, encryption keys, and certificates", "", "P1 p7"),
 ("Management", "Name the 5 Azure Advisor recommendation categories.", "Reliability, Security, Performance, Operational Excellence, Cost", "", "P1 p7"),
 ("Management", "Name the 3 components of Azure Service Health.", "Azure Status, Service Health, Resource Health", "", "P1 p7"),
 ("Management", "What is Azure Monitor?", "Records telemetry from apps for metrics/logs, supports real-time alerts and auto-scaling", "", "P1 p8"),
 ("Management", "What is Azure Log Analytics?", "The tool to write and run queries against the data Azure Monitor collects", "", "P1 p8"),
 ("Management", "What is Application Insights?", "An Azure Monitor feature that monitors performance and usage of web applications", "", "P1 p8"),
 ("Cost", "Name the factors that affect Azure cost.", "Resource type, consumption, maintenance, geography, subscription type", "", "P1 p6"),
 ("Cost", "What is the Azure Pricing Calculator?", "Calculates the cost of provisioned Azure services for future deployments", "", "P1 p6"),
 ("Cost", "What is the Azure Total Cost of Ownership (TCO) Calculator?", "Compares on-premises expenses to an Azure cloud deployment", "", "P1 p6"),
 ("Cost", "Name the 3 types of cost alerts.", "Budget alerts, Credit alerts, Department spending quota alerts", "", "P1 p6"),
 ("Governance", "What is Microsoft Purview?", "A unified data governance, risk and compliance platform", "", "P1 p6"),
 ("Cost", "What are Reservations best for?", "Stable, predictable workloads (1-year or 3-year commitment)", "", "P1 p8"),
 ("Cost", "What is the Azure savings plan for compute?", "Commit to an hourly spend amount for 1 or 3 years to get discounts on compute usage", "", "P1 p8"),
 ("Cost", "What are Spot Virtual Machines?", "VMs that use unused Azure capacity at reduced prices for interruptible workloads", "", "P1 p8"),
 ("Management", "What is Azure Cloud Shell?", "A browser-based shell that supports both Azure PowerShell and the Azure CLI (Bash)", "", "P1 p9"),
 ("Migration", "What is Azure Migrate?", "A unified migration platform from on-premises to the cloud", "", "P1 p5"),
 ("Hybrid", "What is Azure Arc?", "Extends Azure management and governance to on-premises and multicloud (e.g., Kubernetes, databases)", "", "P1 p1"),
 ("Hybrid", "What is Azure VMware Solution?", "Runs VMware workloads on Azure", "", "P1 p1"),
 ("Compliance", "What is the Service Trust Portal?", "Provides access to Microsoft's Security, Privacy, and Compliance resources", "", "P1 p7"),
]

# ---- P3 extra concepts not already covered ----
ta += [
 ("Security", "What does the Zero Trust model assume?", "That a breach is happening; access is granted by authentication, not location (uses Just-In-Time access)", "", "P3 3.1"),
 ("Identity", "Conditional Access grants or denies access based on which signals?", "Identity, Device, and Location", "", "P3 3.2"),
 ("Security", "What are the two Azure DDoS Protection tiers?", "Basic (always on, free) and Standard (paid)", "", "P3 3.5"),
 ("Support", "Which support plan offers 24x7 phone and email support?", "Standard", "Standard", "P3 3.9"),
 ("Support", "Which support plan includes support for some non-Microsoft products?", "Premier", "Premier", "P3 3.9"),
 ("Subscriptions", "Name the subscription types covered by the reviewer.", "Enterprise Agreement, CSP (Cloud Solution Provider), and the Free account", "", "P3 3.9"),
 ("SLA", "What is RPO (Recovery Point Objective)?", "The maximum acceptable amount of data loss, measured in time", "", "P3 3.9"),
 ("SLA", "What is RTO (Recovery Time Objective)?", "How long it takes to restore service", "", "P3 3.9"),
 ("Lifecycle", "Put the Azure feature lifecycle stages in order.", "Private Preview -> Public Preview -> General Availability (GA)", "", "P3 3.9"),
 ("Cost", "What is the Azure Hybrid Benefit?", "Reusing existing on-premises licenses (e.g., Windows/SQL Server) in Azure", "", "P3 3.8"),
 ("Architecture", "What protects against a single datacenter's update/reboot or hardware fault?", "An Availability Set", "Availability Set", "P3 2.1"),
 ("Security", "Which Defense in Depth layer filters large-scale attacks like DDoS?", "Perimeter", "Perimeter", "P3 3.1"),
]

ws = wb.create_sheet("Use Cases (Type Answer)")
headers = ["#", "Topic", "Question (player types the answer)", "Correct Answer (from PDF)",
           "Suggested Short Answer (<=20 chars)", "Source (PDF)"]
ws.append(headers)
style_header(ws, 1, len(headers))
for i, (topic, q, ans, short, src) in enumerate(ta, start=1):
    ws.append([i, topic, q, ans, short, src])
set_widths(ws, [5, 18, 70, 48, 28, 18])
for r_ in range(2, ws.max_row + 1):
    for c_ in range(1, len(headers) + 1):
        ws.cell(row=r_, column=c_).alignment = WRAP
        ws.cell(row=r_, column=c_).border = BORDER
    ws.cell(row=r_, column=1).alignment = CENTER
ws.freeze_panes = "A2"

# =====================================================================
# SHEET 5: REDUNDANT QUESTIONS
# =====================================================================
redundant = [
 ("ARM templates use JSON", "Type Answer", "P2 Q28; P3 Practice Test Q8; P1 p7", "Keep ONE (e.g., P2 Q28). Same fact asked 3x."),
 ("ExpressRoute = private connection, no public internet", "Type Answer / Matching", "P2 Q7; P2 Q13; P3 Practice Test Q3; P1 p3-4; P3 2.4", "Keep ONE. Appears 5x across all PDFs."),
 ("VPN gateway = encrypted tunnel over the public internet", "Type Answer / Matching", "P2 Q7; P2 Q13; P1 p4; P3 2.4", "P2 Q13 repeats P2 Q7 almost verbatim - drop one."),
 ("Spot pricing = cheapest, interruptible VMs", "Type Answer", "P3 Domain 3 Q1; P3 Practice Test Q9; P1 p8", "Keep ONE. Asked twice in P3 plus stated in P1."),
 ("Azure Advisor = recommendations / find unused VMs", "Type Answer", "P2 Q32; P3 Domain 3 Q5; P1 p7", "Two different angles (unused VMs vs recommendations) - keep at most one."),
 ("Hybrid = most flexible / on-prem + cloud mix", "Type Answer", "P3 Domain 1 Q1; P3 Practice Test Q1", "Same concept; Practice Test Q1 is the scenario version. Keep one."),
 ("Private cloud = total control over hardware/security", "Type Answer", "P3 Domain 1 Q5; P3 Practice Test Q11", "Duplicate concept. Keep one."),
 ("Azure Policy restricts allowed regions", "Type Answer", "P2 Q25; P3 Practice Test Q4; P3 3.6", "Same fact asked 2x (plus stated in P3). Keep one."),
 ("Resource group deletion deletes all resources inside", "Type Answer", "P3 Domain 2 Q4; P1 p2; P3 2.2 / Final reviewer", "Keep the P3 Domain 2 Q4 question; the rest are notes."),
 ("Resource moved to a new RG inherits that RG's permissions", "Type Answer", "P3 Practice Test Q7; P1 p2; Final reviewer", "Keep one question version."),
 ("Availability Zones - minimum of 3 per supporting region", "Type Answer", "P3 Domain 2 Q1; P1 p2; Final reviewer", "Keep P3 Domain 2 Q1; others are notes."),
 ("Archive tier = 180+ days, rarely accessed", "Type Answer", "P3 Domain 2 Q2; P1 p4-5; P3 2.5", "Keep one tier question; the others are the tier table."),
 ("GRS = 16 nines, replicates to secondary region", "Type Answer", "P3 Practice Test Q5; P1 p4; P3 2.5", "Keep one. Same durability fact."),
 ("Azure Data Box = physical transfer of ~40-80 TB", "Type Answer", "P3 Domain 2 Q5; P1 p5; P3 2.6", "Keep one question; others are notes."),
 ("MFA = two DIFFERENT factor types (e.g., password + fingerprint)", "MC / Type Answer", "P2 Q11; P3 Practice Test Q14; P1 p6; P3 3.2", "P2 Q11 (definition) and P3 Q14 (MC) overlap. Keep both ONLY if you want one MC + one type."),
 ("Authentication = proving identity (vs authorization)", "Type Answer / True-False", "P3 Practice Test Q6; P2 Q5 (statement 2); P3 3.2", "P2 Q5 tests the inverse as True/False; P3 Q6 asks directly. Mild overlap."),
 ("Service Health = check first for Azure-wide outage", "Type Answer", "P2 Q33; P3 Practice Test Q12; P1 p7", "P2 Q33 (which tool monitors health) overlaps P3 Q12 (check first). Keep one."),
 ("Defender for Cloud = security recommendations / Free tier", "Type Answer", "P2 Q33; P3 Practice Test Q15; P1 p6; P3 3.5", "Different angles (recommendations vs Free tier) - okay to keep both, but related."),
 ("Tags = name/value metadata on resources", "Type Answer / True-False", "P2 Q30; P2 Q31; P1 p7; P3 3.6", "P2 Q31 (assign tags to every resource) and Q30 statements overlap conceptually."),
 ("Initiative = a group/collection of Azure policies", "Type Answer", "P3 Domain 3 Q3; P1 p7; P3 3.4/3.6", "Keep one question; the rest are notes."),
 ("Shared Responsibility - provider always owns physical layer", "Type Answer", "P2 Q12; P3 Domain 3 Q4; P1 p1; P3 3.1", "P2 Q12 (IaaS provider parts) overlaps P3 Q4. Keep one."),
 ("Lift-and-shift migration => IaaS", "Type Answer", "P2 Q18; P1 p1 (IaaS use cases); P3 1.2", "Keep P2 Q18; P1/P3 state it as a fact."),
 ("Disk pricing = quality (SSD/HDD) + size", "Type Answer", "P3 Practice Test Q10; P3 2.5", "Stated once and asked once - keep the question only."),
 ("Elasticity = automatic scaling", "Type Answer", "P3 Domain 1 Q3; P1 p1; P3 1.3 / Final reviewer", "Keep one question; the rest are notes."),
 ("SaaS example = Microsoft 365 Online", "Type Answer", "P3 Domain 1 Q4; P1 p1; P3 1.2", "Keep the question; P1/P3 state it as an example."),
]
ws = wb.create_sheet("Redundant Questions")
headers = ["Concept (asked more than once)", "Question Type", "Where it appears", "Recommendation"]
ws.append(["These questions/facts repeat across the PDFs. Pick ONE source per concept so your Kahoot does not ask the same thing twice."])
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
ws.cell(row=1, column=1).font = NOTE_FONT
ws.cell(row=1, column=1).alignment = WRAP
ws.append(headers)
style_header(ws, 2, len(headers))
for row in redundant:
    ws.append(list(row))
set_widths(ws, [48, 22, 42, 52])
for r_ in range(3, ws.max_row + 1):
    for c_ in range(1, len(headers) + 1):
        ws.cell(row=r_, column=c_).alignment = WRAP
        ws.cell(row=r_, column=c_).border = BORDER
ws.freeze_panes = "A3"

# =====================================================================
# SHEET 6: SOURCE MAP
# =====================================================================
ws = wb.create_sheet("Source Map")
set_widths(ws, [10, 60, 40])
ws.append(["Code", "PDF File", "What it contributed"])
style_header(ws, 1, 3)
rows = [
 ("P1", "AZ-900 Fundamentals Reviewer.pdf", "Fundamentals notes -> most Type-Answer definitions"),
 ("P2", "AZ900_Practice_Exam_With_Answers.pdf", "Q1-Q36 -> all True/False + 4 Multiple Choice + scenarios"),
 ("P3", "RADOVAN_AZ-900_ Microsoft Azure Fundamentals Reviewer.pdf", "Domain & practice questions -> Type-Answer + 1 MC (Q14)"),
]
for row in rows:
    ws.append(list(row))
for r_ in range(2, ws.max_row + 1):
    for c_ in range(1, 4):
        ws.cell(row=r_, column=c_).alignment = WRAP
        ws.cell(row=r_, column=c_).border = BORDER

# =====================================================================
# SAVE + VALIDATION REPORT
# =====================================================================
wb.save(OUT)

# Validation: counts + character-limit checks
print("Saved:", OUT)
print("Multiple Choice questions :", len(mc))
print("True/False questions      :", len(tf))
print("Type-Answer questions     :", len(ta))
print("Redundancy entries        :", len(redundant))
print("TOTAL quiz items          :", len(mc) + len(tf) + len(ta))
print()
warn = 0
for q, *rest in mc:
    if len(q) > 120:
        print(f"[WARN] MC question >120 chars ({len(q)}): {q}"); warn += 1
for stmt, *_ in tf:
    if len(stmt) > 120:
        print(f"[WARN] TF question >120 chars ({len(stmt)}): {stmt}"); warn += 1
for topic, q, ans, short, src in ta:
    if len(q) > 120:
        print(f"[WARN] TA question >120 chars ({len(q)}): {q}"); warn += 1
print(f"\nCharacter-limit warnings: {warn}")
