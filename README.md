# 🚀 M365 Assessment Consultant Toolkit

A Draft Version toolkit to perform Microsoft 365 security assessments and generate an **interactive enterprise dashboard (Fluent UI style)**.
Next update will contain will be related with SPO.

---

## 📌 Features

✅ Automated M365 Assessment (PowerShell)  
✅ Risk & Findings Report (CSV)  
✅ 30/60/90 Roadmap  
✅ Enterprise Dashboard UI (Intune-like)  
✅ Filtering + Drill-down panel  
✅ Charts + KPI visualization  
✅ Export to Excel (multi-sheet XLSX)  

---

## 📂 Project Structure

---

## ⚙️ Prerequisites

- PowerShell 5.1+
- Microsoft Graph modules
- Required permissions:
  - Directory.Read.All
  - SecurityEvents.Read.All
  - Reports.Read.All

---

## 🚀 Step-by-Step Usage

### ✅ Step 1 — Run Assessment Script

```powershell
.\scripts\M365_Assessment_Report_Generator_Consultant_FINAL.ps1

📂 Output:
M365_Assessment_Consultant_YYYYMMDD_HHMMSS

###✅ Step 2 — Copy Dashboard Script
copy .\scripts\M365_Assessment_Dashboard_Generator_ENTERPRISE_v4.ps1 <OutputFolder>\

###✅ Step 3 — Run Dashboard Generator
cd <OutputFolder>
.\M365_Assessment_Dashboard_Generator_ENTERPRISE_v4.ps1 -OutputRoot "."


###✅ Step 4 — Open Dashboard
00_Dashboard_ENTERPRISE.html
