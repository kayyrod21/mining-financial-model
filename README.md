# ⚡ GridEdge Compute Center – Hybrid Bitcoin + AI Financial Model

Welcome to the open-source financial modeling and pitch material 
repository for **GridEdge Compute Center** — a sovereign-grade Bitcoin 
mining and GPU hosting facility powered by geothermal and solar energy in 
El Salvador.

> Developed by **Bitcoin Mining Consulting LLC**

---

## 🧠 Project Summary

GridEdge is a **Phase I: 5MW modular hybrid compute facility** designed to serve two critical 
sectors, with planned expansion to 15MW+ in future phases:

-   **Bitcoin Mining (ASIC infrastructure):** Daily BTC generation using 
    clean, low-cost energy from El Salvador’s volcano grid.
-   **AI Compute Leasing (GPU infrastructure):** Infrastructure-as-a-Service 
    (IaaS) for researchers, startups, and decentralized AI workloads.

The facility reuses waste heat through a wellness center (hydrotherapy + 
sauna) reducing OpEx while improving community well-being.

---

## 🌎 Strategic Context

> “$900B in annual data center CapEx by 2028 will rival the entire S&P500.” – Morgan Stanley  

> “AI infrastructure spend will reach $4 trillion by 2030.” – Jensen Huang, NVIDIA

The sovereign AI and Bitcoin infrastructure era is beginning. El 
Salvador's renewable energy, political will, and digital ecosystem offer a 
rare opportunity to build a geopolitically aligned compute center outside 
of traditional VC and surveillance models.

---

## 💰 Financial Overview

| Component             | Estimated Value      |
|----------------------|----------------------|
| Total CapEx (Phase I)| $6.2M USD            |
| Revenue (GPU Lease)  | $85,000/month        |
| Revenue (BTC Mining) | $17,000 - $18,000/month |
| Break-even Timeline  | ~3.5 years (base case) |

Power source mix: **70% Geothermal** + **Solar Redundancy** + **Grid 
Backup**

---

## 📊 What’s Inside This Repo
```bash
mining-financial-model/
├── create_financial_model.py         ← Script to generate financial_model.xlsx
├── financial_model.xlsx              ← Output file with CapEx, revenue, ROI, and summary
├── assumptions.md                     ← Key inputs & economic scenarios
├── graphs/                           ← Auto-generated chart images (CapEx, ROI, etc.)
│   ├── capex_breakdown.png           ← CapEx pie chart  
│   └── capex_detailed.png            ← CapEx detailed bar chart
├── scripts/                          ← Python scripts to generate visual charts
├── presentation/                     ← Pitch deck (PDF, Canva or Figma exports)
└── README.md                         ← This file
```

### 🐍 Python Scripts

**`create_financial_model.py`** generates the Excel-based financial model for the GridEdge Compute Center. It creates worksheets for CapEx breakdown, revenue forecasts, operating expenses, ROI timeline, executive summary, phased build plan, and financing options. The model targets ~$6M CapEx with optimized revenue streams and ~$132K monthly OpEx.

**`scripts/generate_capex_chart.py`** reads the Excel model and creates visual charts showing CapEx distribution by category (Equipment, Facility, Power & Cooling, Legal/Admin, and Contingency).

**`scripts/generate_roi_chart.py`** creates ROI timeline visualization showing break-even forecast and monthly cash flow analysis over 5 years.

## 📈 Visuals

The following charts are auto-generated from the financial model:

-   **[CapEx Breakdown](graphs/capex_breakdown.png)** - Investment distribution pie chart
-   **[CapEx Detailed](graphs/capex_detailed.png)** - Detailed cost breakdown by category
-   **[ROI Timeline](graphs/roi_chart.png)** - Break-even forecast and cumulative cash flow
-   **[Cash Flow Analysis](graphs/cashflow_detailed.png)** - Monthly revenue vs expenses breakdown

*Note: Charts reflect Phase I model assumptions. Current projections show challenging cash flow - model may need revenue optimization or OpEx reduction.*

---

## ⚙️ Next Steps

-   Clone this repo and update the Excel model for your own region or scale
-   Run the chart scripts to generate visuals for investor decks
-   Fork or contribute if you’re building similar infrastructure in a 
    sovereign context

---

## 💰 Financing Options

| Asset | Financing Type | Example Lenders/Platforms | Notes |
|---|---|---|---|
| Bitcoin (self-mined) | BTC-backed Line of Credit | Unchained, Ledn, Local El Salvador Lenders | Use freshly mined BTC as collateral for stablecoin loans (Liquid USDt). |
| GPUs (NVIDIA A6000/H100) | Hardware-backed Leasing | Coreweave, Lambda, Hydra Host | Finance GPU acquisition via lease-to-own or sale-leaseback agreements. |
| ASIC Miners (S21/T21) | Equipment Financing | Asset-backed loans from crypto-friendly funds | Use the hardware itself as collateral for financing the purchase. |
| Modular Containers/Infra | Equipment-based Financing | Private credit, off-balance sheet leasing | Collateralize the physical infrastructure for private loans. |
| Forward Sales | Prepaid GPU Contracts | Direct sales to AI startups, render farms | Sell future GPU access at a discount to fund current OpEx and CapEx. |
| Carbon Credits | Green Financing | ESG funds, Methane capture grants | Monetize methane flaring partnerships for carbon credits and grant funding. |
| Sovereign Partnership | Government Grants/Loans | El Salvador National Bitcoin Office | Explore partnerships for building sovereign AI infrastructure, potentially unlocking grants. |

---

## 📫 Contact

**Kaylee Rodriguez**  
Founder, Bitcoin Mining Consulting LLC  
📧 kayleebmc@gmail.com  
📍 El Salvador / United States

---

## 📜 License

**Creative Commons BY-NC-SA 4.0**  
>   You may remix, adapt, and build upon this work **non-commercially**, as 
    long as you credit and license your new creations under the same terms.
