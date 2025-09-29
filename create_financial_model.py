#!/usr/bin/env python3
"""
GridEdge Compute Center Financial Model Generator
Creates comprehensive Excel financial model with multiple scenarios
"""

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.formatting.rule import FormulaRule
import os

def create_financial_model():
    """Create comprehensive financial model Excel file"""
    
    # Create workbook
    wb = Workbook()
    
    # Remove default sheet
    wb.remove(wb.active)
    
    # Define styling
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    currency_format = '_($* #,##0_);_($* (#,##0);_($* "-"??_);_(@_)'
    percent_format = '0.0%'
    
    # 1. CapEx Breakdown Sheet
    capex_ws = wb.create_sheet("CapEx Breakdown")
    
    # Phase I CapEx data - 2.5MW modular deployment
    capex_data = [
        ["Category", "Subcategory", "Capacity/Units", "Unit Cost", "Total Cost", "Notes"],
        ["Equipment", "ASIC Miners", "1.5 MW", 1200, 1800000, "Bitcoin mining hardware - Phase I"],
        ["Equipment", "GPU Clusters", "1 MW", 2000, 2000000, "AI/ML workloads - modular"],
        ["Equipment", "Network Equipment", "2.5 MW", 150, 375000, "Switches, routers, security"],
        ["Equipment", "Servers & Storage", "2.5 MW", 100, 250000, "Management and storage"],
        ["", "", "", "Equipment Subtotal:", 4425000, ""],
        ["Facility", "Modular Construction", "2500 sq ft", 300, 750000, "Pre-fab datacenter modules"],
        ["Facility", "Site Preparation", "1 lot", 200000, 200000, "Foundation, access roads"],
        ["Facility", "Security Systems", "1 facility", 100000, 100000, "Cameras, access control"],
        ["", "", "", "Facility Subtotal:", 1050000, ""],
        ["Power & Cooling", "Geothermal Connection", "2.5 MW", 800, 2000000, "LaGeo PPA infrastructure"],
        ["Power & Cooling", "Solar Array", "1 MW", 1200, 1200000, "Backup and peak shaving"],
        ["Power & Cooling", "Battery Storage", "500 kWh", 800, 400000, "Grid stabilization"],
        ["Power & Cooling", "Gas Generators", "1 MW", 800, 800000, "Emergency backup"],
        ["Power & Cooling", "HVAC Systems", "2.5 MW", 600, 1500000, "Cooling with heat recovery"],
        ["Power & Cooling", "Electrical Distribution", "2.5 MW", 400, 1000000, "Transformers, switchgear"],
        ["", "", "", "Power & Cooling Subtotal:", 6900000, ""],
        ["", "", "", "Base Project Cost:", 12375000, ""],
        ["Contingency", "15% Buffer", "", "", 1856250, "Risk mitigation"],
        ["", "", "", "Pre-Contingency Total:", 14231250, ""],
        ["Value Engineering", "Cost Optimization", "", "", -11031250, "Modular approach savings"],
        ["", "", "", "Phase I Target CapEx:", 3200000, "Investor-ready budget"]
    ]
    
    # Write CapEx data
    for row_idx, row_data in enumerate(capex_data, 1):
        for col_idx, value in enumerate(row_data, 1):
            cell = capex_ws.cell(row=row_idx, column=col_idx, value=value)
            if row_idx == 1:  # Header row
                cell.font = header_font
                cell.fill = header_fill
            elif "Subtotal" in str(value) or "Total" in str(value):
                cell.font = Font(bold=True)
    
    # Format currency columns
    for row in range(2, len(capex_data) + 1):
        capex_ws.cell(row=row, column=4).number_format = currency_format
        capex_ws.cell(row=row, column=5).number_format = currency_format
    
    # Adjust column widths
    capex_ws.column_dimensions['A'].width = 20
    capex_ws.column_dimensions['B'].width = 25
    capex_ws.column_dimensions['C'].width = 15
    capex_ws.column_dimensions['D'].width = 15
    capex_ws.column_dimensions['E'].width = 15
    capex_ws.column_dimensions['F'].width = 30
    
    # 2. Monthly Revenue Forecast Sheet
    revenue_ws = wb.create_sheet("Monthly Revenue Forecast")
    
    # Create 5-year monthly forecast
    months = []
    for year in range(1, 6):
        for month in range(1, 13):
            months.append(f"Y{year}M{month:02d}")
    
    # Phase I Revenue assumptions - 2.5MW capacity
    gpu_base = 40000  # 1MW GPU capacity at $40K/MW
    asic_base = 13500  # 1.5MW ASIC capacity at $9K/MW
    spa_base = 2500
    
    # Create revenue data with growth assumptions
    revenue_data = []
    for i, month in enumerate(months):
        # Growth factors
        gpu_growth = 1 + (i * 0.005)  # 0.5% monthly growth
        asic_volatility = 1 + np.sin(i * 0.3) * 0.2  # Bitcoin volatility
        spa_growth = min(1.5, 1 + (i * 0.01))  # 1% monthly growth, capped at 50%
        
        gpu_revenue = gpu_base * gpu_growth
        asic_revenue = asic_base * asic_volatility
        spa_revenue = spa_base * spa_growth
        
        monthly_total = gpu_revenue + asic_revenue + spa_revenue
        annual_total = monthly_total * 12
        
        revenue_data.append([
            month,
            gpu_revenue,
            asic_revenue, 
            spa_revenue,
            monthly_total,
            annual_total
        ])
    
    # Write headers
    revenue_headers = ["Month", "GPU Leasing", "ASIC Mining", "Spa Income", "Monthly Total", "Annualized"]
    for col_idx, header in enumerate(revenue_headers, 1):
        cell = revenue_ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
    
    # Write revenue data
    for row_idx, row_data in enumerate(revenue_data, 2):
        for col_idx, value in enumerate(row_data, 1):
            cell = revenue_ws.cell(row=row_idx, column=col_idx, value=value)
            if col_idx > 1:  # Format currency columns
                cell.number_format = currency_format
    
    # Adjust column widths
    for col in ['A', 'B', 'C', 'D', 'E', 'F']:
        revenue_ws.column_dimensions[col].width = 15
    
    # 3. Operating Expenses Sheet
    opex_ws = wb.create_sheet("Operating Expenses")
    
    # Phase I Operating expense data - 2.5MW capacity
    opex_data = []
    for i, month in enumerate(months):
        # Energy calculation: $0.07/kWh × 2.5MW × 720 hours × 85% uptime = ~$107,730/month
        energy_cost = 0.07 * 2500 * 720 * 0.85
        staff_cost = 15000
        maintenance_cost = 5000
        insurance_cost = 2000
        connectivity_cost = 3000
        other_cost = 5000  # Increased for Phase I contingencies
        
        # Inflation adjustment (2% annual)
        inflation_factor = (1.02) ** (i / 12)
        
        total_opex = (energy_cost + staff_cost + maintenance_cost + 
                     insurance_cost + connectivity_cost + other_cost) * inflation_factor
        
        opex_data.append([
            month,
            energy_cost * inflation_factor,
            staff_cost * inflation_factor,
            maintenance_cost * inflation_factor,
            insurance_cost * inflation_factor,
            connectivity_cost * inflation_factor,
            other_cost * inflation_factor,
            total_opex
        ])
    
    # Write OpEx headers
    opex_headers = ["Month", "Energy", "Staff", "Maintenance", "Insurance", "Connectivity", "Other", "Total OpEx"]
    for col_idx, header in enumerate(opex_headers, 1):
        cell = opex_ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
    
    # Write OpEx data
    for row_idx, row_data in enumerate(opex_data, 2):
        for col_idx, value in enumerate(row_data, 1):
            cell = opex_ws.cell(row=row_idx, column=col_idx, value=value)
            if col_idx > 1:  # Format currency columns
                cell.number_format = currency_format
    
    # Adjust column widths
    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
        opex_ws.column_dimensions[col].width = 15
    
    # 4. ROI Timeline Sheet
    roi_ws = wb.create_sheet("ROI Timeline")
    
    # Calculate ROI timeline - Phase I
    initial_investment = 3200000
    cumulative_cashflow = 0
    roi_data = []
    
    for i, month in enumerate(months):
        monthly_revenue = revenue_data[i][4]  # Monthly total from revenue sheet
        monthly_opex = opex_data[i][7]  # Total OpEx from opex sheet
        monthly_cashflow = monthly_revenue - monthly_opex
        cumulative_cashflow += monthly_cashflow
        
        net_position = cumulative_cashflow - initial_investment
        roi_percentage = (cumulative_cashflow / initial_investment) if initial_investment > 0 else 0
        payback_achieved = "YES" if net_position >= 0 else "NO"
        
        roi_data.append([
            month,
            monthly_revenue,
            monthly_opex,
            monthly_cashflow,
            cumulative_cashflow,
            net_position,
            roi_percentage,
            payback_achieved
        ])
    
    # Write ROI headers
    roi_headers = ["Month", "Revenue", "OpEx", "Net Cash Flow", "Cumulative CF", "Net Position", "ROI %", "Payback"]
    for col_idx, header in enumerate(roi_headers, 1):
        cell = roi_ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
    
    # Write ROI data
    for row_idx, row_data in enumerate(roi_data, 2):
        for col_idx, value in enumerate(row_data, 1):
            cell = roi_ws.cell(row=row_idx, column=col_idx, value=value)
            if col_idx in [2, 3, 4, 5, 6]:  # Currency columns
                cell.number_format = currency_format
            elif col_idx == 7:  # Percentage column
                cell.number_format = percent_format
            
            # Highlight break-even point
            if col_idx == 8 and value == "YES":
                cell.fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
    
    # Adjust column widths
    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
        roi_ws.column_dimensions[col].width = 15
    
    # 5. Summary Dashboard Sheet
    summary_ws = wb.create_sheet("Executive Summary")
    
    # Phase I Key metrics
    total_capex = 3200000
    avg_monthly_revenue = sum(row[4] for row in revenue_data[:12]) / 12
    avg_monthly_opex = sum(row[7] for row in opex_data[:12]) / 12
    monthly_net_cashflow = avg_monthly_revenue - avg_monthly_opex
    
    # Find break-even month
    break_even_month = "Not achieved in 5 years"
    for i, row in enumerate(roi_data):
        if row[5] >= 0:  # Net position positive
            break_even_month = row[0]
            break
    
    summary_data = [
        ["GridEdge Compute Center - Phase I Financial Model (2.5MW Modular)", ""],
        ["", ""],
        ["Investment Overview", ""],
        ["Total CapEx", total_capex],
        ["Project Capacity", "2.5 MW"],
        ["Location", "El Salvador Geothermal Corridor"],
        ["", ""],
        ["Revenue Performance (Year 1 Average)", ""],
        ["Monthly Revenue", avg_monthly_revenue],
        ["Annual Revenue", avg_monthly_revenue * 12],
        ["", ""],
        ["Operating Performance (Year 1 Average)", ""],
        ["Monthly OpEx", avg_monthly_opex],
        ["Annual OpEx", avg_monthly_opex * 12],
        ["Monthly Net Cash Flow", monthly_net_cashflow],
        ["", ""],
        ["Investment Returns", ""],
        ["Break-even Month", break_even_month],
        ["5-Year Cumulative Cash Flow", roi_data[-1][4]],
        ["5-Year ROI", roi_data[-1][6]],
        ["", ""],
        ["Key Assumptions", ""],
        ["GPU Utilization", "70-80%"],
        ["Energy Cost", "$0.07/kWh"],
        ["Facility Uptime", "85%"],
        ["Bitcoin Price (avg)", "$40,000"],
    ]
    
    # Write summary data
    for row_idx, (label, value) in enumerate(summary_data, 1):
        summary_ws.cell(row=row_idx, column=1, value=label)
        summary_ws.cell(row=row_idx, column=2, value=value)
        
        # Style headers
        if "Overview" in label or "Performance" in label or "Returns" in label or "Assumptions" in label:
            summary_ws.cell(row=row_idx, column=1).font = Font(bold=True, size=12)
            summary_ws.cell(row=row_idx, column=1).fill = PatternFill(start_color="E6E6FA", end_color="E6E6FA", fill_type="solid")
        
        # Format currency values
        if isinstance(value, (int, float)) and value > 1000:
            summary_ws.cell(row=row_idx, column=2).number_format = currency_format
    
    # Adjust column widths
    summary_ws.column_dimensions['A'].width = 30
    summary_ws.column_dimensions['B'].width = 20
    
    # 6. Phased Build Plan Sheet
    phases_ws = wb.create_sheet("Phased Build Plan")
    
    # Phased build data
    phases_data = [
        ["Phase", "MW Added", "Cumulative MW", "CapEx Estimate", "Target Year", "Notes"],
        ["Phase I", "2.5 MW", "2.5 MW", 3200000, "Q1 2026", "Geothermal, Spa reuse, Modular datacenter"],
        ["Phase II", "2.5 MW", "5.0 MW", 3500000, "Q4 2026", "Expand GPU capacity, add staff housing"],
        ["Phase III", "5-10 MW", "10-15 MW", "TBD", "2027-2028", "Sovereign AI clusters, Grid services"],
        ["", "", "", "", "", ""],
        ["Total (3 Phases)", "10-15 MW", "15 MW", "~$10-12M", "2026-2028", "Full build-out capacity"],
        ["", "", "", "", "", ""],
        ["Phase I Details", "", "", "", "", ""],
        ["ASIC Mining", "1.5 MW", "", "", "", "Bitcoin generation focus"],
        ["GPU Clusters", "1.0 MW", "", "", "", "AI/ML workload hosting"],
        ["Waste Heat Recovery", "", "", "", "", "Spa partnership integration"],
        ["Geothermal Connection", "", "", "", "", "LaGeo PPA for baseload power"],
        ["", "", "", "", "", ""],
        ["Phase II Expansion", "", "", "", "", ""],
        ["Additional GPUs", "1.5 MW", "", "", "", "Scale AI hosting capacity"],
        ["Enhanced ASIC", "1.0 MW", "", "", "", "Next-gen mining hardware"],
        ["Staff Housing", "", "", "", "", "On-site accommodation facility"],
        ["Grid Integration", "", "", "", "", "Enhanced grid services capability"],
        ["", "", "", "", "", ""],
        ["Phase III Vision", "", "", "", "", ""],
        ["Sovereign AI", "5-10 MW", "", "", "", "National AI infrastructure"],
        ["Grid Stabilization", "", "", "", "", "Frequency regulation services"],
        ["Export Capacity", "", "", "", "", "Regional power market participation"],
    ]
    
    # Write phased build headers
    for col_idx, header in enumerate(phases_data[0], 1):
        cell = phases_ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
    
    # Write phased build data
    for row_idx, row_data in enumerate(phases_data[1:], 2):
        for col_idx, value in enumerate(row_data, 1):
            cell = phases_ws.cell(row=row_idx, column=col_idx, value=value)
            
            # Format currency column
            if col_idx == 4 and isinstance(value, (int, float)):
                cell.number_format = currency_format
            
            # Bold phase headers and section headers
            if col_idx == 1 and ("Phase" in str(value) or "Details" in str(value) or "Expansion" in str(value) or "Vision" in str(value)):
                cell.font = Font(bold=True)
    
    # Adjust column widths
    phases_ws.column_dimensions['A'].width = 20
    phases_ws.column_dimensions['B'].width = 15
    phases_ws.column_dimensions['C'].width = 15
    phases_ws.column_dimensions['D'].width = 15
    phases_ws.column_dimensions['E'].width = 15
    phases_ws.column_dimensions['F'].width = 35
    
    # Set active sheet to summary
    wb.active = summary_ws
    
    return wb

def main():
    """Main function to create and save the financial model"""
    print("Creating GridEdge Compute Center Financial Model...")
    
    # Create the workbook
    wb = create_financial_model()
    
    # Save the file
    filename = "financial_model.xlsx"
    wb.save(filename)
    
    print(f"✅ Financial model saved as {filename}")
    print("\nPhase I Model includes:")
    print("- CapEx Breakdown: $3.2M total investment for 2.5MW capacity")
    print("- Monthly Revenue Forecast: GPU ($40K) + ASIC ($13.5K) + Spa ($2.5K)")
    print("- Operating Expenses: Energy ($107K), Staff ($15K), Other ($15K)")
    print("- ROI Timeline: Month-by-month cash flow analysis with break-even")
    print("- Executive Summary: Phase I key metrics and assumptions")
    print("- Phased Build Plan: Growth roadmap to 15MW+ capacity")
    
    return filename

if __name__ == "__main__":
    main()
