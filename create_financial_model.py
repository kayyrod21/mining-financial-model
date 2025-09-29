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
    
    # CapEx data
    capex_data = [
        ["Category", "Subcategory", "Capacity/Units", "Unit Cost", "Total Cost", "Notes"],
        ["Equipment", "ASIC Miners", "2 MW", 1500, 3000000, "Bitcoin mining hardware"],
        ["Equipment", "GPU Clusters", "1 MW", 2500, 2500000, "AI/ML workloads"],
        ["Equipment", "Network Equipment", "5 MW", 200, 1000000, "Switches, routers, security"],
        ["Equipment", "Servers & Storage", "5 MW", 300, 1500000, "Management and storage"],
        ["", "", "", "Equipment Subtotal:", 8000000, ""],
        ["Facility", "Modular Construction", "5000 sq ft", 400, 2000000, "Pre-fab datacenter modules"],
        ["Facility", "Site Preparation", "1 lot", 500000, 500000, "Foundation, access roads"],
        ["Facility", "Security Systems", "1 facility", 250000, 250000, "Cameras, access control"],
        ["", "", "", "Facility Subtotal:", 2750000, ""],
        ["Power & Cooling", "Geothermal Connection", "3 MW", 5000, 15000000, "LaGeo PPA infrastructure"],
        ["Power & Cooling", "Solar Array", "2 MW", 1500, 3000000, "Rooftop and ground mount"],
        ["Power & Cooling", "Battery Storage", "1000 kWh", 1000, 1000000, "Grid stabilization"],
        ["Power & Cooling", "Gas Generators", "2 MW", 1000, 2000000, "Backup power"],
        ["Power & Cooling", "HVAC Systems", "5 MW", 800, 4000000, "Cooling with heat recovery"],
        ["Power & Cooling", "Electrical Distribution", "5 MW", 600, 3000000, "Transformers, switchgear"],
        ["", "", "", "Power & Cooling Subtotal:", 28000000, ""],
        ["", "", "", "Base Project Cost:", 38750000, ""],
        ["Contingency", "15% Buffer", "", "", 5812500, "Risk mitigation"],
        ["", "", "", "Total CapEx:", 44562500, ""],
        ["", "", "", "Target CapEx:", 32000000, "Scaled to budget"]
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
    
    # Revenue assumptions
    gpu_base = 80000
    asic_base = 18000  # 2MW at $9K/MW
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
    
    # Operating expense data
    opex_data = []
    for i, month in enumerate(months):
        # Energy calculation: $0.07/kWh × 5MW × 720 hours × 85% uptime
        energy_cost = 0.07 * 5000 * 720 * 0.85
        staff_cost = 15000
        maintenance_cost = 5000
        insurance_cost = 2000
        connectivity_cost = 3000
        other_cost = 1500
        
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
    
    # Calculate ROI timeline
    initial_investment = 32000000
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
    
    # Key metrics
    total_capex = 32000000
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
        ["GridEdge Compute Center - Financial Summary", ""],
        ["", ""],
        ["Investment Overview", ""],
        ["Total CapEx", total_capex],
        ["Project Capacity", "5 MW"],
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
    print("\nModel includes:")
    print("- CapEx Breakdown: Detailed equipment and infrastructure costs")
    print("- Monthly Revenue Forecast: 5-year revenue projections")
    print("- Operating Expenses: Energy, staff, and maintenance costs")
    print("- ROI Timeline: Month-by-month cash flow analysis")
    print("- Executive Summary: Key metrics and assumptions")
    
    return filename

if __name__ == "__main__":
    main()
