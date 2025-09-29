#!/usr/bin/env python3
"""
GridEdge Compute Center - Scenario Modeling
Generates scenario-based financial projections and ROI charts
"""

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from openpyxl import load_workbook
import os

def read_roi_data(excel_file, sheet_name="ROI Timeline"):
    """Read ROI timeline data from the Excel file"""
    wb = load_workbook(excel_file)
    ws = wb[sheet_name]

    data = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] and row[4]:
            data.append({
                'month': row[0],
                'monthly_revenue': row[1] if row[1] else 0,
                'monthly_opex': row[2] if row[2] else 0,
                'net_cashflow': row[3] if row[3] else 0,
                'cumulative_cf': row[4] if row[4] else 0,
                'net_position': row[5] if row[5] else 0,
                'roi_percent': row[6] if row[6] else 0,
                'payback': row[7] if row[7] else "NO"
            })

    return pd.DataFrame(data)

def create_roi_chart(df, output_path, chart_title):
    """Create ROI timeline chart showing break-even forecast"""
    fig, ax = plt.subplots(figsize=(14, 8))
    months = range(1, len(df) + 1)

    ax.plot(months, df['cumulative_cf'] / 1_000_000, linewidth=3, color='#1f4e79', label='Cumulative Cash Flow')
    ax.axhline(y=0, color='red', linestyle='--', linewidth=2, alpha=0.7, label='Break-even Line')

    net_positions = df['net_position'].values
    break_even_month = next((i + 1 for i, net_pos in enumerate(net_positions) if net_pos >= 0), None)

    if break_even_month:
        break_even_cf = df.iloc[break_even_month - 1]['cumulative_cf'] / 1_000_000
        ax.plot(break_even_month, break_even_cf, 'go', markersize=12, label=f'Break-even: Month {break_even_month}')
        ax.annotate(f'Break-even\nMonth {break_even_month}', 
                    xy=(break_even_month, break_even_cf),
                    xytext=(break_even_month + 1, break_even_cf + 0.5),
                    arrowprops=dict(arrowstyle='->', color='green', lw=2),
                    fontsize=11, fontweight='bold', color='green',
                    bbox=dict(boxstyle="round", facecolor='lightgreen', alpha=0.7))
    else:
        min_idx = df['cumulative_cf'].idxmin()
        min_month = min_idx + 1
        min_value = df.iloc[min_idx]['cumulative_cf'] / 1_000_000
        ax.plot(min_month, min_value, 'ro', markersize=12, label=f'Maximum Loss: Month {min_month}')
        ax.annotate(f'Maximum Loss\nMonth {min_month}\n${min_value:.1f}M', 
                    xy=(min_month, min_value),
                    xytext=(min_month + 6, min_value - 0.5),
                    arrowprops=dict(arrowstyle='->', color='red', lw=2),
                    fontsize=11, fontweight='bold', color='red',
                    bbox=dict(boxstyle="round", facecolor='lightcoral', alpha=0.7))

    ax.set_xlabel('Month', fontsize=12, fontweight='bold')
    ax.set_ylabel('Cumulative Cash Flow (Millions USD)', fontsize=12, fontweight='bold')
    ax.set_title(chart_title, fontsize=16, fontweight='bold', pad=20)
    ax.grid(True, alpha=0.3, linestyle='--')

    ax.set_xlim(0, len(df) + 2)
    year_ticks = list(range(0, len(df) + 1, 12))
    ax.set_xticks(year_ticks)
    ax.set_xticklabels([f'Year {i}' if i > 0 else 'Start' for i in range(len(year_ticks))])
    ax.set_xticks(range(1, len(df) + 1), minor=True)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x:.1f}M'))
    ax.legend(loc='upper left', fontsize=10, framealpha=0.9)

    final_cf = df.iloc[-1]['cumulative_cf'] / 1_000_000
    final_roi = df.iloc[-1]['roi_percent'] * 100
    avg_net = df['net_cashflow'].mean()

    summary = f"""5-Year Summary:
• Final Cash Flow: ${final_cf:.1f}M
• ROI: {final_roi:.1f}%
• Avg Monthly Net: ${avg_net/1000:.0f}K
• CapEx: $6.2M"""
    ax.text(0.02, 0.98, summary, transform=ax.transAxes,
            verticalalignment='top', fontsize=10,
            bbox=dict(boxstyle="round", facecolor='lightblue', alpha=0.8))

    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches='tight', facecolor='white')
    print(f"✅ Scenario chart saved to: {output_path}")
    return fig, break_even_month

def generate_scenario_models(excel_file="financial_model.xlsx", graphs_dir="graphs"):
    """Generate and save charts for different scenarios."""
    os.makedirs(graphs_dir, exist_ok=True)
    try:
        base_df_original = read_roi_data(excel_file, "ROI Timeline")
    except FileNotFoundError:
        print(f"❌ Error: {excel_file} not found.")
        return
    if base_df_original.empty:
        print("❌ Error: ROI Timeline sheet is empty.")
        return

    capex = 6200000  # USD
    scenarios = {
        "bear": {
            "gpu_revenue": 80000,
            "btc_revenue": 15000,
            "opex": 153000,
            "title": "GridEdge 5MW – Bear Case (Low Utilization, High Power Costs)",
            "filename": "scenario_bear.png"
        },
        "base": {
            "gpu_revenue": 120000,
            "btc_revenue": 18000,
            "opex": 138000,
            "title": "GridEdge 5MW – Base Case",
            "filename": "scenario_base.png"
        },
        "bull": {
            "gpu_revenue": 250000,
            "btc_revenue": 25000,
            "opex": 116000,
            "title": "GridEdge 5MW – Bull Case (High Utilization, Low Power Costs)",
            "filename": "scenario_bull.png"
        }
    }

    for name, s in scenarios.items():
        df = base_df_original.copy()
        df['monthly_revenue'] = s['gpu_revenue'] + s['btc_revenue']
        df['monthly_opex'] = s['opex']
        df['net_cashflow'] = df['monthly_revenue'] - df['monthly_opex']
        df['cumulative_cf'] = df['net_cashflow'].cumsum()
        df['net_position'] = df['cumulative_cf'] - capex
        df['roi_percent'] = (df['cumulative_cf'] / capex)
        df['payback'] = df['net_position'].apply(lambda x: "YES" if x >= 0 else "NO")

        output_path = os.path.join(graphs_dir, s['filename'])
        create_roi_chart(df, output_path, chart_title=s['title'])

    print("\n✅ Scenario Charts Generated:")
    for s in scenarios.values():
        print(f"   • {s['title']}: graphs/{s['filename']}")

def main():
    generate_scenario_models()

if __name__ == "__main__":
    main()
