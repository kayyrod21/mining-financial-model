#!/usr/bin/env python3
"""
GridEdge Compute Center - ROI Timeline Chart Generator
Generates ROI and cash flow visualizations from financial_model.xlsx
"""

import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
import os

def read_roi_data(excel_file, sheet_name="ROI Timeline"):
    """Extract ROI timeline data from Excel sheet"""
    wb = load_workbook(excel_file, data_only=True)
    ws = wb[sheet_name]

    data = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] and row[4]:
            data.append({
                'month': row[0],
                'monthly_revenue': row[1] or 0,
                'monthly_opex': row[2] or 0,
                'net_cashflow': row[3] or 0,
                'cumulative_cf': row[4] or 0,
                'net_position': row[5] or 0,
                'roi_percent': row[6] or 0,
                'payback': row[7] or "NO"
            })

    return pd.DataFrame(data)

def create_roi_chart(df, output_path):
    """Generate break-even chart from cumulative cash flow"""
    fig, ax = plt.subplots(figsize=(12, 6))
    months = range(1, len(df) + 1)

    # Plot cumulative CF
    ax.plot(months, df['cumulative_cf'] / 1e6, label='Cumulative Cash Flow', linewidth=3, color='green')
    ax.axhline(0, color='red', linestyle='--', label='Break-even Line')

    # Determine break-even or min loss point
    break_even_month = next((i+1 for i, x in enumerate(df['net_position']) if x >= 0), None)

    if break_even_month:
        val = df.iloc[break_even_month - 1]['cumulative_cf'] / 1e6
        ax.plot(break_even_month, val, 'go', markersize=10, label=f'Break-even: Month {break_even_month}')
        ax.annotate(f'Break-even\nMonth {break_even_month}', xy=(break_even_month, val),
                    xytext=(break_even_month + 3, val + 0.5),
                    arrowprops=dict(arrowstyle='->', color='green'),
                    fontsize=10, bbox=dict(boxstyle="round", fc="lightgreen", alpha=0.7))
    else:
        min_idx = df['cumulative_cf'].idxmin()
        min_val = df.iloc[min_idx]['cumulative_cf'] / 1e6
        ax.plot(min_idx+1, min_val, 'ro', markersize=10, label='Max Loss')
        ax.annotate(f'Max Loss\nMonth {min_idx+1}\n${min_val:.1f}M', xy=(min_idx+1, min_val),
                    xytext=(min_idx+5, min_val - 0.5),
                    arrowprops=dict(arrowstyle='->', color='red'),
                    fontsize=10, bbox=dict(boxstyle="round", fc="lightcoral", alpha=0.7))

    ax.set_title("GridEdge Phase I (5MW) – Break-even Forecast", fontsize=14, fontweight='bold')
    ax.set_xlabel("Month")
    ax.set_ylabel("Cumulative Cash Flow (USD Millions)")
    ax.legend()
    ax.grid(True, linestyle='--', alpha=0.6)

    # Add summary box
    final_cf = df.iloc[-1]['cumulative_cf'] / 1e6
    final_roi = df.iloc[-1]['roi_percent'] * 100
    avg_net = df['net_cashflow'].mean() / 1e3
    ax.text(0.01, 0.98,
            f"5-Year Summary:\n"
            f"• Final CF: ${final_cf:.1f}M\n"
            f"• ROI: {final_roi:.1f}%\n"
            f"• Avg Net CF: ${avg_net:.0f}K/mo\n"
            f"• CapEx: $6.2M",
            transform=ax.transAxes,
            verticalalignment='top',
            bbox=dict(boxstyle="round", facecolor="lightblue", alpha=0.8),
            fontsize=9)

    plt.tight_layout()
    plt.savefig(output_path, dpi=300)
    print(f"✅ ROI chart saved to: {output_path}")
    return fig, break_even_month

def main():
    excel_file = "financial_model.xlsx"
    output_path = "graphs/roi_chart.png"

    if not os.path.exists(excel_file):
        print(f"❌ Error: Missing {excel_file}.")
        return

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    print("📈 Generating ROI Timeline...")

    df = read_roi_data(excel_file)
    fig, break_even = create_roi_chart(df, output_path)

    print("\n📊 Summary:")
    if break_even:
        print(f"• Break-even reached: Month {break_even}")
    else:
        print("• Break-even not achieved within 5 years")

    print(f"• Final CF: ${df.iloc[-1]['cumulative_cf']/1e6:.1f}M")
    print(f"• ROI: {df.iloc[-1]['roi_percent']*100:.1f}%")
    print(f"• Avg Monthly Net CF: ${df['net_cashflow'].mean()/1e3:.0f}K")

if __name__ == "__main__":
    main()
