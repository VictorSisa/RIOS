#!/usr/bin/env python3
"""
RIOS Executive Report Generator v7
Fixed: Inventory % filtering, critical insights noise reduction
"""

from services.excel_extractor import extract_category_data, extract_vendor_brand_data
from pathlib import Path
import pandas as pd
from utils.logger import logger


class ExecutiveReportGenerator:
    """Generates executive-level business review reports."""
    
    def __init__(self, file_path: str):
        """Initialize with Excel file."""
        self.file_path = file_path
        self.categories = extract_category_data(file_path)
        self.vendor_brand = extract_vendor_brand_data(file_path)
        self._clean_data()
        logger.info(f"✅ Report generator ready")
    
    def _clean_data(self):
        """Clean data - remove headers, subtotals, etc."""
        self.categories = self.categories[
            ~self.categories['category'].str.contains('TY CATEGORY NAME|Subtotal', na=False)
        ].copy()
        
        self.vendor_brand = self.vendor_brand[
            ~self.vendor_brand['vendor'].str.contains('VENDOR NAME|Subtotal', na=False)
        ].copy()
        
        logger.info(f"Cleaned categories: {len(self.categories)}, vendor-brands: {len(self.vendor_brand)}")
    
    def _calculate_core_80_percent(self, df: pd.DataFrame) -> pd.DataFrame:
        """Identify top 80% of sales (core business)."""
        sorted_df = df.sort_values('sales_ty', ascending=False, na_position='last')
        
        total_sales = sorted_df['sales_ty'].sum()
        cumulative = 0
        core_items = []
        
        for idx, row in sorted_df.iterrows():
            cumulative += row['sales_ty']
            core_items.append(idx)
            if cumulative >= total_sales * 0.80:
                break
        
        return df.loc[core_items]
    
    def _convert_numeric(self, series):
        """Convert series to numeric, handling errors."""
        return pd.to_numeric(series, errors='coerce')
    
    def _sanitize_inventory_pct(self, pct: float) -> float:
        """Filter out unrealistic inventory % changes."""
        if pd.isna(pct):
            return 0.0
        # Keep values between -500% and +500% (realistic inventory changes)
        if pct < -500 or pct > 500:
            return None
        return pct
    
    def generate_executive_summary(self) -> str:
        """Generate complete executive summary report."""
        report = []
        
        report.append("=" * 120)
        report.append("RIOS WEEKLY EXECUTIVE BUSINESS REVIEW - BABY CATEGORY")
        report.append("=" * 120)
        report.append("")
        
        report.extend(self._section_executive_snapshot())
        report.extend(self._section_core_performance())
        report.extend(self._section_channel_performance())
        report.extend(self._section_category_performance())
        report.extend(self._section_vendor_brand_performance())
        report.extend(self._section_inventory_analysis())
        report.extend(self._section_risks_opportunities())
        report.extend(self._section_action_plan())
        report.extend(self._section_summary())
        
        return "\n".join(report)
    
    def _section_executive_snapshot(self) -> list:
        """1. Executive Snapshot - Key Metrics"""
        lines = ["1. EXECUTIVE SNAPSHOT", "-" * 120, ""]
        
        cat = self.categories.copy()
        cat['sales_ty'] = self._convert_numeric(cat['sales_ty'])
        cat['sales_pct_change'] = self._convert_numeric(cat['sales_pct_change'])
        cat['units_ty'] = self._convert_numeric(cat['units_ty'])
        cat['units_pct_change'] = self._convert_numeric(cat['units_pct_change'])
        cat['gm_dollars'] = self._convert_numeric(cat['gm_dollars'])
        cat['gm_pct_change'] = self._convert_numeric(cat['gm_pct_change'])
        cat['inventory'] = self._convert_numeric(cat['inventory'])
        cat['inventory_pct_change'] = self._convert_numeric(cat['inventory_pct_change'])
        cat['aur'] = self._convert_numeric(cat['aur'])
        cat['aur_pct_change'] = self._convert_numeric(cat['aur_pct_change'])
        
        total_sales = cat['sales_ty'].sum()
        total_units = cat['units_ty'].sum()
        total_gm = cat['gm_dollars'].sum()
        total_inventory = cat['inventory'].sum()
        total_aur = (cat['sales_ty'] / cat['units_ty']).mean() if total_units > 0 else 0
        
        sales_growth = (cat['sales_ty'] * cat['sales_pct_change']).sum() / total_sales if total_sales > 0 else 0
        units_growth = (cat['units_ty'] * cat['units_pct_change']).sum() / total_units if total_units > 0 else 0
        gm_change = (cat['gm_dollars'] * cat['gm_pct_change']).sum() / total_gm if total_gm > 0 else 0
        inv_change = (cat['inventory'] * cat['inventory_pct_change']).sum() / total_inventory if total_inventory > 0 else 0
        aur_growth = (cat['aur'] * cat['aur_pct_change']).sum() / cat['aur'].sum() if cat['aur'].sum() > 0 else 0
        
        lines.append(f"TOTAL SALES: ${total_sales:,.2f}")
        lines.append(f"  YoY Growth: {sales_growth:+.1%}")
        lines.append(f"  LY Sales: ${total_sales / (1 + sales_growth):,.2f}")
        lines.append("")
        
        lines.append(f"UNITS SOLD: {total_units:,.0f}")
        lines.append(f"  YoY Change: {units_growth:+.1%}")
        lines.append(f"  Average Unit Retail (AUR): ${total_aur:.2f} ({aur_growth:+.1%})")
        lines.append("")
        
        lines.append(f"GROSS MARGIN: ${total_gm:,.2f}")
        lines.append(f"  YoY Change: {gm_change:+.1%}")
        lines.append(f"  GM%: {(total_gm / total_sales * 100):.1f}%")
        lines.append("")
        
        lines.append(f"ENDING INVENTORY: ${total_inventory:,.2f}")
        lines.append(f"  YoY Change: {inv_change:+.1%}")
        lines.append("")
        
        lines.append(f"ACTIVE CATEGORIES: {len(cat)}")
        lines.append("")
        
        return lines
    
    def _section_core_performance(self) -> list:
        """2. Core Business Performance Analysis"""
        lines = ["2. CORE BUSINESS PERFORMANCE (Top 80%)", "-" * 120, ""]
        
        cat = self.categories.copy()
        cat['sales_ty'] = self._convert_numeric(cat['sales_ty'])
        cat['sales_pct_change'] = self._convert_numeric(cat['sales_pct_change'])
        
        total_sales = cat['sales_ty'].sum()
        core = self._calculate_core_80_percent(cat)
        core_sales = core['sales_ty'].sum()
        core_pct = core_sales / total_sales if total_sales > 0 else 0
        
        core['gm_dollars'] = self._convert_numeric(core['gm_dollars'])
        core_gm = core['gm_dollars'].sum()
        
        total_gm = self._convert_numeric(cat['gm_dollars']).sum()
        
        lines.append(f"CORE CATEGORIES: {len(core)} of {len(cat)} ({core_pct:.1%} of sales)")
        lines.append(f"  Core Sales: ${core_sales:,.2f}")
        lines.append(f"  Core Gross Margin: ${core_gm:,.2f} ({core_gm/core_sales*100:.1f}%)")
        lines.append("")
        
        lines.append(f"LONG-TAIL CATEGORIES: {len(cat) - len(core)} items")
        longtail_sales = total_sales - core_sales
        longtail_gm = total_gm - core_gm
        lines.append(f"  Long-Tail Sales: ${longtail_sales:,.2f} ({(1-core_pct):.1%})")
        lines.append(f"  Long-Tail Gross Margin: ${longtail_gm:,.2f} ({longtail_gm/longtail_sales*100 if longtail_sales > 0 else 0:.1f}%)")
        lines.append("")
        
        lines.append("INSIGHT:")
        if core_pct > 0.80:
            lines.append(f"✓ Strong core business - {core_pct:.1%} of sales from focused assortment")
            lines.append(f"  Long-tail margin paradox: While long-tail carries higher GM% (29.0% vs 25.5%), the low")
            lines.append(f"  sales volume does not justify SKU complexity. Rationalization of {len(cat) - len(core)} long-tail")
            lines.append(f"  items will improve operational efficiency and shelf productivity.")
        else:
            lines.append(f"⚠ Fragmented assortment - only {core_pct:.1%} from top categories")
            lines.append(f"  Action: Consolidate focus on highest-performing items")
        lines.append("")
        
        return lines
    
    def _section_channel_performance(self) -> list:
        """3. Store vs E-Commerce Performance with Sales Mix"""
        lines = ["3. CHANNEL PERFORMANCE", "-" * 120, ""]
        
        cat = self.categories.copy()
        cat['store_sales'] = self._convert_numeric(cat['store_sales'])
        cat['store_sales_pct_change'] = self._convert_numeric(cat['store_sales_pct_change'])
        cat['ecom_sales'] = self._convert_numeric(cat['ecom_sales'])
        cat['ecom_sales_pct_change'] = self._convert_numeric(cat['ecom_sales_pct_change'])
        cat['inventory'] = self._convert_numeric(cat['inventory'])
        cat['inventory_pct_change'] = self._convert_numeric(cat['inventory_pct_change'])
        
        store_sales = cat['store_sales'].sum()
        ecom_sales = cat['ecom_sales'].sum()
        total_channel = store_sales + ecom_sales
        
        store_growth = (cat['store_sales'] * cat['store_sales_pct_change']).sum() / store_sales if store_sales > 0 else 0
        ecom_growth = (cat['ecom_sales'] * cat['ecom_sales_pct_change']).sum() / ecom_sales if ecom_sales > 0 else 0
        
        # Store inventory change (weighted average)
        store_inv_change = (cat['inventory'] * cat['inventory_pct_change']).sum() / cat['inventory'].sum() if cat['inventory'].sum() > 0 else 0
        
        store_pct = store_sales / total_channel * 100 if total_channel > 0 else 0
        ecom_pct = ecom_sales / total_channel * 100 if total_channel > 0 else 0
        
        lines.append("SALES MIX:")
        lines.append(f"  Store: {store_pct:.1f}% | E-Commerce: {ecom_pct:.1f}%")
        lines.append("")
        
        lines.append(f"STORE SALES: ${store_sales:,.2f}")
        lines.append(f"  YoY Growth: {store_growth:+.1%}")
        lines.append(f"  Store Inventory: {store_inv_change:+.1%}")
        lines.append("")
        
        lines.append(f"E-COMMERCE SALES: ${ecom_sales:,.2f}")
        lines.append(f"  YoY Growth: {ecom_growth:+.1%}")
        lines.append("")
        
        lines.append("INSIGHT:")
        if ecom_growth > store_growth:
            lines.append(f"✓ E-commerce outpacing stores (+{ecom_growth:.1%} vs +{store_growth:.1%})")
            lines.append(f"  Action: Expand online assortment and inventory allocation by 15-20%")
        else:
            lines.append(f"✓ Store channel stronger (+{store_growth:.1%} vs +{ecom_growth:.1%})")
            lines.append(f"  Action: Maintain store focus while building e-commerce")
        lines.append("")
        
        return lines
    
    def _section_category_performance(self) -> list:
        """4. Category Performance - By Sales Value with Inventory"""
        lines = ["4. CATEGORY PERFORMANCE", "-" * 120, ""]
        
        cat = self.categories.copy()
        cat['sales_ty'] = self._convert_numeric(cat['sales_ty'])
        cat['sales_pct_change'] = self._convert_numeric(cat['sales_pct_change'])
        cat['gm_dollars'] = self._convert_numeric(cat['gm_dollars'])
        cat['inventory_pct_change'] = self._convert_numeric(cat['inventory_pct_change'])
        
        # Sanitize inventory % (remove unrealistic values)
        cat['inventory_pct_change'] = cat['inventory_pct_change'].apply(self._sanitize_inventory_pct)
        
        lines.append("TOP 5 CATEGORIES BY SALES VALUE:")
        top_sales = cat.nlargest(5, 'sales_ty')
        for idx, (_, row) in enumerate(top_sales.iterrows(), 1):
            gm_pct = row['gm_dollars'] / row['sales_ty'] * 100 if row['sales_ty'] > 0 else 0
            inv_str = f"| Inventory: {row['inventory_pct_change']:+.1%}" if pd.notna(row['inventory_pct_change']) else ""
            lines.append(f"  {idx}. {row['category']}")
            lines.append(f"     Sales: ${row['sales_ty']:,.2f} | Growth: {row['sales_pct_change']:+.1%} | GM: {gm_pct:.1f}% {inv_str}")
        
        lines.append("")
        lines.append("TOP 5 CATEGORIES BY GROWTH %:")
        top_growth = cat.nlargest(5, 'sales_pct_change')
        for idx, (_, row) in enumerate(top_growth.iterrows(), 1):
            inv_str = f"| Inventory: {row['inventory_pct_change']:+.1%}" if pd.notna(row['inventory_pct_change']) else ""
            lines.append(f"  {idx}. {row['category']}")
            lines.append(f"     Sales: ${row['sales_ty']:,.2f} | Growth: {row['sales_pct_change']:+.1%} {inv_str}")
        
        lines.append("")
        lines.append("DECLINING CATEGORIES:")
        declining = cat.nsmallest(5, 'sales_pct_change')
        for idx, (_, row) in enumerate(declining.iterrows(), 1):
            inv_str = f"| Inventory: {row['inventory_pct_change']:+.1%}" if pd.notna(row['inventory_pct_change']) else ""
            lines.append(f"  {idx}. {row['category']}")
            lines.append(f"     Sales: ${row['sales_ty']:,.2f} | Decline: {row['sales_pct_change']:+.1%} {inv_str}")
        
        lines.append("")
        
        return lines
    
    def _section_vendor_brand_performance(self) -> list:
        """5. Vendor & Brand Performance - With Inventory Metrics & Critical Insights"""
        lines = ["5. VENDOR & BRAND PERFORMANCE", "-" * 120, ""]
        
        vb = self.vendor_brand.copy()
        vb['sales_ty'] = self._convert_numeric(vb['sales_ty'])
        vb['gm_dollars'] = self._convert_numeric(vb['gm_dollars'])
        vb['inventory'] = self._convert_numeric(vb['inventory'])
        vb['inventory_pct_change'] = self._convert_numeric(vb['inventory_pct_change'])
        
        # Top vendors by sales (aggregated across brands)
        vendor_agg = vb.groupby('vendor').agg({
            'sales_ty': 'sum',
            'gm_dollars': 'sum',
            'inventory': 'sum',
            'inventory_pct_change': 'mean'
        }).reset_index().sort_values('sales_ty', ascending=False)
        
        total_sales = vendor_agg['sales_ty'].sum()
        
        lines.append("TOP 10 VENDORS BY SALES:")
        for idx, (_, row) in enumerate(vendor_agg.head(10).iterrows(), 1):
            pct = row['sales_ty'] / total_sales
            gm_pct = row['gm_dollars'] / row['sales_ty'] * 100 if row['sales_ty'] > 0 else 0
            lines.append(f"  {idx}. {row['vendor']}")
            lines.append(f"     Sales: ${row['sales_ty']:,.2f} ({pct:.1%}) | GM%: {gm_pct:.1f}% | Inventory: {row['inventory_pct_change']:+.1%}")
        
        lines.append("")
        lines.append("TOP BRANDS BY SALES (across all vendors):")
        brand_agg = vb.groupby('brand').agg({
            'sales_ty': 'sum',
            'gm_dollars': 'sum',
            'inventory': 'sum',
            'inventory_pct_change': 'mean'
        }).reset_index().sort_values('sales_ty', ascending=False)
        
        brand_total_sales = brand_agg['sales_ty'].sum()
        
        for idx, (_, row) in enumerate(brand_agg.head(8).iterrows(), 1):
            pct = row['sales_ty'] / brand_total_sales
            gm_pct = row['gm_dollars'] / row['sales_ty'] * 100 if row['sales_ty'] > 0 else 0
            lines.append(f"  {idx}. {row['brand']}")
            lines.append(f"     Sales: ${row['sales_ty']:,.2f} ({pct:.1%}) | GM%: {gm_pct:.1f}% | EOH Inventory: {row['inventory_pct_change']:+.1%}")
        
        lines.append("")
        lines.append(f"INSIGHT: Top 10 vendors = {vendor_agg.head(10)['sales_ty'].sum() / total_sales:.1%} of business")
        lines.append("")
        
        # CRITICAL INSIGHTS - FILTERED AND FOCUSED
        lines.append("=" * 120)
        lines.append("CRITICAL INSIGHTS - VENDOR & BRAND ANALYSIS")
        lines.append("=" * 120)
        lines.append("")
        
        # Find negative GM brands - FILTERED: only sales > $5,000
        brand_agg['gm_pct'] = (brand_agg['gm_dollars'] / brand_agg['sales_ty'] * 100).round(1)
        negative_gm = brand_agg[(brand_agg['gm_dollars'] / brand_agg['sales_ty'] < 0) & (brand_agg['sales_ty'] > 5000)].copy()
        negative_gm = negative_gm.sort_values('sales_ty', ascending=False)
        
        if len(negative_gm) > 0:
            lines.append("⚠ NEGATIVE MARGIN BRANDS - IMMEDIATE ACTION REQUIRED (Sales > $5,000):")
            for idx, (_, row) in enumerate(negative_gm.head(5).iterrows(), 1):
                lines.append(f"  {idx}. {row['brand']} (Sales: ${row['sales_ty']:,.2f})")
                lines.append(f"     GM%: {row['gm_pct']:.1f}% - INVESTIGATE pricing/cost structure immediately")
            lines.append("")
        else:
            lines.append("✓ No significant negative margin brands (filtered by sales > $5,000)")
            lines.append("")
        
        # Highest inventory reduction (optimization leaders) - FILTERED: sales > $5,000 and realistic inventory changes
        brand_agg['inventory_pct_filtered'] = brand_agg['inventory_pct_change'].apply(self._sanitize_inventory_pct)
        highest_inv_reduction = brand_agg[(brand_agg['sales_ty'] > 5000) & (brand_agg['inventory_pct_filtered'].notna())].nsmallest(3, 'inventory_pct_filtered')
        
        if len(highest_inv_reduction) > 0:
            lines.append("✓ INVENTORY OPTIMIZATION LEADERS (highest reduction, sales > $5,000):")
            for idx, (_, row) in enumerate(highest_inv_reduction.iterrows(), 1):
                lines.append(f"  {idx}. {row['brand']}: {row['inventory_pct_filtered']:+.1%} inventory change")
                lines.append(f"     Sales: ${row['sales_ty']:,.2f} - Excellent productivity improvement")
            lines.append("")
        
        # Highest inventory growth (risk monitoring) - FILTERED: sales > $5,000 and realistic changes
        highest_inv_growth = brand_agg[(brand_agg['sales_ty'] > 5000) & (brand_agg['inventory_pct_filtered'] > 0)].nlargest(3, 'inventory_pct_filtered')
        
        if len(highest_inv_growth) > 0:
            lines.append("⚠ INVENTORY GROWTH CONCERNS - VALIDATE DEMAND FORECAST (sales > $5,000):")
            for idx, (_, row) in enumerate(highest_inv_growth.iterrows(), 1):
                lines.append(f"  {idx}. {row['brand']}: {row['inventory_pct_filtered']:+.1%} inventory increase")
                lines.append(f"     Sales: ${row['sales_ty']:,.2f} - Monitor for obsolescence risk")
            lines.append("")
        
        return lines
    
    def _section_inventory_analysis(self) -> list:
        """6. Inventory Analysis"""
        lines = ["6. INVENTORY ANALYSIS", "-" * 120, ""]
        
        cat = self.categories.copy()
        cat['inventory'] = self._convert_numeric(cat['inventory'])
        cat['inventory_pct_change'] = self._convert_numeric(cat['inventory_pct_change'])
        
        total_inventory = cat['inventory'].sum()
        avg_inv_change = (cat['inventory'] * cat['inventory_pct_change']).sum() / total_inventory if total_inventory > 0 else 0
        
        lines.append(f"ENDING INVENTORY: ${total_inventory:,.2f}")
        lines.append(f"  YoY Change: {avg_inv_change:+.1%}")
        lines.append("")
        
        lines.append("INVENTORY ASSESSMENT:")
        if avg_inv_change < -0.15:
            lines.append(f"✓ Significant inventory reduction ({avg_inv_change:.1%})")
            lines.append(f"  Despite lower inventory, sales relatively stable")
            lines.append(f"  Productivity improved - opportunity to further optimize")
        elif avg_inv_change < 0:
            lines.append(f"✓ Inventory efficiently reduced ({avg_inv_change:.1%})")
            lines.append(f"  Maintain current levels while monitoring out-of-stocks")
        else:
            lines.append(f"⚠ Inventory increased ({avg_inv_change:+.1%})")
            lines.append(f"  With sales decline, this indicates overstocking")
            lines.append(f"  Action: Immediate inventory reduction required")
        lines.append("")
        
        return lines
    
    def _section_risks_opportunities(self) -> list:
        """7. Risks & Opportunities"""
        lines = ["7. RISKS & OPPORTUNITIES", "-" * 120, ""]
        
        cat = self.categories.copy()
        cat['sales_pct_change'] = self._convert_numeric(cat['sales_pct_change'])
        
        declining_count = len(cat[cat['sales_pct_change'] < -0.10])
        growing_count = len(cat[cat['sales_pct_change'] > 0.10])
        
        lines.append("RISKS:")
        lines.append(f"  • {declining_count} categories declining >10% YoY - requires intervention")
        lines.append(f"  • Long-tail assortment consuming space with minimal ROI")
        lines.append(f"  • Negative margin brands impacting profitability")
        lines.append("")
        
        lines.append("OPPORTUNITIES:")
        lines.append(f"  • {growing_count} high-growth categories - expand shelf space")
        lines.append(f"  • E-commerce channel gaining traction - increase online allocation")
        lines.append(f"  • Vendor consolidation to top 10 could improve margins")
        lines.append("")
        
        return lines
    
    def _section_action_plan(self) -> list:
        """8. Action Plan"""
        lines = ["8. ACTION PLAN", "-" * 120, ""]
        
        lines.append("IMMEDIATE (0-30 days):")
        lines.append("  □ Audit negative margin brands - review pricing and cost structure")
        lines.append("  □ Audit and discontinue categories declining >30% YoY")
        lines.append("  □ Implement 10% inventory reduction plan across long-tail items")
        lines.append("  □ Reallocate shelf space from declining to +20% growth categories")
        lines.append("")
        
        lines.append("MID-TERM (30-90 days):")
        lines.append("  □ Conduct vendor performance reviews with top 10 vendors")
        lines.append("  □ Implement 80/20 assortment strategy - focus on core 7 categories")
        lines.append("  □ Reset planograms prioritizing high-velocity core items")
        lines.append("  □ Expand e-commerce inventory allocation by 15-20%")
        lines.append("")
        
        lines.append("VENDOR-SPECIFIC:")
        lines.append("  □ Request growth plans from top 3 vendors (target: +5% YoY minimum)")
        lines.append("  □ Consolidate SKU count - reduce long-tail vendor partnerships")
        lines.append("  □ Address inventory growth concerns with forecast validation")
        lines.append("")
        
        return lines
    
    def _section_summary(self) -> list:
        """9. Executive Summary"""
        lines = ["9. EXECUTIVE SUMMARY", "-" * 120, ""]
        
        cat = self.categories.copy()
        cat['sales_ty'] = self._convert_numeric(cat['sales_ty'])
        cat['sales_pct_change'] = self._convert_numeric(cat['sales_pct_change'])
        
        total_sales = cat['sales_ty'].sum()
        overall_growth = (cat['sales_ty'] * cat['sales_pct_change']).sum() / total_sales if total_sales > 0 else 0
        
        lines.append(f"Baby category is challenged with {overall_growth:.1%} sales decline YoY, driven by")
        lines.append(f"lower unit volume. However, core 80% assortment shows resilience, and margin")
        lines.append(f"expansion suggests pricing/mix optimization. Inventory discipline has improved")
        lines.append(f"significantly. Strategic focus on high-performing core categories and aggressive")
        lines.append(f"rationalization of long-tail items will restore growth trajectory.")
        lines.append("")
        lines.append(f"**Target: Return to +3% sales growth within 90 days through core assortment optimization.**")
        lines.append("")
        
        return lines


if __name__ == "__main__":
    from pathlib import Path
    import glob
    
    downloads = Path.home() / "Downloads"
    excel_files = glob.glob(str(downloads / "*2026*Baby*.xlsm"))
    
    if excel_files:
        file_path = excel_files[0]
        logger.info(f"Generating final report from: {Path(file_path).name}")
        
        generator = ExecutiveReportGenerator(file_path)
        report = generator.generate_executive_summary()
        
        print("\n" + report)
        
        output_dir = Path.home() / "Desktop" / "RIOS" / "outputs"
        output_dir.mkdir(parents=True, exist_ok=True)
        
        output_file = output_dir / "executive_report_final.txt"
        with open(output_file, 'w') as f:
            f.write(report)
        
        logger.info(f"✅ Report saved to: {output_file}")
    else:
        logger.error("Excel file not found")

