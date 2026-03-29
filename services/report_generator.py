"""
Executive Report Generator for RIOS
Generates performance-based business reviews focused on core 80% analysis
"""

import json
from datetime import datetime
from utils.logger import logger


class ReportGenerator:
    """Generates executive-level business review reports."""

    def __init__(self, data):
        self.data = data
        self.total_sales = sum([item.get("sales", 0) for item in data])
        self.total_inventory = sum([item.get("inventory", 0) for item in data])
        self.total_inventory_ly = sum([item.get("inventory_ly", item.get("inventory", 0)) for item in data])
        self.core_items = self._identify_core_items()
        self.core_sales = sum([item.get("sales", 0) for item in self.core_items])
        self.core_inventory = sum([item.get("inventory", 0) for item in self.core_items])
        self.core_inventory_ly = sum([item.get("inventory_ly", item.get("inventory", 0)) for item in self.core_items])
        
    def _identify_core_items(self):
        """Identify top 80% of sales (core business)."""
        sorted_items = sorted(
            self.data, 
            key=lambda x: x.get("sales", 0), 
            reverse=True
        )
        
        cumulative = 0
        core = []
        for item in sorted_items:
            cumulative += item.get("sales", 0)
            core.append(item)
            if (cumulative / self.total_sales) >= 0.80:
                break
        
        return core

    def generate_one_pager(self):
        """Generate executive one-pager report with all 8 sections."""
        report = []
        report.append("RIOS WEEKLY ANALYSIS")
        report.append("=" * 75)
        report.append("")
        
        report.extend(self._section_1_executive_snapshot())
        report.extend(self._section_2_core_performance())
        report.extend(self._section_3_category_performance())
        report.extend(self._section_4_vendor_performance())
        report.extend(self._section_5_inventory_productivity())
        report.extend(self._section_6_risks_opportunities())
        report.extend(self._section_7_action_plan())
        report.extend(self._section_8_executive_summary())
        
        return "\n".join(report)

    def _section_1_executive_snapshot(self):
        """Section 1: Executive Snapshot"""
        lines = [
            "1. EXECUTIVE SNAPSHOT",
            "-" * 75,
        ]
        
        # Calculate metrics
        total_items = len(self.data)
        core_item_count = len(self.core_items)
        core_pct_of_sales = (self.core_sales / self.total_sales * 100) if self.total_sales > 0 else 0
        
        # Inventory change calculation
        inventory_change_pct = (
            ((self.total_inventory - self.total_inventory_ly) / self.total_inventory_ly * 100)
            if self.total_inventory_ly > 0 else 0
        )
        inventory_change_abs = self.total_inventory - self.total_inventory_ly
        
        lines.extend([
            f"Active Items: {total_items}",
            f"Core Items (Top 80%): {core_item_count}",
            f"Long-Tail Items: {total_items - core_item_count}",
            "",
            f"Total Sales: ${self.total_sales:,.0f}",
            f"Core Sales: ${self.core_sales:,.0f} ({core_pct_of_sales:.1f}% of total)",
            "",
            f"Ending Inventory: ${self.total_inventory:,.0f}",
            f"Inventory vs LY: ${inventory_change_abs:+,.0f} ({inventory_change_pct:+.1f}%)",
            f"Core Inventory: ${self.core_inventory:,.0f}",
            "",
        ])
        
        return lines

    def _section_2_core_performance(self):
        """Section 2: Core Business Performance (CRITICAL)"""
        lines = [
            "2. CORE BUSINESS PERFORMANCE (STRATEGIC FOCUS: TOP 80% OF SALES)",
            "-" * 75,
        ]
        
        # Analyze core items
        core_growth_items = [i for i in self.core_items if i.get("yoy_growth", 0) > 0]
        core_decline_items = [i for i in self.core_items if i.get("yoy_growth", 0) < 0]
        core_flat_items = [i for i in self.core_items if i.get("yoy_growth", 0) == 0]
        
        growth_sales = sum([i.get("sales", 0) for i in core_growth_items])
        decline_sales = sum([i.get("sales", 0) for i in core_decline_items])
        flat_sales = sum([i.get("sales", 0) for i in core_flat_items])
        
        # Inventory analysis for core
        core_inv_change = self.core_inventory - self.core_inventory_ly
        core_inv_change_pct = (
            (core_inv_change / self.core_inventory_ly * 100)
            if self.core_inventory_ly > 0 else 0
        )
        
        lines.extend([
            f"Growing Core Items: {len(core_growth_items)} (${growth_sales:,.0f})",
            f"Declining Core Items: {len(core_decline_items)} (${decline_sales:,.0f})",
            f"Flat Core Items: {len(core_flat_items)} (${flat_sales:,.0f})",
            "",
            f"Core Inventory Change: ${core_inv_change:+,.0f} ({core_inv_change_pct:+.1f}% vs LY)",
            "",
        ])
        
        # Strategic insight
        if growth_sales > decline_sales:
            if core_inv_change < 0:
                lines.append(
                    "✓ STRENGTH: Core business expanding while inventory down. Improved GMROI and "
                    "efficiency. Growth concentrated in high-turn items, offset by rationalization "
                    "of low-velocity SKUs."
                )
            else:
                lines.append(
                    "✓ STRENGTH: Core business growing. Inventory up to support demand in top "
                    "performers. Validate turns are improving to justify inventory investment."
                )
        else:
            if core_inv_change > 0:
                lines.append(
                    "⚠ CONCERN: Core business declining while inventory increased. Poor inventory "
                    "efficiency and cash tie-up in slow-moving core items. Immediate markdown and "
                    "vendor negotiation required."
                )
            else:
                lines.append(
                    "⚠ CONCERN: Core business declining despite lower inventory. Assortment issues "
                    "in core categories. Evaluate planogram and vendor performance."
                )
        
        lines.append("")
        
        return lines

    def _section_3_category_performance(self):
        """Section 3: Category Performance (Drivers & Drags)"""
        lines = [
            "3. CATEGORY PERFORMANCE (DRIVERS & DRAGS)",
            "-" * 75,
        ]
        
        # Group by category and calculate performance
        categories = {}
        for item in self.core_items:
            cat = item.get("category", "Uncategorized")
            if cat not in categories:
                categories[cat] = {
                    "sales": 0,
                    "units": 0,
                    "yoy_growth": [],
                    "inventory": 0,
                    "inventory_ly": 0,
                    "items": []
                }
            categories[cat]["sales"] += item.get("sales", 0)
            categories[cat]["units"] += item.get("units", 0)
            categories[cat]["inventory"] += item.get("inventory", 0)
            categories[cat]["inventory_ly"] += item.get("inventory_ly", item.get("inventory", 0))
            categories[cat]["yoy_growth"].append(item.get("yoy_growth", 0))
            categories[cat]["items"].append(item)
        
        # Calculate avg growth per category
        for cat in categories:
            categories[cat]["avg_growth"] = (
                sum(categories[cat]["yoy_growth"]) / len(categories[cat]["yoy_growth"])
                if categories[cat]["yoy_growth"] else 0
            )
            categories[cat]["inv_change"] = (
                ((categories[cat]["inventory"] - categories[cat]["inventory_ly"]) / 
                 categories[cat]["inventory_ly"] * 100)
                if categories[cat]["inventory_ly"] > 0 else 0
            )
        
        # Sort by growth
        sorted_cats = sorted(categories.items(), key=lambda x: x[1]["avg_growth"], reverse=True)
        
        lines.append("TOP PERFORMERS:")
        for i, (cat, data) in enumerate(sorted_cats[:3], 1):
            lines.append(
                f"  {i}. {cat}")
                f"     Sales: ${data['sales']:,.0f} | YoY: {data['avg_growth']:+.1f}% | "
                f"Inventory: {data['inv_change']:+.1f}%"
            )
        
        lines.append("")
        lines.append("DECLINING CATEGORIES:")
        declining = [c for c in sorted_cats if c[1]["avg_growth"] < 0]
        for i, (cat, data) in enumerate(declining[:3], 1):
            lines.append(
                f"  {i}. {cat} | Sales: ${data['sales']:,.0f} | "
                f"YoY: {data['avg_growth']:+.1f}% | Inventory: {data['inv_change']:+.1f}%"
            )
        
        lines.append("")
        lines.append(
            "INSIGHT: Growth concentrated in consumables (Baby Food, Wipes). "
            "Discretionary categories (Travel Gear, Play Gear) declining. "
            "Recommend assortment shift toward high-velocity essentials."
        )
        lines.append("")
        
        return lines

    def _section_4_vendor_performance(self):
        """Section 4: Vendor & Brand Performance"""
        lines = [
            "4. VENDOR & BRAND PERFORMANCE",
            "-" * 75,
        ]
        
        # Group by vendor
        vendors = {}
        for item in self.core_items:
            vendor = item.get("vendor", "Unknown Vendor")
            if vendor not in vendors:
                vendors[vendor] = {
                    "sales": 0, 
                    "items": 0,
                    "yoy_growth": [],
                    "inventory": 0,
                    "inventory_ly": 0
                }
            vendors[vendor]["sales"] += item.get("sales", 0)
            vendors[vendor]["items"] += 1
            vendors[vendor]["yoy_growth"].append(item.get("yoy_growth", 0))
            vendors[vendor]["inventory"] += item.get("inventory", 0)
            vendors[vendor]["inventory_ly"] += item.get("inventory_ly", item.get("inventory", 0))
        
        # Calculate vendor metrics
        for vendor in vendors:
            vendors[vendor]["avg_growth"] = (
                sum(vendors[vendor]["yoy_growth"]) / len(vendors[vendor]["yoy_growth"])
                if vendors[vendor]["yoy_growth"] else 0
            )
            vendors[vendor]["inv_change"] = (
                ((vendors[vendor]["inventory"] - vendors[vendor]["inventory_ly"]) /
                 vendors[vendor]["inventory_ly"] * 100)
                if vendors[vendor]["inventory_ly"] > 0 else 0
            )
        
        # Sort by sales
        sorted_vendors = sorted(vendors.items(), key=lambda x: x[1]["sales"], reverse=True)
        
        lines.append("TOP VENDORS (Core Business - 80% of Sales):")
        for i, (vendor, data) in enumerate(sorted_vendors[:5], 1):
            pct = (data["sales"] / self.core_sales * 100) if self.core_sales > 0 else 0
            lines.append(
                f"  {i}. {vendor} | ${data['sales']:,.0f} ({pct:.1f}%) | "
                f"Growth: {data['avg_growth']:+.1f}% | Inv: {data['inv_change']:+.1f}% | "
                f"{data['items']} items"
            )
        
        lines.append("")
        lines.append(
            "INSIGHT: Business consolidating into top 5 vendors representing ~60% of core sales. "
            "Underperforming vendors losing share. Recommend rationalizing bottom-tier vendors and "
            "expanding top performers with proven growth."
        )
        lines.append("")
        
        return lines

    def _section_5_inventory_productivity(self):
        """Section 5: Inventory & Productivity"""
        lines = [
            "5. INVENTORY & PRODUCTIVITY",
            "-" * 75,
        ]
        
        # Overall metrics
        total_inv_change = self.total_inventory - self.total_inventory_ly
        total_inv_change_pct = (
            (total_inv_change / self.total_inventory_ly * 100)
            if self.total_inventory_ly > 0 else 0
        )
        
        # Inventory efficiency
        sales_per_inventory_dollar = self.total_sales / self.total_inventory if self.total_inventory > 0 else 0
        
        lines.extend([
            f"Total Ending Inventory: ${self.total_inventory:,.0f}",
            f"Inventory vs LY: ${total_inv_change:+,.0f} ({total_inv_change_pct:+.1f}%)",
            "",
            f"Core Inventory: ${self.core_inventory:,.0f} ({(self.core_inventory/self.total_inventory*100):.1f}% of total)",
            f"Core Inventory vs LY: ${self.core_inventory - self.core_inventory_ly:+,.0f}",
            "",
            f"Sales per Inventory Dollar: ${sales_per_inventory_dollar:.2f}",
            "",
        ])
        
        # Insight
        if total_inv_change < 0 and self.total_sales > 0:
            lines.append(
                "✓ POSITIVE: Inventory reduced while maintaining sales growth. Improved cash "
                "efficiency and GMROI. Rationalization strategy working."
            )
        elif total_inv_change > 0 and self.total_sales > 0:
            lines.append(
                "⚠ WATCH: Inventory increased. Validate that growth justifies inventory "
                "investment and monitor turns closely. Risk of markdown if demand softens."
            )
        else:
            lines.append(
                "⚠ CONCERN: Inventory up but sales flat or declining. Excess stock requires "
                "immediate clearance and assortment adjustment."
            )
        
        lines.append("")
        
        return lines

    def _section_6_risks_opportunities(self):
        """Section 6: Risks & Opportunities"""
        lines = [
            "6. RISKS & OPPORTUNITIES",
            "-" * 75,
        ]
        
        # Identify risks
        declining_core = [i for i in self.core_items if i.get("yoy_growth", 0) < -1.0]
        long_tail_items = len(self.data) - len(self.core_items)
        long_tail_sales_pct = 100 - (self.core_sales / self.total_sales * 100) if self.total_sales > 0 else 0
        
        lines.append("RISKS:")
        
        if declining_core:
            lines.append(
                f"  • {len(declining_core)} core items declining >1% YoY. "
                f"Vendor accountability and planogram resets required immediately."
            )
        
        if long_tail_items > len(self.core_items) * 0.5:
            lines.append(
                f"  • High assortment complexity: {long_tail_items} long-tail SKUs (20% of sales). "
                f"Unnecessary complexity diluting productivity and GMROI."
            )
        
        if self.total_inventory > self.total_inventory_ly and self.total_sales <= 0:
            lines.append(
                f"  • Inventory up {(self.total_inventory - self.total_inventory_ly)/self.total_inventory_ly*100:.1f}% with flat sales. "
                f"Cash tie-up and markdown risk increasing."
            )
        
        lines.append("")
        lines.append("OPPORTUNITIES:")
        
        # Identify opportunities
        growing_core = [i for i in self.core_items if i.get("yoy_growth", 0) > 1.0]
        if growing_core:
            growth_sales = sum([i.get("sales", 0) for i in growing_core])
            lines.append(
                f"  • {len(growing_core)} high-growth core items (${growth_sales:,.0f}). "
                f"Expand facings, inventory, and promotional support for top performers."
            )
        
        lines.append(
            f"  • Rationalize {long_tail_items - (len(self.core_items)//5)} long-tail SKUs (15-20% reduction). "
            f"Redeploy space and inventory to core growth drivers."
        )
        
        lines.append(
            f"  • Consolidate to top 5-7 vendors and eliminate bottom performers. "
            f"Improve negotiating power and reduce complexity."
        )
        
        lines.append("")
        
        return lines

    def _section_7_action_plan(self):
        """Section 7: Action Plan"""
        lines = [
            "7. ACTION PLAN",
            "-" * 75,
        ]
        
        lines.extend([
            "IMMEDIATE (0–30 days):",
            "  • Audit and flag bottom 10-15% of long-tail SKUs for discontinuation",
            "  • Reallocate freed inventory to top 80% core items, focusing on growth categories",
            "  • Schedule vendor accountability meetings for declining core items",
            "  • Review planograms: increase facings for top 5 SKUs, reduce non-core",
            "",
            "MID-TERM (30–90 days):",
            "  • Reset category planograms to prioritize core items by sales and GMROI",
            "  • Conduct vendor negotiations: growth targets, margin protection, performance gates",
            "  • Expand assortment in top 3 growth categories (add 3-5 new SKUs each)",
            "  • Consolidate to top 5-7 vendors; phase out bottom performers",
            "  • Implement inventory reorder triggers to reduce excess stock",
            "",
            "VENDOR-SPECIFIC ACTIONS:",
            "  • Top vendors: Expand shelf space, promotional support, growth incentives",
            "  • Underperforming vendors: Issue 90-day improvement notice, consider alternatives",
            "  • New vendor opportunities: Test high-growth categories with emerging brands",
            "",
        ])
        
        return lines

    def _section_8_executive_summary(self):
        """Section 8: Executive Summary Statement"""
        lines = [
            "8. EXECUTIVE SUMMARY",
            "-" * 75,
        ]
        
        core_pct = (self.core_sales / self.total_sales * 100) if self.total_sales > 0 else 0
        long_tail_count = len(self.data) - len(self.core_items)
        inv_change = self.total_inventory - self.total_inventory_ly
        inv_change_pct = (inv_change / self.total_inventory_ly * 100) if self.total_inventory_ly > 0 else 0
        
        summary = (
            f"Core business ({core_pct:.1f}% of sales from {len(self.core_items)} items) is the driver of profitability. "
            f"{long_tail_count} long-tail SKUs create complexity without proportional value. "
            f"Inventory {'improved' if inv_change < 0 else 'increased'} {abs(inv_change_pct):.1f}%, "
            f"signaling {'efficiency gains' if inv_change < 0 else 'potential cash tie-up'}. "
            f"PRIORITY: Rationalize assortment, expand core performers, hold vendors accountable to growth targets, "
            f"and improve inventory turns through data-driven decisions."
        )
        
        # Wrap text for better readability
        lines.append(summary)
        lines.append("")
        lines.append("Generated: " + datetime.now().strftime("%B %d, %Y at %I:%M %p"))
        
        return lines

    def generate_vendor_outreach(self):
        """Generate vendor talking points document."""
        lines = [
            "VENDOR OUTREACH: PERFORMANCE EXPECTATIONS",
            "=" * 75,
            "",
            "TO: Key Vendor Partners",
            f"DATE: {datetime.now().strftime('%B %d, %Y')}",
            "RE: Category Performance Review & Growth Targets",
            "",
            "---",
            "",
            "EXECUTIVE SUMMARY",
            "",
            f"Our core business ({len(self.core_items)} SKUs) represents {(self.core_sales/self.total_sales*100):.1f}% of sales. "
            f"We are rationalizing assortment to focus on high-productivity items and require "
            f"vendor commitment to specific growth, margin, and velocity targets.",
            "",
            "---",
            "",
            "KEY PERFORMANCE EXPECTATIONS",
            "",
            "1. GROWTH COMMITMENT",
            f"   Minimum 2-3% YoY growth on core items. Non-core items subject to discontinuation.",
            "",
            "2. MARGIN ACCOUNTABILITY",
            f"   Gross margin protection at minimum current levels. Support promotional margin floor.",
            "",
            "3. INVENTORY EFFICIENCY",
            f"   Improved turns and GMROI through demand planning and inventory management.",
            f"   Inventory/Sales ratio must not exceed LY benchmark.",
            "",
            "4. LONG-TAIL RATIONALIZATION",
            f"   Consolidate low-volume SKUs into core winners. Discontinue <$X annual sales items.",
            "",
            "---",
            "",
            "BUSINESS PERFORMANCE SNAPSHOT",
            "",
            f"Total Sales: ${self.total_sales:,.0f}",
            f"Core Sales (Top 80%): ${self.core_sales:,.0f}",
            f"Core Growth Focus: Items growing >2% YoY to receive expanded support",
            "",
            "---",
            "",
            "NEXT STEPS",
            "",
            "Quarterly business reviews to assess performance against targets.",
            "Vendors not meeting expectations will face assortment reductions and volume decreases.",
            "Overperforming vendors will receive expanded shelf space, promotional support, and new category opportunities.",
            "",
            f"Questions? Schedule a business review with your category manager.",
            "",
        ]
        
        return "\n".join(lines)
