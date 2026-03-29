"""
Report Generator for RIOS
Generates executive one-pager and vendor outreach reports from structured sales data.
"""

from __future__ import annotations

from typing import Any

from utils.logger import logger

# ---------------------------------------------------------------------------
# Constants — extracted to avoid magic numbers throughout the module
# ---------------------------------------------------------------------------
CORE_SALES_THRESHOLD: float = 0.80  # 80/20 rule: top 80% of sales = core business
TOP_N_CATEGORIES: int = 3           # Number of top/bottom categories to display
TOP_N_VENDORS: int = 5              # Number of top vendors to display in one-pager
OUTREACH_VENDORS: int = 5           # Number of vendors in outreach document

SEPARATOR: str = "=" * 100
SUBSEPARATOR: str = "-" * 100


class ReportGenerator:
    """
    Generates an executive one-pager and vendor outreach document from structured
    sales data following the 80/20 Pareto principle.

    Args:
        data: List of item dicts with keys: category, vendor, brand, sku,
              sales, units, yoy_growth, gross_margin_pct, inventory.

    Raises:
        TypeError: If *data* is not a list.
        ValueError: If items in *data* are missing required keys.

    Example:
        >>> items = [
        ...     {"category": "Diapers", "vendor": "Acme", "brand": "SoftBaby",
        ...      "sku": "SKU001", "sales": 50000.0, "units": 1200.0,
        ...      "yoy_growth": 5.2, "gross_margin_pct": 28.5, "inventory": 10000.0}
        ... ]
        >>> rg = ReportGenerator(items)
        >>> one_pager = rg.generate_one_pager()
        >>> outreach = rg.generate_vendor_outreach()
    """

    REQUIRED_KEYS: frozenset = frozenset(
        {"category", "vendor", "brand", "sku", "sales", "units",
         "yoy_growth", "gross_margin_pct", "inventory"}
    )

    def __init__(self, data: list) -> None:
        self._validate_data(data)
        self.data: list = data
        logger.info("ReportGenerator initialized with %d items", len(data))

    # -----------------------------------------------------------------------
    # Validation & utilities
    # -----------------------------------------------------------------------

    def _validate_data(self, data: list) -> None:
        """
        Validate the incoming data structure before processing.

        Args:
            data: The data list to validate.

        Raises:
            TypeError: If *data* is not a list.
            ValueError: If items are missing required keys.
        """
        if not isinstance(data, list):
            raise TypeError(
                f"data must be a list, got {type(data).__name__}"
            )
        if len(data) == 0:
            logger.warning(
                "ReportGenerator received empty data — reports will be minimal"
            )
            return
        missing = self.REQUIRED_KEYS - set(data[0].keys())
        if missing:
            raise ValueError(
                f"Data items are missing required keys: {sorted(missing)}"
            )

    def _safe_float(self, value: Any, default: float = 0.0) -> float:
        """
        Safely convert *value* to float, returning *default* on failure or NaN.

        Args:
            value:   The value to convert.
            default: Fallback value when conversion fails.

        Returns:
            Float representation of *value*, or *default*.
        """
        try:
            result = float(value)
            # Treat NaN as the default
            return result if result == result else default
        except (TypeError, ValueError):
            return default

    # -----------------------------------------------------------------------
    # Aggregate helpers
    # -----------------------------------------------------------------------

    def _total_sales(self) -> float:
        """Return total sales across all items."""
        return sum(self._safe_float(item.get("sales")) for item in self.data)

    def _total_units(self) -> float:
        """Return total units across all items."""
        return sum(self._safe_float(item.get("units")) for item in self.data)

    def _total_inventory(self) -> float:
        """Return total inventory value across all items."""
        return sum(self._safe_float(item.get("inventory")) for item in self.data)

    def _core_items(self) -> list:
        """
        Return items representing the top 80% of sales (Pareto core).

        Returns:
            Sorted list of item dicts whose cumulative sales reach the
            CORE_SALES_THRESHOLD.
        """
        sorted_items = sorted(
            self.data,
            key=lambda x: self._safe_float(x.get("sales")),
            reverse=True,
        )
        total = self._total_sales()
        if total == 0:
            return sorted_items

        cumulative = 0.0
        core: list = []
        for item in sorted_items:
            cumulative += self._safe_float(item.get("sales"))
            core.append(item)
            if cumulative >= total * CORE_SALES_THRESHOLD:
                break
        return core

    def _categories_summary(self) -> dict:
        """
        Aggregate sales data by category.

        Returns:
            Dict mapping category name to aggregated metrics
            (sales, units, inventory, avg_growth, avg_gm_pct).
        """
        cats: dict = {}
        for item in self.data:
            cat = str(item.get("category") or "Unknown")
            sales = self._safe_float(item.get("sales"))
            units = self._safe_float(item.get("units"))
            gm = self._safe_float(item.get("gross_margin_pct"))
            growth = self._safe_float(item.get("yoy_growth"))
            inv = self._safe_float(item.get("inventory"))

            if cat not in cats:
                cats[cat] = {
                    "sales": 0.0, "units": 0.0, "inventory": 0.0,
                    "gm_sum": 0.0, "gm_count": 0,
                    "growth_sum": 0.0, "growth_count": 0,
                }
            cats[cat]["sales"] += sales
            cats[cat]["units"] += units
            cats[cat]["inventory"] += inv
            if gm != 0.0:
                cats[cat]["gm_sum"] += gm
                cats[cat]["gm_count"] += 1
            if growth != 0.0:
                cats[cat]["growth_sum"] += growth
                cats[cat]["growth_count"] += 1

        result: dict = {}
        for cat, agg in cats.items():
            result[cat] = {
                "sales": agg["sales"],
                "units": agg["units"],
                "inventory": agg["inventory"],
                "avg_gm_pct": (
                    agg["gm_sum"] / agg["gm_count"]
                    if agg["gm_count"] > 0 else 0.0
                ),
                "avg_growth": (
                    agg["growth_sum"] / agg["growth_count"]
                    if agg["growth_count"] > 0 else 0.0
                ),
            }
        return result

    def _vendors_summary(self) -> dict:
        """
        Aggregate sales data by vendor.

        Returns:
            Dict mapping vendor name to aggregated metrics
            (sales, units, inventory, avg_gm_pct).
        """
        vendors: dict = {}
        for item in self.data:
            vendor = str(item.get("vendor") or "Unknown")
            sales = self._safe_float(item.get("sales"))
            units = self._safe_float(item.get("units"))
            gm = self._safe_float(item.get("gross_margin_pct"))
            inv = self._safe_float(item.get("inventory"))

            if vendor not in vendors:
                vendors[vendor] = {
                    "sales": 0.0, "units": 0.0, "inventory": 0.0,
                    "gm_sum": 0.0, "gm_count": 0,
                }
            vendors[vendor]["sales"] += sales
            vendors[vendor]["units"] += units
            vendors[vendor]["inventory"] += inv
            if gm != 0.0:
                vendors[vendor]["gm_sum"] += gm
                vendors[vendor]["gm_count"] += 1

        result: dict = {}
        for vendor, agg in vendors.items():
            result[vendor] = {
                "sales": agg["sales"],
                "units": agg["units"],
                "inventory": agg["inventory"],
                "avg_gm_pct": (
                    agg["gm_sum"] / agg["gm_count"]
                    if agg["gm_count"] > 0 else 0.0
                ),
            }
        return result

    # -----------------------------------------------------------------------
    # Public report generators
    # -----------------------------------------------------------------------

    def generate_one_pager(self) -> str:
        """
        Generate the 8-section executive one-pager report.

        Returns:
            Multi-line string containing the full report.

        Raises:
            Exception: Re-raises any unexpected error after logging it.

        Example:
            >>> report = rg.generate_one_pager()
            >>> assert "EXECUTIVE SNAPSHOT" in report
        """
        try:
            lines: list = []
            lines.append(SEPARATOR)
            lines.append("RIOS WEEKLY EXECUTIVE BUSINESS REVIEW")
            lines.append(SEPARATOR)
            lines.append("")

            lines.extend(self._section_executive_snapshot())
            lines.extend(self._section_core_performance())
            lines.extend(self._section_category_performance())
            lines.extend(self._section_vendor_performance())
            lines.extend(self._section_inventory_productivity())
            lines.extend(self._section_risks_opportunities())
            lines.extend(self._section_action_plan())
            lines.extend(self._section_executive_summary())

            report = "\n".join(lines)
            logger.info("One-pager generated successfully (%d lines)", len(lines))
            return report
        except Exception as exc:
            logger.error("Error generating one-pager: %s", exc, exc_info=True)
            raise

    def generate_vendor_outreach(self) -> str:
        """
        Generate vendor outreach talking-points document.

        Returns:
            Multi-line string containing per-vendor talking points.

        Raises:
            Exception: Re-raises any unexpected error after logging it.

        Example:
            >>> doc = rg.generate_vendor_outreach()
            >>> assert "VENDOR OUTREACH" in doc
        """
        try:
            vendors = self._vendors_summary()
            total_sales = self._total_sales()
            sorted_vendors = sorted(
                vendors.items(),
                key=lambda x: x[1]["sales"],
                reverse=True,
            )

            lines: list = []
            lines.append(SEPARATOR)
            lines.append("RIOS VENDOR OUTREACH — WEEKLY TALKING POINTS")
            lines.append(SEPARATOR)
            lines.append("")

            if not sorted_vendors:
                lines.append("No vendor data available.")
                return "\n".join(lines)

            for i, (vendor, data) in enumerate(sorted_vendors[:OUTREACH_VENDORS], 1):
                pct = (
                    data["sales"] / total_sales * 100
                    if total_sales > 0 else 0.0
                )
                lines.append(f"{i}. {vendor}")
                lines.append(
                    f"   Sales: ${data['sales']:,.0f} ({pct:.1f}% of total)"
                )
                lines.append(f"   Units: {data['units']:,.0f}")
                lines.append(f"   Avg GM%: {data['avg_gm_pct']:.1f}%")
                lines.append(f"   Inventory: ${data['inventory']:,.0f}")
                lines.append("")
                lines.append("   TALKING POINTS:")
                if data["avg_gm_pct"] < 20:
                    lines.append(
                        "   • Discuss cost reduction opportunities to improve margin"
                    )
                if total_sales > 0 and data["sales"] / total_sales > 0.10:
                    lines.append(
                        "   • Major partner — schedule QBR and growth planning session"
                    )
                lines.append(
                    "   • Review assortment performance and in-stock position"
                )
                lines.append(
                    "   • Explore promotional co-op funding opportunities"
                )
                lines.append("")

            report = "\n".join(lines)
            logger.info(
                "Vendor outreach generated successfully (%d vendors)",
                len(sorted_vendors),
            )
            return report
        except Exception as exc:
            logger.error(
                "Error generating vendor outreach: %s", exc, exc_info=True
            )
            raise

    # -----------------------------------------------------------------------
    # Section helpers — each returns list[str] to be joined by generate_*
    # -----------------------------------------------------------------------

    def _section_executive_snapshot(self) -> list:
        """
        Section 1: Executive Snapshot — high-level KPIs.

        Returns:
            List of formatted lines for this section.
        """
        lines = ["1. EXECUTIVE SNAPSHOT", SUBSEPARATOR, ""]

        total_sales = self._total_sales()
        total_units = self._total_units()
        total_inv = self._total_inventory()

        # Weighted-average YoY growth
        weighted_growth = (
            sum(
                self._safe_float(item.get("sales"))
                * self._safe_float(item.get("yoy_growth"))
                for item in self.data
            ) / total_sales
            if total_sales > 0 else 0.0
        )

        # Weighted-average gross margin %
        weighted_gm = (
            sum(
                self._safe_float(item.get("sales"))
                * self._safe_float(item.get("gross_margin_pct"))
                for item in self.data
            ) / total_sales
            if total_sales > 0 else 0.0
        )

        aur = total_sales / total_units if total_units > 0 else 0.0
        active_cats = len(
            {str(item.get("category", "")) for item in self.data}
        )

        lines.append(f"TOTAL SALES: ${total_sales:,.2f}")
        lines.append(f"  YoY Growth: {weighted_growth:+.1f}%")
        lines.append("")
        lines.append(f"TOTAL UNITS: {total_units:,.0f}")
        lines.append(f"  Average Unit Retail (AUR): ${aur:.2f}")
        lines.append("")
        lines.append(f"GROSS MARGIN: {weighted_gm:.1f}%")
        lines.append("")
        lines.append(f"ENDING INVENTORY: ${total_inv:,.2f}")
        lines.append("")
        lines.append(f"ACTIVE CATEGORIES: {active_cats}")
        lines.append("")
        return lines

    def _section_core_performance(self) -> list:
        """
        Section 2: Core Business Performance (Top 80% of Sales).

        Returns:
            List of formatted lines for this section.
        """
        lines = ["2. CORE BUSINESS PERFORMANCE (Top 80%)", SUBSEPARATOR, ""]

        total_sales = self._total_sales()
        core = self._core_items()
        core_sales = sum(self._safe_float(item.get("sales")) for item in core)
        core_pct = core_sales / total_sales * 100 if total_sales > 0 else 0.0
        longtail_count = len(self.data) - len(core)

        lines.append(f"Core Items (Top 80% of Sales): {len(core)} items")
        lines.append(f"Core Sales: ${core_sales:,.2f} ({core_pct:.1f}% of total)")
        lines.append(f"Long-Tail Items: {longtail_count} items")
        lines.append("")

        if core_pct >= CORE_SALES_THRESHOLD * 100:
            lines.append(
                f"✓ Strong core business — {core_pct:.1f}% of sales from "
                "focused assortment"
            )
        else:
            lines.append(
                f"⚠ Fragmented assortment — only {core_pct:.1f}% from top items"
            )
            lines.append(
                "  Action: Consolidate focus on highest-performing SKUs"
            )
        lines.append("")
        return lines

    def _section_category_performance(self) -> list:
        """
        Section 3: Category Performance — top and declining categories.

        Returns:
            List of formatted lines for this section.
        """
        lines = ["3. CATEGORY PERFORMANCE", SUBSEPARATOR, ""]

        cats = self._categories_summary()
        if not cats:
            lines.append("No category data available.")
            lines.append("")
            return lines

        sorted_cats = sorted(
            cats.items(), key=lambda x: x[1]["avg_growth"], reverse=True
        )

        lines.append("TOP PERFORMERS:")
        for i, (cat, data) in enumerate(sorted_cats[:TOP_N_CATEGORIES], 1):
            lines.append(
                f"  {i}. {cat} | "
                f"Sales: ${data['sales']:,.0f} | YoY: {data['avg_growth']:+.1f}%"
            )
        lines.append("")

        lines.append("DECLINING CATEGORIES:")
        declining = sorted_cats[-TOP_N_CATEGORIES:]
        for i, (cat, data) in enumerate(reversed(declining), 1):
            lines.append(
                f"  {i}. {cat} | "
                f"Sales: ${data['sales']:,.0f} | YoY: {data['avg_growth']:+.1f}%"
            )
        lines.append("")
        return lines

    def _section_vendor_performance(self) -> list:
        """
        Section 4: Vendor Performance — top vendors by sales.

        Returns:
            List of formatted lines for this section.
        """
        lines = ["4. VENDOR PERFORMANCE", SUBSEPARATOR, ""]

        vendors = self._vendors_summary()
        if not vendors:
            lines.append("No vendor data available.")
            lines.append("")
            return lines

        total_sales = self._total_sales()
        sorted_vendors = sorted(
            vendors.items(), key=lambda x: x[1]["sales"], reverse=True
        )

        lines.append("TOP VENDORS BY SALES:")
        for i, (vendor, data) in enumerate(sorted_vendors[:TOP_N_VENDORS], 1):
            pct = data["sales"] / total_sales * 100 if total_sales > 0 else 0.0
            lines.append(
                f"  {i}. {vendor} | ${data['sales']:,.0f} ({pct:.1f}%)"
            )
        lines.append("")
        return lines

    def _section_inventory_productivity(self) -> list:
        """
        Section 5: Inventory & Productivity analysis.

        Returns:
            List of formatted lines for this section.
        """
        lines = ["5. INVENTORY & PRODUCTIVITY", SUBSEPARATOR, ""]

        total_sales = self._total_sales()
        total_inv = self._total_inventory()
        ratio = total_sales / total_inv if total_inv > 0 else 0.0
        sell_through = (
            total_sales / (total_sales + total_inv) * 100
            if (total_sales + total_inv) > 0 else 0.0
        )

        lines.append(f"Total Inventory Value: ${total_inv:,.2f}")
        if total_inv > 0:
            lines.append(f"Sales-to-Inventory Ratio: {ratio:.2f}x")
        else:
            lines.append("Sales-to-Inventory Ratio: N/A")
        lines.append(f"Estimated Sell-Through: {sell_through:.1f}%")
        lines.append("")

        if sell_through >= 30:
            lines.append("✓ Inventory productivity supports efficiency improvement")
        else:
            lines.append(
                "⚠ Low sell-through — review stock levels and reorder points"
            )
        lines.append("")
        return lines

    def _section_risks_opportunities(self) -> list:
        """
        Section 6: Risks & Opportunities.

        Returns:
            List of formatted lines for this section.
        """
        lines = ["6. RISKS & OPPORTUNITIES", SUBSEPARATOR, ""]

        cats = self._categories_summary()
        declining_count = sum(
            1 for data in cats.values() if data["avg_growth"] < -10
        )
        growing_count = sum(
            1 for data in cats.values() if data["avg_growth"] > 10
        )

        lines.append("RISKS:")
        lines.append(
            f"  • {declining_count} categories declining >10% YoY"
            " — requires intervention"
        )
        lines.append("  • Inventory management needs optimization")
        lines.append("  • Long-tail assortment diluting profitability")
        lines.append("")

        lines.append("OPPORTUNITIES:")
        lines.append(
            f"  • {growing_count} high-growth categories"
            " — expand assortment and shelf space"
        )
        lines.append("  • Vendor consolidation potential to improve margins")
        lines.append(
            "  • Rationalize low-velocity SKUs to improve operational efficiency"
        )
        lines.append("")
        return lines

    def _section_action_plan(self) -> list:
        """
        Section 7: Action Plan — immediate and mid-term priorities.

        Returns:
            List of formatted lines for this section.
        """
        lines = ["7. ACTION PLAN", SUBSEPARATOR, ""]

        lines.append("IMMEDIATE (0-30 days):")
        lines.append("  • Audit long-tail SKUs for discontinuation")
        lines.append("  • Review declining category assortments")
        lines.append("  • Adjust inventory levels where necessary")
        lines.append("")

        lines.append("MID-TERM (30-90 days):")
        lines.append("  • Reset planograms for core items")
        lines.append("  • Conduct vendor performance reviews")
        lines.append("  • Implement assortment optimizations")
        lines.append("")
        return lines

    def _section_executive_summary(self) -> list:
        """
        Section 8: Executive Summary — strategic narrative.

        Returns:
            List of formatted lines for this section.
        """
        lines = ["8. EXECUTIVE SUMMARY", SUBSEPARATOR, ""]

        total_sales = self._total_sales()
        overall_growth = (
            sum(
                self._safe_float(item.get("sales"))
                * self._safe_float(item.get("yoy_growth"))
                for item in self.data
            ) / total_sales
            if total_sales > 0 else 0.0
        )
        core = self._core_items()

        lines.append(
            f"Category business performance is driven by {len(core)} core items "
            f"representing the top {CORE_SALES_THRESHOLD:.0%} of sales."
        )
        lines.append(
            f"Overall YoY growth stands at {overall_growth:+.1f}%. "
            "Continued focus on high-turn items and rationalization"
        )
        lines.append(
            "of long-tail assortment will improve overall category "
            "productivity and profitability."
        )
        lines.append("")
        return lines
