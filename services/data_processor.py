"""
Data Processor for RIOS
Extracts and analyzes Excel report data
"""

import pandas as pd
from utils.logger import logger
from services.report_generator import ReportGenerator


class DataProcessor:
    """Processes Excel files and generates analysis."""

    def analyze(self, file_path):
        """
        Analyze Excel file and generate reports.
        
        Args:
            file_path: Path to Excel file
            
        Returns:
            dict with 'one_pager' and 'vendor_outreach' keys
        """
        try:
            logger.info(f"📊 Processing file: {file_path}")
            
            # Read Excel file
            df = pd.read_excel(file_path)
            logger.info(f"✅ Loaded {len(df)} rows from Excel")
            
            # Normalize column names
            df.columns = [col.lower().strip().replace(" ", "_") for col in df.columns]
            
            # Extract data structure
            data = self._extract_data(df)
            logger.info(f"✅ Extracted {len(data)} items for analysis")
            
            # Generate reports
            report_gen = ReportGenerator(data)
            
            one_pager = report_gen.generate_one_pager()
            vendor_outreach = report_gen.generate_vendor_outreach()
            
            logger.info("✅ Analysis complete")
            
            return {
                "one_pager": one_pager,
                "vendor_outreach": vendor_outreach,
            }
            
        except Exception as e:
            logger.error(f"Error processing file: {e}", exc_info=True)
            raise

    def _extract_data(self, df):
        """Extract structured data from DataFrame."""
        data = []
        
        for idx, row in df.iterrows():
            item = {
                "category": self._safe_get(row, "category"),
                "vendor": self._safe_get(row, "vendor"),
                "brand": self._safe_get(row, "brand"),
                "sku": self._safe_get(row, "sku"),
                "sales": self._safe_float(row, "sales"),
                "units": self._safe_float(row, "units"),
                "yoy_growth": self._safe_float(row, "yoy_growth"),
                "gross_margin_pct": self._safe_float(row, "gross_margin_%"),
                "inventory": self._safe_float(row, "inventory"),
            }
            
            data.append(item)
        
        return data

    def _safe_get(self, row, col_name):
        """Safely get value from row."""
        try:
            return str(row.get(col_name, "N/A")).strip()
        except:
            return "N/A"

    def _safe_float(self, row, col_name):
        """Safely convert to float."""
        try:
            val = row.get(col_name, 0)
            if pd.isna(val):
                return 0.0
            return float(val)
        except:
            return 0.0
