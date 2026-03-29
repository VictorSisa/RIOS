"""
Excel Data Extractor for RIOS
Modular extraction of structured data from Excel files
"""

import pandas as pd
from pathlib import Path
from typing import Dict, List, Optional
from utils.logger import logger


class ExcelExtractor:
    """Extracts structured data from Excel files."""
    
    def __init__(self, file_path: str):
        """Initialize extractor with file path."""
        self.file_path = Path(file_path)
        self.xls = None
        self._load_file()
    
    def _load_file(self):
        """Load Excel file."""
        try:
            self.xls = pd.ExcelFile(self.file_path)
            logger.info(f"✅ Loaded Excel file: {self.file_path.name}")
        except Exception as e:
            logger.error(f"Error loading Excel file: {e}")
            raise
    
    def _get_column_by_letter(self, df: pd.DataFrame, col_letter: str) -> Optional[str]:
        """Convert column letter to index and return column name."""
        col_index = sum((ord(char) - ord('A') + 1) * 26 ** i 
                       for i, char in enumerate(reversed(col_letter.upper()))) - 1
        
        if col_index < len(df.columns):
            return df.columns[col_index]
        return None
    
    def extract_category_summary(self, sheet_name: str = "Category - Summary") -> pd.DataFrame:
        """Extract category-level summary data."""
        try:
            logger.info(f"Extracting category summary from: {sheet_name}")
            
            df = pd.read_excel(self.xls, sheet_name=sheet_name)
            logger.info(f"   Loaded {len(df)} rows, {len(df.columns)} columns")
            
            column_map = {
                'B': 'category',
                'D': 'sales_ty',
                'F': 'sales_pct_change',
                'G': 'units_ty',
                'I': 'units_pct_change',
                'J': 'aur',
                'L': 'aur_pct_change',
                'S': 'gm_dollars',
                'U': 'gm_pct_change',
                'X': 'inventory',
                'Z': 'inventory_pct_change',
                'AR': 'store_sales',
                'AT': 'store_sales_pct_change',
                'BW': 'ecom_sales',
                'BY': 'ecom_sales_pct_change'
            }
            
            extracted_data = {}
            for col_letter, output_name in column_map.items():
                col_name = self._get_column_by_letter(df, col_letter)
                if col_name is not None:
                    extracted_data[output_name] = df[col_name]
                    logger.info(f"   ✓ Column {col_letter}: {col_name}")
            
            result = pd.DataFrame(extracted_data)
            result = result.dropna(how='all')
            
            if 'category' in result.columns:
                result = result[result['category'].notna()]
            
            logger.info(f"   ✅ Extracted {len(result)} category rows")
            return result
        
        except Exception as e:
            logger.error(f"Error extracting category summary: {e}")
            raise
    
    def extract_vendor_brand_summary(self, sheet_name: str = "Vendor->Brand - Summary") -> pd.DataFrame:
        """Extract vendor-brand-level summary data with inventory metrics."""
        try:
            logger.info(f"Extracting vendor-brand summary from: {sheet_name}")
            
            df = pd.read_excel(self.xls, sheet_name=sheet_name)
            logger.info(f"   Loaded {len(df)} rows, {len(df.columns)} columns")
            
            column_map = {
                'A': 'vendor',
                'C': 'brand',
                'D': 'sales_ty',
                'F': 'sales_pct_change',
                'G': 'units_ty',
                'I': 'units_pct_change',
                'J': 'aur',
                'L': 'aur_pct_change',
                'S': 'gm_dollars',
                'U': 'gm_pct_change',
                'X': 'inventory',
                'Z': 'inventory_pct_change'
            }
            
            extracted_data = {}
            for col_letter, output_name in column_map.items():
                col_name = self._get_column_by_letter(df, col_letter)
                if col_name is not None:
                    extracted_data[output_name] = df[col_name]
                    logger.info(f"   ✓ Column {col_letter}: {col_name}")
            
            result = pd.DataFrame(extracted_data)
            result = result.dropna(how='all')
            
            if 'vendor' in result.columns:
                result = result[result['vendor'].notna()]
            
            logger.info(f"   ✅ Extracted {len(result)} vendor-brand rows")
            return result
        
        except Exception as e:
            logger.error(f"Error extracting vendor-brand summary: {e}")
            raise
    
    def extract_vendor_summary(self, sheet_name: str = "Vendor->Brand - Summary") -> pd.DataFrame:
        """Extract vendor-level summary data (aggregated)."""
        try:
            logger.info(f"Extracting vendor summary from: {sheet_name}")
            
            vb_data = self.extract_vendor_brand_summary(sheet_name)
            
            vb_data = vb_data[
                ~vb_data['vendor'].str.contains('Subtotal', na=False)
            ].copy()
            
            for col in ['sales_ty', 'gm_dollars', 'inventory', 'inventory_pct_change']:
                vb_data[col] = pd.to_numeric(vb_data[col], errors='coerce')
            
            vendor_summary = vb_data.groupby('vendor').agg({
                'sales_ty': 'sum',
                'gm_dollars': 'sum',
                'inventory': 'sum',
                'inventory_pct_change': 'mean'
            }).reset_index()
            
            logger.info(f"   ✅ Aggregated to {len(vendor_summary)} vendors")
            return vendor_summary
        
        except Exception as e:
            logger.error(f"Error extracting vendor summary: {e}")
            raise


def extract_category_data(file_path: str) -> pd.DataFrame:
    """Quick function to extract category summary data."""
    extractor = ExcelExtractor(file_path)
    return extractor.extract_category_summary()


def extract_vendor_brand_data(file_path: str) -> pd.DataFrame:
    """Quick function to extract vendor-brand summary data."""
    extractor = ExcelExtractor(file_path)
    return extractor.extract_vendor_brand_summary()


def extract_vendor_data(file_path: str) -> pd.DataFrame:
    """Quick function to extract vendor summary data."""
    extractor = ExcelExtractor(file_path)
    return extractor.extract_vendor_summary()

