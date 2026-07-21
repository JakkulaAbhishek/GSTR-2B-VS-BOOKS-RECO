# ============================================================================
# GST Recon Pro v9.2.3 - Enterprise Multi-Strategy GST Reconciliation Engine
# ============================================================================
# Fixed: Strict matching criteria, no false positives
# Author: Abhishek Jakkula
# Email: jakkulaabhishek5@gmail.com
# Version: 9.2.3 (Strict • Confidence‑Filtered • Credit Note Support)
# Last Updated: July 2026
# License: Proprietary - Enterprise Edition
# ============================================================================

import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import warnings
import logging
import sys
import time
import json
import base64
import os
import traceback
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple, Union, Any, Set
from dataclasses import dataclass, field
from enum import Enum
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import xlsxwriter
from io import BytesIO
from difflib import SequenceMatcher
from collections import defaultdict, Counter
import functools
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import queue
import random
import hashlib
import gc

# Suppress warnings
warnings.filterwarnings('ignore')
pd.options.mode.chained_assignment = None
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', 100)
pd.set_option('display.float_format', lambda x: '%.2f' % x)

# ============================================================================
# VERSION INFORMATION
# ============================================================================

VERSION = "9.3.0"
VERSION_NAME = "Data Quality, ITC Hints & CLI Mode"
BUILD_DATE = "2026-07-22"
ENHANCEMENTS = [
    "Stricter matching thresholds (doc sim ≥0.7–0.95)",
    "Tax difference limit reduced to 1×tolerance",
    "Aggregate matching only 1‑to‑1 with doc sim ≥0.7",
    "Confidence filtering (min 0.5) eliminates weak matches",
    "Credit note matching requires reference doc sim ≥0.8",
    "Negative value matching requires doc sim ≥0.7",
    "No false positives – only invoices present in both datasets are matched"
]

# ============================================================================
# LOGGING CONFIGURATION
# ============================================================================

class LoggerSetup:
    """Centralized logging configuration"""
    
    _instance = None
    _lock = threading.Lock()
    
    def __new__(cls):
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super(LoggerSetup, cls).__new__(cls)
                    cls._instance._initialize()
        return cls._instance
    
    def _initialize(self):
        self.logger = logging.getLogger('GSTReconPro')
        self.logger.setLevel(logging.INFO)
        
        for handler in self.logger.handlers[:]:
            self.logger.removeHandler(handler)
        
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.INFO)
        console_format = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        console_handler.setFormatter(console_format)
        self.logger.addHandler(console_handler)
        
        try:
            file_handler = logging.FileHandler('gst_recon.log')
            file_handler.setLevel(logging.DEBUG)
            file_format = logging.Formatter(
                '%(asctime)s - %(name)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s'
            )
            file_handler.setFormatter(file_format)
            self.logger.addHandler(file_handler)
        except:
            pass
    
    def get_logger(self):
        return self.logger

# ============================================================================
# SAMPLE DATA GENERATOR WITH CREDIT NOTE SUPPORT
# ============================================================================

class SampleDataGenerator:
    """Generate sample GST data with credit note support"""
    
    @staticmethod
    def generate_gstin() -> str:
        """Generate a valid GSTIN format"""
        state_code = f"{random.randint(1, 37):02d}"
        pan = ''.join(random.choices('ABCDEFGHIJKLMNOPQRSTUVWXYZ', k=5)) + \
              ''.join(random.choices('0123456789', k=4)) + \
              random.choice('ABCDEFGHIJKLMNOPQRSTUVWXYZ')
        entity = random.choice('0123456789')
        return f"{state_code}{pan}{entity}Z{random.choice('0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ')}"
    
    @staticmethod
    def generate_invoice_number(doc_type: str = 'INVOICE') -> str:
        """Generate a realistic document number"""
        prefixes = {
            'INVOICE': ['INV', 'BILL', 'INV-', 'GST-', 'TAX-', 'SALES-', 'PUR-', ''],
            'CREDIT NOTE': ['CN', 'CRN', 'CN-', 'CREDIT-', 'CR-NOTE-', ''],
            'DEBIT NOTE': ['DN', 'DBN', 'DN-', 'DEBIT-', 'DB-NOTE-', '']
        }
        
        prefix_list = prefixes.get(doc_type, prefixes['INVOICE'])
        prefix = random.choice(prefix_list)
        year = random.choice(['2024', '2025', '2026'])
        month = f"{random.randint(1, 12):02d}"
        number = f"{random.randint(1, 9999):04d}"
        
        formats = [
            f"{prefix}{year}{month}{number}",
            f"{prefix}{number}/{year}-{month}",
            f"{prefix}{year}-{month}-{number}",
            f"{prefix}{number}",
            f"{prefix}{random.randint(100000, 999999)}"
        ]
        
        if doc_type == 'CREDIT NOTE':
            formats.extend([
                f"CN-{year}-{month}-{number}",
                f"CRN{number}",
                f"CREDIT-{number}"
            ])
        elif doc_type == 'DEBIT NOTE':
            formats.extend([
                f"DN-{year}-{month}-{number}",
                f"DBN{number}",
                f"DEBIT-{number}"
            ])
        
        return random.choice(formats)
    
    @staticmethod
    def generate_supplier_name() -> str:
        """Generate a realistic supplier name"""
        prefixes = ['M/s ', 'Shri ', '', '']
        names = [
            'ABC Traders', 'XYZ Enterprises', 'Ram & Sons', 'Sita Suppliers',
            'Ganesh Agencies', 'Lakshmi Enterprises', 'Saraswati Traders',
            'Durga Suppliers', 'Krishna Enterprises', 'Rama Traders',
            'Shiva Agencies', 'Parvati Suppliers', 'Ganpati Enterprises',
            'Sai Traders', 'Baba Agencies', 'Mata Suppliers',
            'Bharat Enterprises', 'India Traders', 'Asian Suppliers',
            'Global Enterprises', 'United Traders', 'National Agencies',
            'City Suppliers', 'Town Enterprises', 'Metro Traders'
        ]
        return random.choice(prefixes) + random.choice(names)
    
    @staticmethod
    def generate_amount(min_val: float = 1000, max_val: float = 100000, negative: bool = False) -> float:
        """Generate a random amount (positive or negative)"""
        amount = round(random.uniform(min_val, max_val), 2)
        if negative:
            amount = -amount
        return amount
    
    @staticmethod
    def generate_date(start_date: str = '2024-01-01', end_date: str = '2026-12-31') -> datetime:
        """Generate a random date"""
        start = datetime.strptime(start_date, '%Y-%m-%d')
        end = datetime.strptime(end_date, '%Y-%m-%d')
        delta = end - start
        random_days = random.randint(0, delta.days)
        return start + timedelta(days=random_days)
    
    @staticmethod
    def get_reference_invoice(gstr_data: pd.DataFrame, idx: int) -> Optional[str]:
        """Get a reference invoice number for credit notes"""
        if not gstr_data.empty and idx < len(gstr_data):
            return gstr_data.iloc[idx]['DOCUMENT NUMBER']
        return f"INV-{random.randint(1000, 9999)}"
    
    @staticmethod
    def generate_gstr_2b_data(num_records: int = 100, credit_note_ratio: float = 0.15) -> pd.DataFrame:
        """Generate sample GSTR-2B data with credit notes"""
        data = []
        
        num_credit_notes = int(num_records * credit_note_ratio)
        num_invoices = num_records - num_credit_notes
        
        for i in range(num_invoices):
            gstin = SampleDataGenerator.generate_gstin()
            invoice = SampleDataGenerator.generate_invoice_number('INVOICE')
            date = SampleDataGenerator.generate_date()
            taxable = SampleDataGenerator.generate_amount(1000, 50000, negative=False)
            
            igst = round(taxable * random.uniform(0.05, 0.18), 2) if random.choice([True, False]) else 0
            cgst = round(taxable * random.uniform(0.025, 0.09), 2) if igst == 0 else 0
            sgst = cgst if igst == 0 else 0
            
            data.append({
                'SUPPLIER GSTIN': gstin,
                'SUPPLIER NAME': SampleDataGenerator.generate_supplier_name(),
                'DOCUMENT NUMBER': invoice,
                'DOCUMENT DATE': date.strftime('%d-%m-%Y'),
                'TAXABLE VALUE': taxable,
                'IGST': igst,
                'CGST': cgst,
                'SGST': sgst,
                'TOTAL TAX': igst + cgst + sgst,
                'TOTAL VALUE': taxable + igst + cgst + sgst,
                'DOC TYPE': 'INVOICE',
                'REFERENCE DOCUMENT': '',
                'ORIGINAL AMOUNT': taxable
            })
        
        for i in range(num_credit_notes):
            gstin = SampleDataGenerator.generate_gstin()
            ref_invoice = SampleDataGenerator.generate_invoice_number('INVOICE')
            cn_number = SampleDataGenerator.generate_invoice_number('CREDIT NOTE')
            date = SampleDataGenerator.generate_date()
            
            original_taxable = SampleDataGenerator.generate_amount(1000, 50000, negative=False)
            credit_percentage = random.uniform(0.1, 1.0)
            taxable = -round(original_taxable * credit_percentage, 2)
            
            igst = round(taxable * random.uniform(0.05, 0.18), 2) if random.choice([True, False]) else 0
            cgst = round(taxable * random.uniform(0.025, 0.09), 2) if igst == 0 else 0
            sgst = cgst if igst == 0 else 0
            
            data.append({
                'SUPPLIER GSTIN': gstin,
                'SUPPLIER NAME': SampleDataGenerator.generate_supplier_name(),
                'DOCUMENT NUMBER': cn_number,
                'DOCUMENT DATE': date.strftime('%d-%m-%Y'),
                'TAXABLE VALUE': taxable,
                'IGST': igst,
                'CGST': cgst,
                'SGST': sgst,
                'TOTAL TAX': igst + cgst + sgst,
                'TOTAL VALUE': taxable + igst + cgst + sgst,
                'DOC TYPE': 'CREDIT NOTE',
                'REFERENCE DOCUMENT': ref_invoice,
                'ORIGINAL AMOUNT': original_taxable
            })
        
        num_debit_notes = int(num_records * 0.03)
        for i in range(num_debit_notes):
            gstin = SampleDataGenerator.generate_gstin()
            dn_number = SampleDataGenerator.generate_invoice_number('DEBIT NOTE')
            date = SampleDataGenerator.generate_date()
            taxable = SampleDataGenerator.generate_amount(1000, 50000, negative=True)
            
            igst = round(taxable * random.uniform(0.05, 0.18), 2) if random.choice([True, False]) else 0
            cgst = round(taxable * random.uniform(0.025, 0.09), 2) if igst == 0 else 0
            sgst = cgst if igst == 0 else 0
            
            data.append({
                'SUPPLIER GSTIN': gstin,
                'SUPPLIER NAME': SampleDataGenerator.generate_supplier_name(),
                'DOCUMENT NUMBER': dn_number,
                'DOCUMENT DATE': date.strftime('%d-%m-%Y'),
                'TAXABLE VALUE': taxable,
                'IGST': igst,
                'CGST': cgst,
                'SGST': sgst,
                'TOTAL TAX': igst + cgst + sgst,
                'TOTAL VALUE': taxable + igst + cgst + sgst,
                'DOC TYPE': 'DEBIT NOTE',
                'REFERENCE DOCUMENT': '',
                'ORIGINAL AMOUNT': abs(taxable)
            })
        
        random.shuffle(data)
        return pd.DataFrame(data)
    
    @staticmethod
    def generate_purchase_register_data(df_2b: Optional[pd.DataFrame] = None, num_records: int = 100,
                                         match_rate: float = 0.85, credit_note_ratio: float = 0.15) -> pd.DataFrame:
        """Generate sample Purchase Register data with credit notes.

        IMPORTANT: To produce a PR sample that actually overlaps with a GSTR-2B sample
        (so the two files can be reconciled meaningfully), pass the SAME df_2b dataframe
        that was used/shown as the GSTR-2B sample. If df_2b is not provided, a fresh
        GSTR-2B dataset is generated internally (backward-compatible fallback), but note
        that it will NOT match any separately-generated GSTR-2B sample.
        """
        if df_2b is not None:
            gstr_data = df_2b.reset_index(drop=True)
        else:
            gstr_data = SampleDataGenerator.generate_gstr_2b_data(num_records, credit_note_ratio)
        
        data = []
        matched_indices = set()
        
        for idx, row in gstr_data.iterrows():
            if random.random() < match_rate:
                taxable = row['TAXABLE VALUE']
                
                if random.random() < 0.15:
                    variation = random.uniform(-50, 50)
                    taxable = taxable + variation
                    if row['DOC TYPE'] == 'CREDIT NOTE' and taxable > 0:
                        taxable = -taxable
                
                date = datetime.strptime(row['DOCUMENT DATE'], '%d-%m-%Y')
                if random.random() < 0.10:
                    date = date + timedelta(days=random.randint(-5, 5))
                
                if taxable != 0:
                    igst = round(taxable * random.uniform(0.05, 0.18), 2) if row['IGST'] != 0 else 0
                    cgst = round(taxable * random.uniform(0.025, 0.09), 2) if row['CGST'] != 0 else 0
                    sgst = cgst if row['SGST'] != 0 else 0
                else:
                    igst = cgst = sgst = 0
                
                data.append({
                    'SUPPLIER GSTIN': row['SUPPLIER GSTIN'],
                    'SUPPLIER NAME': row['SUPPLIER NAME'],
                    'DOCUMENT NUMBER': row['DOCUMENT NUMBER'],
                    'DOCUMENT DATE': date.strftime('%d-%m-%Y'),
                    'TAXABLE VALUE': taxable,
                    'IGST': igst,
                    'CGST': cgst,
                    'SGST': sgst,
                    'TOTAL TAX': igst + cgst + sgst,
                    'TOTAL VALUE': taxable + igst + cgst + sgst,
                    'DOC TYPE': row['DOC TYPE'],
                    'REFERENCE DOCUMENT': row.get('REFERENCE DOCUMENT', ''),
                    'ORIGINAL AMOUNT': row.get('ORIGINAL AMOUNT', abs(taxable))
                })
                matched_indices.add(idx)
        
        additional_needed = num_records - len(data)
        for i in range(max(0, additional_needed)):
            doc_type = random.choice(['INVOICE', 'CREDIT NOTE', 'DEBIT NOTE'])
            is_credit = doc_type in ['CREDIT NOTE', 'DEBIT NOTE']
            
            taxable = SampleDataGenerator.generate_amount(1000, 50000, negative=is_credit)
            
            if is_credit:
                igst = round(taxable * random.uniform(0.05, 0.18), 2) if random.choice([True, False]) else 0
                cgst = round(taxable * random.uniform(0.025, 0.09), 2) if igst == 0 else 0
                sgst = cgst if igst == 0 else 0
            else:
                igst = round(taxable * random.uniform(0.05, 0.18), 2) if random.choice([True, False]) else 0
                cgst = round(taxable * random.uniform(0.025, 0.09), 2) if igst == 0 else 0
                sgst = cgst if igst == 0 else 0
            
            data.append({
                'SUPPLIER GSTIN': SampleDataGenerator.generate_gstin(),
                'SUPPLIER NAME': SampleDataGenerator.generate_supplier_name(),
                'DOCUMENT NUMBER': SampleDataGenerator.generate_invoice_number(doc_type),
                'DOCUMENT DATE': SampleDataGenerator.generate_date().strftime('%d-%m-%Y'),
                'TAXABLE VALUE': taxable,
                'IGST': igst,
                'CGST': cgst,
                'SGST': sgst,
                'TOTAL TAX': igst + cgst + sgst,
                'TOTAL VALUE': taxable + igst + cgst + sgst,
                'DOC TYPE': doc_type,
                'REFERENCE DOCUMENT': SampleDataGenerator.generate_invoice_number('INVOICE') if is_credit else '',
                'ORIGINAL AMOUNT': abs(taxable)
            })
        
        return pd.DataFrame(data)
    
    @staticmethod
    def create_sample_excel_files():
        """Create sample Excel files for download with credit notes"""
        df_2b = SampleDataGenerator.generate_gstr_2b_data(100, credit_note_ratio=0.15)
        df_pr = SampleDataGenerator.generate_purchase_register_data(df_2b, match_rate=0.85, credit_note_ratio=0.15)
        
        output_2b = BytesIO()
        with pd.ExcelWriter(output_2b, engine='xlsxwriter') as writer:
            df_2b.to_excel(writer, sheet_name='GSTR-2B', index=False)
            
            workbook = writer.book
            worksheet = writer.sheets['GSTR-2B']
            
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#1e293b',
                'font_color': 'white',
                'border': 1
            })
            
            credit_format = workbook.add_format({
                'bg_color': '#fee2e2',
                'font_color': '#dc2626'
            })
            
            for col_num, value in enumerate(df_2b.columns.values):
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, 18)
            
            for row_num in range(len(df_2b)):
                doc_type = df_2b.iloc[row_num]['DOC TYPE']
                if doc_type == 'CREDIT NOTE':
                    for col_num in range(len(df_2b.columns)):
                        worksheet.write(row_num + 1, col_num, df_2b.iloc[row_num, col_num], credit_format)
        
        output_pr = BytesIO()
        with pd.ExcelWriter(output_pr, engine='xlsxwriter') as writer:
            df_pr.to_excel(writer, sheet_name='Purchase_Register', index=False)
            
            workbook = writer.book
            worksheet = writer.sheets['Purchase_Register']
            
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#1e293b',
                'font_color': 'white',
                'border': 1
            })
            
            credit_format = workbook.add_format({
                'bg_color': '#fee2e2',
                'font_color': '#dc2626'
            })
            
            for col_num, value in enumerate(df_pr.columns.values):
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, 18)
            
            for row_num in range(len(df_pr)):
                doc_type = df_pr.iloc[row_num]['DOC TYPE']
                if doc_type == 'CREDIT NOTE':
                    for col_num in range(len(df_pr.columns)):
                        worksheet.write(row_num + 1, col_num, df_pr.iloc[row_num, col_num], credit_format)
        
        return output_2b.getvalue(), output_pr.getvalue()
    
    @staticmethod
    def get_credit_note_summary(df: pd.DataFrame) -> Dict:
        """Get credit note summary statistics"""
        if df.empty or 'DOC TYPE' not in df.columns:
            return {
                'total_credit_notes': 0,
                'total_credit_amount': 0,
                'total_invoices': len(df)
            }
        
        credit_df = df[df['DOC TYPE'].str.upper().str.contains('CREDIT', na=False)]
        return {
            'total_credit_notes': len(credit_df),
            'total_credit_amount': credit_df['TAXABLE VALUE'].sum() if not credit_df.empty else 0,
            'total_invoices': len(df[~df['DOC TYPE'].str.upper().str.contains('CREDIT', na=False)])
        }

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def extract_pan_from_gstin(gstin: str) -> str:
    """Extract PAN from GSTIN (characters 3-12)"""
    if pd.isna(gstin) or len(str(gstin).strip()) < 15:
        return "UNKNOWN"
    gstin_str = str(gstin).strip().upper()
    if len(gstin_str) >= 12:
        return gstin_str[2:12]
    return "UNKNOWN"

def normalize_document_number(doc_num: str) -> str:
    """Normalize document number by removing special characters"""
    if pd.isna(doc_num) or str(doc_num).strip() == "":
        return "UNKNOWN"
    normalized = re.sub(r'[^A-Z0-9]', '', str(doc_num).upper().strip())
    return normalized.lstrip('0') or "0"

def parse_date(date_str: str) -> Optional[datetime]:
    """Parse date from various formats"""
    if pd.isna(date_str) or str(date_str).strip() == "":
        return None
    
    date_str = str(date_str).strip()
    
    formats = [
        '%d-%m-%Y', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y',
        '%d-%b-%Y', '%d %b %Y', '%b %d, %Y',
        '%Y/%m/%d', '%d.%m.%Y', '%m.%d.%Y',
        '%d-%m-%y', '%d/%m/%y', '%m/%d/%y',
        '%Y%m%d', '%d%m%Y', '%d%m%y'
    ]
    
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    
    try:
        parsed = pd.to_datetime(date_str, errors='coerce')
        if pd.notna(parsed):
            return parsed.to_pydatetime()
    except:
        pass
    
    return None

def validate_gstin_format(gstin: str) -> bool:
    """Validate GSTIN format (15 characters, structural pattern only)"""
    if pd.isna(gstin) or len(str(gstin).strip()) != 15:
        return False
    gstin = str(gstin).strip().upper()
    pattern = r'^[0-9]{2}[A-Z0-9]{10}[0-9]Z[A-Z0-9]{1}$'
    return bool(re.match(pattern, gstin))

# GSTIN checksum alphabet used by GSTN's check-digit algorithm
_GSTIN_CODEPOINTS = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
_GSTIN_CODEPOINT_INDEX = {ch: i for i, ch in enumerate(_GSTIN_CODEPOINTS)}

# State code -> State name mapping (as per GST state jurisdiction codes)
GST_STATE_CODES = {
    '01': 'Jammu and Kashmir', '02': 'Himachal Pradesh', '03': 'Punjab',
    '04': 'Chandigarh', '05': 'Uttarakhand', '06': 'Haryana', '07': 'Delhi',
    '08': 'Rajasthan', '09': 'Uttar Pradesh', '10': 'Bihar', '11': 'Sikkim',
    '12': 'Arunachal Pradesh', '13': 'Nagaland', '14': 'Manipur', '15': 'Mizoram',
    '16': 'Tripura', '17': 'Meghalaya', '18': 'Assam', '19': 'West Bengal',
    '20': 'Jharkhand', '21': 'Odisha', '22': 'Chhattisgarh', '23': 'Madhya Pradesh',
    '24': 'Gujarat', '25': 'Daman and Diu', '26': 'Dadra and Nagar Haveli',
    '27': 'Maharashtra', '28': 'Andhra Pradesh (Old)', '29': 'Karnataka',
    '30': 'Goa', '31': 'Lakshadweep', '32': 'Kerala', '33': 'Tamil Nadu',
    '34': 'Puducherry', '35': 'Andaman and Nicobar Islands', '36': 'Telangana',
    '37': 'Andhra Pradesh', '38': 'Ladakh', '97': 'Other Territory',
}

def validate_gstin_checksum(gstin: str) -> Dict[str, Any]:
    """
    Validate the GSTIN check-digit (15th character) using GSTN's published
    check-digit algorithm. This catches typos/OCR errors that pass the
    structural regex but are not real GSTINs.

    Returns a dict: {
        'format_valid': bool,      # structural pattern check
        'checksum_valid': bool,    # actual check-digit match
        'state_code': str | None,
        'state_name': str | None,
        'error': str | None
    }
    """
    result = {
        'format_valid': False,
        'checksum_valid': False,
        'state_code': None,
        'state_name': None,
        'error': None
    }

    if pd.isna(gstin):
        result['error'] = 'GSTIN is empty'
        return result

    g = str(gstin).strip().upper()

    if len(g) != 15:
        result['error'] = f'GSTIN length is {len(g)}, expected 15'
        return result

    if not validate_gstin_format(g):
        result['error'] = 'GSTIN does not match expected structural pattern'
        return result

    result['format_valid'] = True
    result['state_code'] = g[0:2]
    result['state_name'] = GST_STATE_CODES.get(g[0:2], 'Unknown/Invalid State Code')

    try:
        total = 0
        for i in range(14):
            ch = g[i]
            if ch not in _GSTIN_CODEPOINT_INDEX:
                result['error'] = f'Invalid character "{ch}" at position {i+1}'
                return result
            code = _GSTIN_CODEPOINT_INDEX[ch]
            factor = 1 if i % 2 == 0 else 2
            digit = factor * code
            digit = (digit // 36) + (digit % 36)
            total += digit

        checksum_index = (36 - (total % 36)) % 36
        expected_check_char = _GSTIN_CODEPOINTS[checksum_index]

        result['checksum_valid'] = (g[14] == expected_check_char)
        if not result['checksum_valid']:
            result['error'] = (
                f'Check digit mismatch: expected "{expected_check_char}", found "{g[14]}". '
                f'This GSTIN may contain a typo.'
            )
    except Exception as e:
        result['error'] = f'Checksum computation failed: {str(e)}'

    return result

def get_gstin_state_name(gstin: str) -> str:
    """Quick helper to get state name from a GSTIN's state code prefix"""
    if pd.isna(gstin) or len(str(gstin).strip()) < 2:
        return 'Unknown'
    code = str(gstin).strip()[0:2]
    return GST_STATE_CODES.get(code, 'Unknown/Invalid')

def safe_float_convert(value: Any) -> float:
    """Safely convert value to float"""
    if pd.isna(value) or value is None:
        return 0.0
    try:
        if isinstance(value, str):
            cleaned = re.sub(r'[^\d.\-]', '', value.strip())
            if cleaned:
                return float(cleaned)
            return 0.0
        return float(value)
    except (ValueError, TypeError):
        return 0.0

def safe_string_convert(value: Any) -> str:
    """Safely convert value to string"""
    if pd.isna(value) or value is None:
        return ""
    return str(value).strip()

def get_quarter(month: int) -> int:
    """Get quarter from month number"""
    if pd.isna(month):
        return None
    return (month - 1) // 3 + 1

def is_credit_note(row) -> bool:
    """Check if a row represents a credit note"""
    if pd.isna(row) or row is None:
        return False
    if isinstance(row, dict):
        doc_type = str(row.get('DOC TYPE', '')).upper()
        taxable = row.get('TAXABLE VALUE', 0)
    else:
        doc_type = str(row.get('DOC TYPE', '')).upper()
        taxable = row.get('TAXABLE VALUE', 0)
    return doc_type == 'CREDIT NOTE' or taxable < 0

def is_debit_note(row) -> bool:
    """Check if a row represents a debit note"""
    if pd.isna(row) or row is None:
        return False
    if isinstance(row, dict):
        doc_type = str(row.get('DOC TYPE', '')).upper()
    else:
        doc_type = str(row.get('DOC TYPE', '')).upper()
    return doc_type == 'DEBIT NOTE'

def get_document_type(row) -> str:
    """Get normalized document type"""
    if pd.isna(row) or row is None:
        return 'UNKNOWN'
    if isinstance(row, dict):
        doc_type = str(row.get('DOC TYPE', '')).upper()
        taxable = row.get('TAXABLE VALUE', 0)
    else:
        doc_type = str(row.get('DOC TYPE', '')).upper()
        taxable = row.get('TAXABLE VALUE', 0)
    
    if doc_type == 'CREDIT NOTE':
        return 'CREDIT NOTE'
    elif doc_type == 'DEBIT NOTE':
        return 'DEBIT NOTE'
    elif taxable < 0:
        return 'CREDIT NOTE (IMPLIED)'
    else:
        return 'INVOICE'

def get_credit_note_amount(row) -> float:
    """Get the credit note amount (absolute value)"""
    if pd.isna(row) or row is None:
        return 0.0
    if isinstance(row, dict):
        return abs(row.get('TAXABLE VALUE', 0))
    return abs(row.get('TAXABLE VALUE', 0))

# ============================================================================
# ENUMS & DATA CLASSES
# ============================================================================

class MatchStrategy(Enum):
    """Enterprise matching strategy types"""
    EXACT = "exact"
    SMART = "smart"
    VALUE_BASED = "value_based"
    FUZZY_NAME = "fuzzy_name"
    PATTERN_RECOGNITION = "pattern_recognition"
    SEQUENTIAL = "sequential"
    AGGREGATE = "aggregate"
    PERCENTAGE = "percentage"
    WILDCARD = "wildcard"
    AI_ENHANCED = "ai_enhanced"
    CREDIT_NOTE = "credit_note"
    NEGATIVE_VALUE = "negative_value"

class MatchTier(Enum):
    """Tier levels for matching confidence"""
    TIER_1_EXACT = 1
    TIER_2_SMART = 2
    TIER_3_VALUE = 3
    TIER_4_FUZZY = 4
    TIER_5_PATTERN = 5
    TIER_6_SEQUENTIAL = 6
    TIER_7_AGGREGATE = 7
    TIER_8_CREDIT_NOTE = 8
    TIER_9_NEGATIVE = 9

@dataclass
class ReconciliationConfig:
    """Enterprise configuration with credit note support"""
    tolerance_amount: float = 20.0
    date_tolerance_days: int = 7
    fuzzy_threshold: float = 85.0
    enable_reverse_charge: bool = True
    enable_auto_claim: bool = True
    enable_fuzzy_matching: bool = True
    enable_pattern_recognition: bool = True
    enable_sequential_matching: bool = True
    enable_aggregate_matching: bool = True
    enable_percentage_matching: bool = True
    enable_wildcard_matching: bool = True
    enable_ai_enhanced: bool = False
    validate_gstin: bool = True
    strict_financial_year: bool = False
    treat_cdn_negative: bool = True
    max_workers: int = 4
    batch_size: int = 1000
    percentage_tolerance: float = 5.0
    enable_multi_strategy_voting: bool = True
    voting_threshold: int = 3
    enable_parallel_processing: bool = True
    enable_caching: bool = True
    enable_credit_note_matching: bool = True
    enable_negative_value_matching: bool = True
    match_credit_notes_separately: bool = True
    treat_negative_as_credit: bool = True
    min_confidence_for_match: float = 0.5  # New: minimum confidence to accept a match

@dataclass
class ProcessingMetrics:
    """Metrics for processing with credit note tracking"""
    start_time: datetime = field(default_factory=datetime.now)
    end_time: Optional[datetime] = None
    total_records_2b: int = 0
    total_records_pr: int = 0
    matched_records: int = 0
    unmatched_2b: int = 0
    unmatched_pr: int = 0
    strategy_usage: Dict[str, int] = field(default_factory=dict)
    tier_distribution: Dict[int, int] = field(default_factory=dict)
    confidence_distribution: Dict[str, float] = field(default_factory=dict)
    processing_time_seconds: float = 0.0
    errors: List[str] = field(default_factory=list)
    memory_usage_mb: float = 0.0
    credit_notes_2b: int = 0
    credit_notes_pr: int = 0
    matched_credit_notes: int = 0
    unmatched_credit_notes_2b: int = 0
    unmatched_credit_notes_pr: int = 0
    total_credit_amount_2b: float = 0.0
    total_credit_amount_pr: float = 0.0
    credit_note_difference: float = 0.0

# ============================================================================
# DATA PROCESSING UTILITIES
# ============================================================================

class DataQualityAnalyzer:
    """
    Pre-flight data quality checks, run BEFORE reconciliation.
    Surfaces problems that would otherwise silently degrade match quality
    or produce misleading results (bad GSTINs, duplicate invoices,
    inconsistent negative/credit-note flags, unparseable dates, etc.)
    """

    @staticmethod
    def analyze(df: pd.DataFrame, source_name: str = 'Dataset') -> Dict[str, Any]:
        """Run a full suite of quality checks on a single dataframe"""
        report = {
            'source': source_name,
            'total_records': len(df),
            'issues': [],
            'warnings': [],
            'gstin_checksum_failures': [],
            'duplicate_invoices': [],
            'unparseable_dates': 0,
            'negative_without_cn_flag': 0,
            'missing_supplier_name': 0,
            'zero_value_records': 0,
            'state_distribution': {},
        }

        if df.empty:
            report['issues'].append('Dataset is empty')
            return report

        # --- GSTIN checksum validation (sampled for large files to stay fast) ---
        if 'SUPPLIER GSTIN' in df.columns:
            unique_gstins = df['SUPPLIER GSTIN'].dropna().astype(str).unique()
            failures = []
            state_counts = Counter()
            for gstin in unique_gstins:
                res = validate_gstin_checksum(gstin)
                if res['state_name']:
                    state_counts[res['state_name']] += 1
                if res['format_valid'] and not res['checksum_valid']:
                    failures.append({'gstin': gstin, 'reason': res['error']})
                elif not res['format_valid']:
                    failures.append({'gstin': gstin, 'reason': res['error']})

            report['gstin_checksum_failures'] = failures[:100]  # cap for readability
            report['gstin_checksum_failure_count'] = len(failures)
            report['state_distribution'] = dict(state_counts.most_common())

            if failures:
                report['issues'].append(
                    f"{len(failures)} unique GSTIN(s) fail format/checksum validation "
                    f"out of {len(unique_gstins)} unique GSTINs"
                )

        # --- Duplicate invoice detection (same GSTIN + doc number appearing >1x) ---
        if 'SUPPLIER GSTIN' in df.columns and 'DOCUMENT NUMBER' in df.columns:
            dup_key = df['SUPPLIER GSTIN'].astype(str).str.upper().str.strip() + '||' + \
                      df['DOCUMENT NUMBER'].apply(normalize_document_number)
            dup_counts = dup_key.value_counts()
            dup_keys = dup_counts[dup_counts > 1]
            if not dup_keys.empty:
                dup_records = []
                for key, count in dup_keys.items():
                    gstin_part, doc_part = key.split('||', 1)
                    dup_records.append({
                        'gstin': gstin_part,
                        'normalized_doc_number': doc_part,
                        'occurrence_count': int(count)
                    })
                report['duplicate_invoices'] = dup_records[:100]
                report['duplicate_invoice_count'] = len(dup_records)
                report['warnings'].append(
                    f"{len(dup_records)} duplicate GSTIN+invoice-number combination(s) found "
                    f"— duplicate submissions can distort match rates"
                )

        # --- Date parseability ---
        if 'DOCUMENT DATE' in df.columns:
            parsed = df['DOCUMENT DATE'].apply(parse_date)
            unparseable = parsed.isna().sum()
            report['unparseable_dates'] = int(unparseable)
            if unparseable > 0:
                report['warnings'].append(
                    f"{unparseable} record(s) have a document date that could not be parsed"
                )

        # --- Negative taxable value without a DOC TYPE flag ---
        if 'TAXABLE VALUE' in df.columns:
            taxable_numeric = pd.to_numeric(df['TAXABLE VALUE'], errors='coerce').fillna(0)
            negative_mask = taxable_numeric < 0
            if 'DOC TYPE' in df.columns:
                flagged_mask = df['DOC TYPE'].astype(str).str.upper().str.contains('CREDIT|DEBIT', na=False)
                inconsistent = negative_mask & (~flagged_mask)
            else:
                inconsistent = negative_mask
            report['negative_without_cn_flag'] = int(inconsistent.sum())
            if report['negative_without_cn_flag'] > 0:
                report['warnings'].append(
                    f"{report['negative_without_cn_flag']} record(s) have negative taxable value "
                    f"but no CREDIT NOTE/DEBIT NOTE doc-type flag — will be auto-treated as credit notes"
                )
            report['zero_value_records'] = int((taxable_numeric == 0).sum())
            if report['zero_value_records'] > 0:
                report['warnings'].append(
                    f"{report['zero_value_records']} record(s) have zero taxable value"
                )

        # --- Missing supplier name ---
        if 'SUPPLIER NAME' in df.columns:
            missing_name = df['SUPPLIER NAME'].apply(
                lambda x: pd.isna(x) or str(x).strip() == ''
            ).sum()
            report['missing_supplier_name'] = int(missing_name)
            if missing_name > 0:
                report['warnings'].append(f"{missing_name} record(s) missing supplier name")

        report['quality_score'] = DataQualityAnalyzer._compute_quality_score(report)

        return report

    @staticmethod
    def _compute_quality_score(report: Dict[str, Any]) -> float:
        """
        Simple 0-100 heuristic quality score. Not a certification of correctness,
        just a fast signal for how much cleanup a file needs before trusting
        the reconciliation output.
        """
        total = max(report.get('total_records', 1), 1)
        penalty = 0.0
        penalty += min(report.get('gstin_checksum_failure_count', 0) / total, 0.3) * 100 * 0.35
        penalty += min(report.get('duplicate_invoice_count', 0) / total, 0.3) * 100 * 0.25
        penalty += min(report.get('unparseable_dates', 0) / total, 0.3) * 100 * 0.15
        penalty += min(report.get('negative_without_cn_flag', 0) / total, 0.3) * 100 * 0.15
        penalty += min(report.get('missing_supplier_name', 0) / total, 0.3) * 100 * 0.10
        score = max(0.0, 100.0 - penalty)
        return round(score, 1)

    @staticmethod
    def format_report_text(report: Dict[str, Any]) -> str:
        """Render a quality report as readable plain text (for logs / CLI)"""
        lines = [
            f"=== Data Quality Report: {report['source']} ===",
            f"Total records: {report['total_records']}",
            f"Quality score: {report.get('quality_score', 'N/A')}/100",
            ""
        ]
        if report['issues']:
            lines.append("ISSUES:")
            for issue in report['issues']:
                lines.append(f"  ❌ {issue}")
        if report['warnings']:
            lines.append("WARNINGS:")
            for warning in report['warnings']:
                lines.append(f"  ⚠️  {warning}")
        if not report['issues'] and not report['warnings']:
            lines.append("✅ No data quality issues detected.")
        return "\n".join(lines)


class ITCEligibilityEngine:
    """
    Rule-based Input Tax Credit (ITC) eligibility screening.
    This is a heuristic/informational aid only — NOT a substitute for
    professional judgement under Sections 16/17 of the CGST Act. It flags
    records for manual review; it does not make final eligibility determinations.
    """

    # Illustrative list of common blocked-credit keywords per Sec 17(5) —
    # matched against supplier name / description text if present.
    BLOCKED_CREDIT_HINTS = [
        'MOTOR VEHICLE', 'RENT A CAB', 'RENT-A-CAB', 'HEALTH INSURANCE',
        'LIFE INSURANCE', 'CLUB MEMBERSHIP', 'EMPLOYEE TRAVEL', 'FOOD AND BEVERAGE',
        'OUTDOOR CATERING', 'BEAUTY TREATMENT', 'COSMETIC SURGERY', 'PERSONAL CONSUMPTION'
    ]

    @staticmethod
    def evaluate_row(row: pd.Series, filing_date: Optional[datetime] = None,
                      annual_return_due_date: Optional[datetime] = None) -> Dict[str, Any]:
        """
        Evaluate a single reconciled row for ITC eligibility hints.

        filing_date: the date ITC is proposed to be claimed (defaults to today)
        annual_return_due_date: Sec 16(4) cutoff for the relevant FY (30th Nov
            of the following FY, or GSTR-9 due date — pass explicitly per FY)
        """
        result = {
            'eligible_hint': 'Review Required',
            'reasons': [],
        }

        doc_date = row.get('DOC_DATE_PARSED_2B', row.get('DOC_DATE_PARSED', None))
        supplier_name = str(row.get('SUPPLIER NAME_2B', row.get('SUPPLIER NAME', ''))).upper()
        match_status = row.get('MATCH_STATUS', '')

        # Rule 1: Must appear in GSTR-2B to claim (post-2021 rule, Sec 16(2)(aa))
        if match_status == 'Missing in 2B':
            result['eligible_hint'] = 'Not Eligible'
            result['reasons'].append(
                'Invoice not reflected in GSTR-2B — ITC cannot be claimed until supplier reports it'
            )
            return result

        # Rule 2: Section 16(4) time-limit check
        if annual_return_due_date and doc_date is not None and pd.notna(doc_date):
            try:
                if isinstance(doc_date, str):
                    doc_date = parse_date(doc_date)
                if doc_date and doc_date > annual_return_due_date:
                    result['reasons'].append('Document date is after the Sec 16(4) cutoff — check applicability')
            except Exception:
                pass

        # Rule 3: Blocked credit keyword screen (heuristic, needs manual review)
        for hint in ITCEligibilityEngine.BLOCKED_CREDIT_HINTS:
            if hint in supplier_name:
                result['reasons'].append(
                    f'Supplier/description text matches a potentially blocked-credit category ("{hint}") '
                    f'under Sec 17(5) — verify eligibility manually'
                )
                break

        # Rule 4: Low-confidence match — don't auto-claim
        confidence = safe_float_convert(row.get('CONFIDENCE', 0))
        if match_status in ('Suggested', 'Partial') and confidence < 0.7:
            result['reasons'].append(
                f'Match confidence is low ({confidence:.0%}) — verify manually before claiming ITC'
            )

        if not result['reasons'] and match_status == 'Exact':
            result['eligible_hint'] = 'Likely Eligible'
        elif not result['reasons']:
            result['eligible_hint'] = 'Likely Eligible (Review Recommended)'
        else:
            result['eligible_hint'] = 'Review Required'

        return result

    @staticmethod
    def evaluate_dataframe(final_df: pd.DataFrame,
                           annual_return_due_date: Optional[datetime] = None) -> pd.DataFrame:
        """Add ITC eligibility hint columns to a reconciled dataframe"""
        df = final_df.copy()
        if df.empty:
            df['ITC_ELIGIBILITY_HINT'] = pd.Series(dtype=str)
            df['ITC_REVIEW_REASONS'] = pd.Series(dtype=str)
            return df

        hints = []
        reasons_list = []
        for _, row in df.iterrows():
            res = ITCEligibilityEngine.evaluate_row(row, annual_return_due_date=annual_return_due_date)
            hints.append(res['eligible_hint'])
            reasons_list.append('; '.join(res['reasons']) if res['reasons'] else '')

        df['ITC_ELIGIBILITY_HINT'] = hints
        df['ITC_REVIEW_REASONS'] = reasons_list
        return df


class DataProcessor:
    """Advanced data processing utilities with credit note support"""
    
    @staticmethod
    def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
        """Standardize column names. Handles the edge case of a completely
        empty DataFrame (which pandas gives an integer RangeIndex for columns,
        not a string index)."""
        df = df.copy()
        if len(df.columns) == 0:
            return df
        df.columns = [str(c).upper().strip() for c in df.columns]
        return df
    
    @staticmethod
    def clean_data(df: pd.DataFrame) -> pd.DataFrame:
        """Clean and prepare data"""
        df = df.copy()
        df = df.drop_duplicates()
        df = df.dropna(how='all')
        
        string_columns = df.select_dtypes(include=['object', 'string']).columns
        for col in string_columns:
            df[col] = df[col].apply(safe_string_convert)
            df[col] = df[col].replace('nan', '')
            df[col] = df[col].replace('None', '')
        
        return df
    
    @staticmethod
    def normalize_column_mappings(df_2b: pd.DataFrame, df_pr: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """Normalize column mappings between 2B and PR"""
        df_2b = df_2b.copy()
        df_pr = df_pr.copy()
        
        column_mappings = {
            'SUPPLIER GSTIN': ['GSTIN', 'SUPPLIER_GSTIN', 'GSTIN OF SUPPLIER', 'REGISTRATION NUMBER'],
            'DOCUMENT NUMBER': ['INVOICE NUMBER', 'INVOICE NO', 'DOC NO', 'BILL NUMBER', 'DOCUMENT NO'],
            'DOCUMENT DATE': ['INVOICE DATE', 'BILL DATE', 'DATE', 'DOC DATE', 'INVOICE DT'],
            'TAXABLE VALUE': ['TAXABLE AMOUNT', 'TAXABLE VALUE', 'AMOUNT', 'BASE AMOUNT', 'VALUE'],
            'SUPPLIER NAME': ['NAME', 'SUPPLIER', 'VENDOR NAME', 'PARTY NAME', 'SELLER NAME'],
            'IGST': ['IGST AMOUNT', 'IGST TAX', 'INTEGRATED TAX'],
            'CGST': ['CGST AMOUNT', 'CGST TAX', 'CENTRAL TAX'],
            'SGST': ['SGST AMOUNT', 'SGST TAX', 'STATE TAX'],
            'TOTAL TAX': ['TOTAL TAX', 'TAX AMOUNT', 'GST TOTAL'],
            'TOTAL VALUE': ['TOTAL', 'TOTAL AMOUNT', 'INVOICE VALUE', 'BILL VALUE'],
            'DOC TYPE': ['DOC_TYPE', 'DOCUMENT TYPE', 'TYPE', 'DOCUMENT TYPE'],
            'REFERENCE DOCUMENT': ['REF_DOC', 'REFERENCE NO', 'REF DOC', 'REFERENCE INVOICE']
        }
        
        for standard_col, possible_cols in column_mappings.items():
            for df, name in [(df_2b, '2B'), (df_pr, 'PR')]:
                if standard_col not in df.columns:
                    for possible in possible_cols:
                        if possible in df.columns:
                            df.rename(columns={possible: standard_col}, inplace=True)
                            break
        
        return df_2b, df_pr
    
    @staticmethod
    def validate_required_columns(df: pd.DataFrame, required_cols: List[str]) -> bool:
        """Validate required columns exist"""
        for col in required_cols:
            if col not in df.columns:
                return False
        return True
    
    @staticmethod
    def get_missing_columns(df: pd.DataFrame, required_cols: List[str]) -> List[str]:
        """Get list of missing columns"""
        return [col for col in required_cols if col not in df.columns]
    
    @staticmethod
    def optimize_dtypes(df: pd.DataFrame) -> pd.DataFrame:
        """Optimize data types for memory efficiency"""
        df = df.copy()
        
        for col in df.columns:
            if df[col].dtype == 'object':
                if len(df[col].unique()) / len(df[col]) < 0.5:
                    df[col] = df[col].astype('category')
            elif df[col].dtype == 'int64':
                df[col] = pd.to_numeric(df[col], downcast='integer')
            elif df[col].dtype == 'float64':
                df[col] = pd.to_numeric(df[col], downcast='float')
        
        return df
    
    @staticmethod
    def detect_credit_notes(df: pd.DataFrame) -> pd.DataFrame:
        """Detect and mark credit notes in the data"""
        df = df.copy()
        
        df['IS_CREDIT_NOTE'] = False
        df['IS_DEBIT_NOTE'] = False
        df['NEGATIVE_AMOUNT'] = False
        
        if 'DOC TYPE' in df.columns:
            df['IS_CREDIT_NOTE'] = df['DOC TYPE'].str.upper().str.contains('CREDIT', na=False)
            df['IS_DEBIT_NOTE'] = df['DOC TYPE'].str.upper().str.contains('DEBIT', na=False)
        
        if 'TAXABLE VALUE' in df.columns:
            df['NEGATIVE_AMOUNT'] = df['TAXABLE VALUE'] < 0
            if 'DOC TYPE' in df.columns:
                mask = (df['NEGATIVE_AMOUNT']) & (~df['IS_CREDIT_NOTE'])
                df.loc[mask, 'IS_CREDIT_NOTE'] = True
                df.loc[mask, 'DOC TYPE'] = 'CREDIT NOTE'
        
        return df

# ============================================================================
# EXCEL EXPORT ENGINE - ENHANCED SIDE-BY-SIDE COMPARISON
# ============================================================================

class ExcelExportEngine:
    """Advanced Excel Export with Side-by-Side Comparison and Formulas"""
    
    # Color palette for professional formatting
    COLORS = {
        'header_bg': '#1a1a2e',
        'header_font': '#ffffff',
        'matched_bg': '#e8f5e9',
        'matched_font': '#2e7d32',
        'unmatched_bg': '#fff3e0',
        'unmatched_font': '#e65100',
        'credit_note_bg': '#fce4ec',
        'credit_note_font': '#c62828',
        'difference_bg': '#fff9c4',
        'difference_font': '#f57f17',
        'exact_match_bg': '#c8e6c9',
        'partial_match_bg': '#fff9c4',
        'summary_header_bg': '#1b5e20',
        'summary_row_bg': '#e8f5e9',
        'border_color': '#bdbdbd',
        'accent_blue': '#1565c0',
        'accent_purple': '#6a1b9a',
        'white': '#ffffff',
        'black': '#000000'
    }
    
    @staticmethod
    def create_comparison_export(final_df: pd.DataFrame, stats: Dict) -> BytesIO:
        """
        Create a comprehensive side-by-side comparison Excel report
        
        Sheets:
        1. Side-by-Side Comparison
        2. Matched Records  
        3. Missing in 2B
        4. Missing in PR
        5. Credit Notes Summary
        6. Dashboard Summary
        """
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Define formats
            formats = ExcelExportEngine._create_formats(workbook)
            
            # Create all sheets
            ExcelExportEngine._create_comparison_sheet(writer, workbook, final_df, formats, stats)
            ExcelExportEngine._create_matched_sheet(writer, workbook, final_df, formats)
            ExcelExportEngine._create_missing_sheets(writer, workbook, final_df, formats)
            ExcelExportEngine._create_credit_notes_sheet(writer, workbook, final_df, formats)
            ExcelExportEngine._create_dashboard_sheet(writer, workbook, final_df, formats, stats)
        
        output.seek(0)
        return output
    
    @staticmethod
    def _create_formats(workbook) -> Dict:
        """Create all Excel formats with professional styling"""
        formats = {}
        
        # Title format
        formats['title'] = workbook.add_format({
            'bold': True,
            'font_size': 16,
            'font_color': '#1a1a2e',
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Calibri',
            'bottom': 2,
            'bottom_color': '#1a1a2e'
        })
        
        # Subtitle format
        formats['subtitle'] = workbook.add_format({
            'bold': True,
            'font_size': 12,
            'font_color': '#455a64',
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Calibri',
            'bottom': 1,
            'bottom_color': '#90a4ae'
        })
        
        # Section headers
        formats['section_header'] = workbook.add_format({
            'bold': True,
            'font_size': 11,
            'font_color': '#ffffff',
            'bg_color': '#1a1a2e',
            'border': 1,
            'border_color': '#1a1a2e',
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Calibri',
            'num_format': '@'
        })
        
        # 2B column header (Blue)
        formats['header_2b'] = workbook.add_format({
            'bold': True,
            'font_size': 10,
            'font_color': '#ffffff',
            'bg_color': '#1565c0',
            'border': 1,
            'border_color': '#0d47a1',
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Calibri'
        })
        
        # PR column header (Purple)
        formats['header_pr'] = workbook.add_format({
            'bold': True,
            'font_size': 10,
            'font_color': '#ffffff',
            'bg_color': '#6a1b9a',
            'border': 1,
            'border_color': '#4a148c',
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Calibri'
        })
        
        # Difference column header (Red)
        formats['header_diff'] = workbook.add_format({
            'bold': True,
            'font_size': 10,
            'font_color': '#ffffff',
            'bg_color': '#c62828',
            'border': 1,
            'border_color': '#b71c1c',
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Calibri'
        })
        
        # Status column header (Green)
        formats['header_status'] = workbook.add_format({
            'bold': True,
            'font_size': 10,
            'font_color': '#ffffff',
            'bg_color': '#2e7d32',
            'border': 1,
            'border_color': '#1b5e20',
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Calibri'
        })
        
        # Matched data format
        formats['matched_cell'] = workbook.add_format({
            'font_size': 10,
            'border': 1,
            'border_color': '#bdbdbd',
            'align': 'right',
            'font_name': 'Calibri',
            'num_format': '#,##0.00'
        })
        
        # Matched data format for 2B columns
        formats['matched_2b_cell'] = workbook.add_format({
            'font_size': 10,
            'bg_color': '#e3f2fd',
            'border': 1,
            'border_color': '#bdbdbd',
            'align': 'right',
            'font_name': 'Calibri',
            'num_format': '#,##0.00'
        })
        
        # Matched data format for PR columns
        formats['matched_pr_cell'] = workbook.add_format({
            'font_size': 10,
            'bg_color': '#f3e5f5',
            'border': 1,
            'border_color': '#bdbdbd',
            'align': 'right',
            'font_name': 'Calibri',
            'num_format': '#,##0.00'
        })
        
        # Exact match format
        formats['exact_match'] = workbook.add_format({
            'font_size': 10,
            'bg_color': '#c8e6c9',
            'font_color': '#1b5e20',
            'border': 1,
            'border_color': '#a5d6a7',
            'align': 'right',
            'font_name': 'Calibri',
            'bold': True,
            'num_format': '#,##0.00'
        })
        
        # Credit note format
        formats['credit_note'] = workbook.add_format({
            'font_size': 10,
            'bg_color': '#fce4ec',
            'font_color': '#c62828',
            'border': 1,
            'border_color': '#ef9a9a',
            'align': 'right',
            'font_name': 'Calibri',
            'bold': True,
            'num_format': '#,##0.00'
        })
        
        # Difference format (positive)
        formats['diff_positive'] = workbook.add_format({
            'font_size': 10,
            'bg_color': '#fff9c4',
            'font_color': '#f57f17',
            'border': 1,
            'border_color': '#fff176',
            'align': 'right',
            'font_name': 'Calibri',
            'bold': True,
            'num_format': '#,##0.00'
        })
        
        # Difference format (negative)
        formats['diff_negative'] = workbook.add_format({
            'font_size': 10,
            'bg_color': '#ffebee',
            'font_color': '#b71c1c',
            'border': 1,
            'border_color': '#ef9a9a',
            'align': 'right',
            'font_name': 'Calibri',
            'bold': True,
            'num_format': '#,##0.00'
        })
        
        # Difference format (zero)
        formats['diff_zero'] = workbook.add_format({
            'font_size': 10,
            'bg_color': '#e8f5e9',
            'font_color': '#2e7d32',
            'border': 1,
            'border_color': '#a5d6a7',
            'align': 'right',
            'font_name': 'Calibri',
            'bold': True,
            'num_format': '#,##0.00'
        })
        
        # Text format
        formats['text_cell'] = workbook.add_format({
            'font_size': 10,
            'border': 1,
            'border_color': '#bdbdbd',
            'align': 'left',
            'font_name': 'Calibri',
            'text_wrap': True
        })
        
        # Text format for 2B
        formats['text_2b_cell'] = workbook.add_format({
            'font_size': 10,
            'bg_color': '#e3f2fd',
            'border': 1,
            'border_color': '#bdbdbd',
            'align': 'left',
            'font_name': 'Calibri',
            'text_wrap': True
        })
        
        # Text format for PR
        formats['text_pr_cell'] = workbook.add_format({
            'font_size': 10,
            'bg_color': '#f3e5f5',
            'border': 1,
            'border_color': '#bdbdbd',
            'align': 'left',
            'font_name': 'Calibri',
            'text_wrap': True
        })
        
        # Date format
        formats['date_cell'] = workbook.add_format({
            'font_size': 10,
            'border': 1,
            'border_color': '#bdbdbd',
            'align': 'center',
            'font_name': 'Calibri',
            'num_format': 'dd-mm-yyyy'
        })
        
        # Status formats
        formats['status_exact'] = workbook.add_format({
            'font_size': 10,
            'bg_color': '#c8e6c9',
            'font_color': '#1b5e20',
            'border': 1,
            'border_color': '#a5d6a7',
            'align': 'center',
            'font_name': 'Calibri',
            'bold': True
        })
        
        formats['status_suggested'] = workbook.add_format({
            'font_size': 10,
            'bg_color': '#fff9c4',
            'font_color': '#f57f17',
            'border': 1,
            'border_color': '#fff176',
            'align': 'center',
            'font_name': 'Calibri',
            'bold': True
        })
        
        formats['status_missing'] = workbook.add_format({
            'font_size': 10,
            'bg_color': '#ffebee',
            'font_color': '#b71c1c',
            'border': 1,
            'border_color': '#ef9a9a',
            'align': 'center',
            'font_name': 'Calibri',
            'bold': True
        })
        
        # Summary formats
        formats['summary_header'] = workbook.add_format({
            'bold': True,
            'font_size': 11,
            'font_color': '#ffffff',
            'bg_color': '#1b5e20',
            'border': 1,
            'border_color': '#0d47a1',
            'align': 'left',
            'valign': 'vcenter',
            'font_name': 'Calibri'
        })
        
        formats['summary_value'] = workbook.add_format({
            'font_size': 11,
            'border': 1,
            'border_color': '#bdbdbd',
            'align': 'right',
            'font_name': 'Calibri',
            'num_format': '#,##0.00'
        })
        
        formats['summary_label'] = workbook.add_format({
            'font_size': 11,
            'border': 1,
            'border_color': '#bdbdbd',
            'align': 'left',
            'font_name': 'Calibri',
            'bold': True
        })
        
        # Sub-total format
        formats['subtotal'] = workbook.add_format({
            'bold': True,
            'font_size': 10,
            'font_color': '#1a1a2e',
            'bg_color': '#e0e0e0',
            'border': 2,
            'border_color': '#757575',
            'align': 'right',
            'font_name': 'Calibri',
            'num_format': '#,##0.00'
        })
        
        # Grand total format
        formats['grand_total'] = workbook.add_format({
            'bold': True,
            'font_size': 11,
            'font_color': '#ffffff',
            'bg_color': '#1a1a2e',
            'border': 2,
            'border_color': '#000000',
            'align': 'right',
            'font_name': 'Calibri',
            'num_format': '#,##0.00'
        })
        
        return formats

    @staticmethod
    def _column_letter_to_index(col_letter: str) -> int:
        """Convert column letter (A, B, ..., Z, AA, AB, ...) to 0-based index"""
        col_letter = col_letter.upper()
        index = 0
        for char in col_letter:
            index = index * 26 + (ord(char) - ord('A') + 1)
        return index - 1

    @staticmethod
    def _create_comparison_sheet(writer, workbook, final_df, formats, stats):
        """Create the main side-by-side comparison sheet"""
        worksheet = workbook.add_worksheet('Side-by-Side Comparison')
        
        # Set column widths using direct column indices
        col_widths = {
            0: 5,    # S.No
            1: 18,   # Supplier GSTIN (2B)
            2: 18,   # Supplier GSTIN (PR)
            3: 25,   # Supplier Name (2B)
            4: 25,   # Supplier Name (PR)
            5: 18,   # Document Number (2B)
            6: 18,   # Document Number (PR)
            7: 14,   # Document Date (2B)
            8: 14,   # Document Date (PR)
            9: 12,   # Doc Type (2B)
            10: 12,  # Doc Type (PR)
            11: 16,  # Taxable Value (2B)
            12: 16,  # Taxable Value (PR)
            13: 16,  # Taxable Difference
            14: 14,  # IGST (2B)
            15: 14,  # IGST (PR)
            16: 14,  # IGST Difference
            17: 14,  # CGST (2B)
            18: 14,  # CGST (PR)
            19: 14,  # CGST Difference
            20: 14,  # SGST (2B)
            21: 14,  # SGST (PR)
            22: 14,  # SGST Difference
            23: 16,  # Total Tax (2B)
            24: 16,  # Total Tax (PR)
            25: 16,  # Total Tax Difference
            26: 16,  # Total Value (2B)
            27: 16,  # Total Value (PR)
            28: 16,  # Total Value Difference
            29: 14,  # Match Status
            30: 12,  # Confidence
            31: 12,  # Risk Level
            32: 18,  # Reference Document
        }
        
        for col_idx, width in col_widths.items():
            worksheet.set_column(col_idx, col_idx, width)
        
        # Freeze panes (keep title + header rows visible while scrolling)
        worksheet.freeze_panes(4, 1)
        worksheet.set_tab_color('#1565c0')
        worksheet.hide_gridlines(2)
        worksheet.set_zoom(100)
        
        # Title
        worksheet.merge_range(0, 0, 0, 32, 'GST RECONCILIATION - SIDE-BY-SIDE COMPARISON REPORT', formats['title'])
        worksheet.merge_range(1, 0, 1, 32, f'Generated: {datetime.now().strftime("%d-%m-%Y %H:%M:%S")} | Version: {VERSION}', formats['subtitle'])
        worksheet.merge_range(2, 0, 2, 32, f'GSTR-2B vs Purchase Register | Total Records: {stats["processed_records"]} | Match Rate: {stats["match_rate"]:.1f}%', formats['subtitle'])
        
        # Headers - Row 3 (0-indexed)
        headers = [
            (0, 'S.No', 'section_header'),
            # 2B columns
            (1, 'GSTIN (2B)', 'header_2b'),
            (2, 'GSTIN (Books)', 'header_pr'),
            (3, 'Supplier Name (2B)', 'header_2b'),
            (4, 'Supplier Name (Books)', 'header_pr'),
            (5, 'Invoice No (2B)', 'header_2b'),
            (6, 'Invoice No (Books)', 'header_pr'),
            (7, 'Date (2B)', 'header_2b'),
            (8, 'Date (Books)', 'header_pr'),
            (9, 'Doc Type (2B)', 'header_2b'),
            (10, 'Doc Type (Books)', 'header_pr'),
            # Difference columns
            (11, 'Taxable Value (2B)', 'header_2b'),
            (12, 'Taxable Value (Books)', 'header_pr'),
            (13, 'Taxable Diff', 'header_diff'),
            (14, 'IGST (2B)', 'header_2b'),
            (15, 'IGST (Books)', 'header_pr'),
            (16, 'IGST Diff', 'header_diff'),
            (17, 'CGST (2B)', 'header_2b'),
            (18, 'CGST (Books)', 'header_pr'),
            (19, 'CGST Diff', 'header_diff'),
            (20, 'SGST (2B)', 'header_2b'),
            (21, 'SGST (Books)', 'header_pr'),
            (22, 'SGST Diff', 'header_diff'),
            (23, 'Total Tax (2B)', 'header_2b'),
            (24, 'Total Tax (Books)', 'header_pr'),
            (25, 'Tax Diff', 'header_diff'),
            (26, 'Total Value (2B)', 'header_2b'),
            (27, 'Total Value (Books)', 'header_pr'),
            (28, 'Value Diff', 'header_diff'),
            # Status columns
            (29, 'Match Status', 'header_status'),
            (30, 'Confidence', 'header_status'),
            (31, 'Risk Level', 'header_status'),
            (32, 'Ref Document', 'header_status'),
        ]
        
        for col_idx, text, fmt_key in headers:
            worksheet.write(3, col_idx, text, formats[fmt_key])
        
        # Row height for header
        worksheet.set_row(3, 30)
        
        # Write data rows
        row = 4  # 0-indexed row 4 (Excel row 5)
        for idx, data_row in final_df.iterrows():
            # Determine if this is a matched record
            status = safe_string_convert(data_row.get('MATCH_STATUS', 'Missing'))
            is_matched = status in ['Exact', 'Suggested', 'Partial']
            is_credit = data_row.get('IS_CREDIT', False)
            
            # S.No
            worksheet.write(row, 0, idx + 1, formats['text_cell'])
            
            # Get values with defaults - try suffixed columns first, then non-suffixed
            gstin_2b = safe_string_convert(data_row.get('SUPPLIER GSTIN_2B', data_row.get('SUPPLIER GSTIN', '')))
            gstin_pr = safe_string_convert(data_row.get('SUPPLIER GSTIN_PR', data_row.get('SUPPLIER GSTIN', '')))
            name_2b = safe_string_convert(data_row.get('SUPPLIER NAME_2B', data_row.get('SUPPLIER NAME', '')))
            name_pr = safe_string_convert(data_row.get('SUPPLIER NAME_PR', data_row.get('SUPPLIER NAME', '')))
            doc_2b = safe_string_convert(data_row.get('DOCUMENT NUMBER_2B', data_row.get('DOCUMENT NUMBER', '')))
            doc_pr = safe_string_convert(data_row.get('DOCUMENT NUMBER_PR', data_row.get('DOCUMENT NUMBER', '')))
            doc_type_2b = safe_string_convert(data_row.get('DOC TYPE_2B', data_row.get('DOC TYPE', data_row.get('DOC_TYPE_NORM', ''))))
            doc_type_pr = safe_string_convert(data_row.get('DOC TYPE_PR', data_row.get('DOC TYPE', data_row.get('DOC_TYPE_NORM', ''))))
            
            # Taxable values - ensure numeric
            taxable_2b = safe_float_convert(data_row.get('TAXABLE VALUE_2B', data_row.get('TAXABLE VALUE', 0)))
            taxable_pr = safe_float_convert(data_row.get('TAXABLE VALUE_PR', data_row.get('TAXABLE VALUE', 0)))
            
            # Tax values - ensure numeric
            igst_2b = safe_float_convert(data_row.get('IGST_2B', data_row.get('IGST', 0)))
            igst_pr = safe_float_convert(data_row.get('IGST_PR', data_row.get('IGST', 0)))
            cgst_2b = safe_float_convert(data_row.get('CGST_2B', data_row.get('CGST', 0)))
            cgst_pr = safe_float_convert(data_row.get('CGST_PR', data_row.get('CGST', 0)))
            sgst_2b = safe_float_convert(data_row.get('SGST_2B', data_row.get('SGST', 0)))
            sgst_pr = safe_float_convert(data_row.get('SGST_PR', data_row.get('SGST', 0)))
            
            total_tax_2b = igst_2b + cgst_2b + sgst_2b
            total_tax_pr = igst_pr + cgst_pr + sgst_pr
            
            total_value_2b = taxable_2b + total_tax_2b
            total_value_pr = taxable_pr + total_tax_pr
            
            # Select formats based on match status
            if is_matched:
                text_2b_fmt = formats['text_2b_cell']
                text_pr_fmt = formats['text_pr_cell']
                num_2b_fmt = formats['matched_2b_cell']
                num_pr_fmt = formats['matched_pr_cell']
            else:
                text_2b_fmt = formats['text_cell']
                text_pr_fmt = formats['text_cell']
                num_2b_fmt = formats['matched_cell']
                num_pr_fmt = formats['matched_cell']
            
            # Override with credit note formatting
            if is_credit:
                num_2b_fmt = formats['credit_note']
                num_pr_fmt = formats['credit_note']
            
            # Write text columns
            worksheet.write(row, 1, gstin_2b, text_2b_fmt)
            worksheet.write(row, 2, gstin_pr, text_pr_fmt)
            worksheet.write(row, 3, name_2b, text_2b_fmt)
            worksheet.write(row, 4, name_pr, text_pr_fmt)
            worksheet.write(row, 5, doc_2b, text_2b_fmt)
            worksheet.write(row, 6, doc_pr, text_pr_fmt)
            worksheet.write(row, 9, doc_type_2b, text_2b_fmt)
            worksheet.write(row, 10, doc_type_pr, text_pr_fmt)
            
            # Write dates
            date_2b = data_row.get('DOCUMENT DATE_2B', data_row.get('DOCUMENT DATE', data_row.get('DOC_DATE_PARSED_2B', None)))
            date_pr = data_row.get('DOCUMENT DATE_PR', data_row.get('DOCUMENT DATE', data_row.get('DOC_DATE_PARSED_PR', None)))
            
            if date_2b:
                try:
                    if isinstance(date_2b, str):
                        date_2b = pd.to_datetime(date_2b)
                    worksheet.write_datetime(row, 7, date_2b, formats['date_cell'])
                except:
                    worksheet.write(row, 7, str(date_2b), formats['text_cell'])
            else:
                worksheet.write(row, 7, '', formats['text_cell'])
            
            if date_pr:
                try:
                    if isinstance(date_pr, str):
                        date_pr = pd.to_datetime(date_pr)
                    worksheet.write_datetime(row, 8, date_pr, formats['date_cell'])
                except:
                    worksheet.write(row, 8, str(date_pr), formats['text_cell'])
            else:
                worksheet.write(row, 8, '', formats['text_cell'])
            
            # Excel row numbers (1-indexed)
            excel_row = row + 1
            
            # Write numeric columns with formulas
            # Taxable Value
            worksheet.write(row, 11, taxable_2b, num_2b_fmt)
            worksheet.write(row, 12, taxable_pr, num_pr_fmt)
            # Formula for taxable difference - using direct column indices
            diff_fmt = formats['diff_zero'] if abs(taxable_2b - taxable_pr) < 0.01 else formats['diff_positive'] if taxable_2b >= taxable_pr else formats['diff_negative']
            worksheet.write_formula(row, 13, f'=L{excel_row}-M{excel_row}', diff_fmt)
            
            # IGST
            worksheet.write(row, 14, igst_2b, num_2b_fmt)
            worksheet.write(row, 15, igst_pr, num_pr_fmt)
            diff_fmt_igst = formats['diff_zero'] if abs(igst_2b - igst_pr) < 0.01 else formats['diff_positive'] if igst_2b >= igst_pr else formats['diff_negative']
            worksheet.write_formula(row, 16, f'=O{excel_row}-P{excel_row}', diff_fmt_igst)
            
            # CGST
            worksheet.write(row, 17, cgst_2b, num_2b_fmt)
            worksheet.write(row, 18, cgst_pr, num_pr_fmt)
            diff_fmt_cgst = formats['diff_zero'] if abs(cgst_2b - cgst_pr) < 0.01 else formats['diff_positive'] if cgst_2b >= cgst_pr else formats['diff_negative']
            worksheet.write_formula(row, 19, f'=R{excel_row}-S{excel_row}', diff_fmt_cgst)
            
            # SGST
            worksheet.write(row, 20, sgst_2b, num_2b_fmt)
            worksheet.write(row, 21, sgst_pr, num_pr_fmt)
            diff_fmt_sgst = formats['diff_zero'] if abs(sgst_2b - sgst_pr) < 0.01 else formats['diff_positive'] if sgst_2b >= sgst_pr else formats['diff_negative']
            worksheet.write_formula(row, 22, f'=U{excel_row}-V{excel_row}', diff_fmt_sgst)
            
            # Total Tax
            worksheet.write(row, 23, total_tax_2b, num_2b_fmt)
            worksheet.write(row, 24, total_tax_pr, num_pr_fmt)
            diff_fmt_tax = formats['diff_zero'] if abs(total_tax_2b - total_tax_pr) < 0.01 else formats['diff_positive'] if total_tax_2b >= total_tax_pr else formats['diff_negative']
            worksheet.write_formula(row, 25, f'=X{excel_row}-Y{excel_row}', diff_fmt_tax)
            
            # Total Value
            worksheet.write(row, 26, total_value_2b, num_2b_fmt)
            worksheet.write(row, 27, total_value_pr, num_pr_fmt)
            diff_fmt_value = formats['diff_zero'] if abs(total_value_2b - total_value_pr) < 0.01 else formats['diff_positive'] if total_value_2b >= total_value_pr else formats['diff_negative']
            worksheet.write_formula(row, 28, f'=AA{excel_row}-AB{excel_row}', diff_fmt_value)
            
            # Status
            status_fmt = formats['status_exact']
            if status == 'Exact':
                status_fmt = formats['status_exact']
            elif status in ['Suggested', 'Partial']:
                status_fmt = formats['status_suggested']
            elif 'Missing' in status:
                status_fmt = formats['status_missing']
            
            worksheet.write(row, 29, status, status_fmt)
            worksheet.write(row, 30, safe_float_convert(data_row.get('CONFIDENCE', 0)), formats['matched_cell'])
            worksheet.write(row, 31, safe_string_convert(data_row.get('RISK_LEVEL', '')), formats['text_cell'])
            worksheet.write(row, 32, safe_string_convert(data_row.get('REFERENCE DOCUMENT', data_row.get('REFERENCE DOCUMENT_2B', ''))), formats['text_cell'])
            
            row += 1
        
        # Autofilter across the full data range (Excel best-practice: filterable header)
        last_data_row_0idx = row - 1
        if last_data_row_0idx >= 4:
            worksheet.autofilter(3, 0, last_data_row_0idx, 32)
        
        # Conditional formatting: data bars on Confidence, color scale on Difference columns,
        # icon set flags on Taxable Difference so mismatches jump out visually.
        if row > 4:
            first_r, last_r = 4, row - 1
            
            # Confidence column (AE) -> data bar
            worksheet.conditional_format(first_r, 30, last_r, 30, {
                'type': 'data_bar',
                'bar_color': '#1565c0',
                'bar_only': False
            })
            
            # Difference columns -> 3-color scale (red = big variance, white = mid, green = zero)
            for diff_col in [13, 16, 19, 22, 25, 28]:
                worksheet.conditional_format(first_r, diff_col, last_r, diff_col, {
                    'type': '3_color_scale',
                    'min_color': '#2e7d32',
                    'mid_color': '#ffffff',
                    'max_color': '#c62828'
                })
            
            # Taxable Difference -> icon set to flag material mismatches at a glance
            worksheet.conditional_format(first_r, 13, last_r, 13, {
                'type': 'icon_set',
                'icon_style': '3_traffic_lights',
                'reverse_icons': True,
                'icons_only': True
            })
            
            # Match Status column -> highlight "Missing" rows with a solid red fill
            worksheet.conditional_format(first_r, 29, last_r, 29, {
                'type': 'text',
                'criteria': 'containing',
                'value': 'Missing',
                'format': formats['status_missing']
            })
        
        # Add summary section with formulas
        last_data_row = row
        
        summary_start_row = row + 2
        
        worksheet.merge_range(summary_start_row, 0, summary_start_row, 32, 'SUMMARY & TOTALS', formats['title'])
        
        # Summary rows with formulas
        summary_items = [
            ('Total Records', f'=COUNTA(B5:B{last_data_row})'),
            ('Matched Records', f'=COUNTIF(AD5:AD{last_data_row},"Exact")+COUNTIF(AD5:AD{last_data_row},"Suggested")+COUNTIF(AD5:AD{last_data_row},"Partial")'),
            ('Missing in 2B', f'=COUNTIF(AD5:AD{last_data_row},"Missing in 2B")'),
            ('Missing in PR', f'=COUNTIF(AD5:AD{last_data_row},"Missing in PR")'),
            ('Total Taxable Value (2B)', f'=SUM(L5:L{last_data_row})'),
            ('Total Taxable Value (Books)', f'=SUM(M5:M{last_data_row})'),
            ('Taxable Value Difference', f'=SUM(N5:N{last_data_row})'),
            ('Total IGST (2B)', f'=SUM(O5:O{last_data_row})'),
            ('Total IGST (Books)', f'=SUM(P5:P{last_data_row})'),
            ('Total Tax Difference', f'=SUM(Z5:Z{last_data_row})'),
        ]
        
        for i, (label, formula) in enumerate(summary_items):
            r = summary_start_row + 2 + i
            worksheet.write(r, 0, label, formats['summary_label'])
            worksheet.write_formula(r, 1, formula, formats['summary_value'])
    
    @staticmethod
    def _create_matched_sheet(writer, workbook, final_df, formats):
        """Create matched records sheet"""
        matched_df = final_df[final_df['MATCH_STATUS'].isin(['Exact', 'Suggested', 'Partial'])].copy()
        
        worksheet = workbook.add_worksheet('Matched Records')
        
        if matched_df.empty:
            worksheet.write(0, 0, 'No matched records found', formats['title'])
            return
        
        # Columns to export - prefer _2B, _PR or fallback
        export_cols = [
            ('SUPPLIER GSTIN', 'text_cell'),
            ('SUPPLIER NAME', 'text_cell'),
            ('DOCUMENT NUMBER', 'text_cell'),
            ('DOCUMENT DATE', 'date_cell'),
            ('TAXABLE VALUE', 'matched_cell'),
            ('IGST', 'matched_cell'),
            ('CGST', 'matched_cell'),
            ('SGST', 'matched_cell'),
            ('TOTAL TAX', 'matched_cell'),
            ('TOTAL VALUE', 'matched_cell'),
            ('MATCH_STATUS', 'status_exact'),
            ('CONFIDENCE', 'matched_cell'),
        ]
        
        # Write headers
        for col_idx, (col_name, fmt_key) in enumerate(export_cols):
            worksheet.write(0, col_idx, col_name, formats['section_header'])
        
        worksheet.freeze_panes(1, 0)
        worksheet.set_tab_color('#2e7d32')
        worksheet.hide_gridlines(2)
        
        # Write data
        for row_idx, (_, data_row) in enumerate(matched_df.iterrows()):
            for col_idx, (col_name, fmt_key) in enumerate(export_cols):
                # Try to get the value from _2B first, then _PR, then without suffix
                val = None
                if f'{col_name}_2B' in data_row.index:
                    val = data_row[f'{col_name}_2B']
                elif f'{col_name}_PR' in data_row.index:
                    val = data_row[f'{col_name}_PR']
                elif col_name in data_row.index:
                    val = data_row[col_name]
                else:
                    val = ''
                if pd.isna(val):
                    val = ''
                
                # Ensure numeric columns are properly formatted
                if col_name in ['TAXABLE VALUE', 'IGST', 'CGST', 'SGST', 'TOTAL TAX', 'TOTAL VALUE', 'CONFIDENCE']:
                    val = safe_float_convert(val)
                
                worksheet.write(row_idx + 1, col_idx, val, formats[fmt_key])
        
        for col_idx, (col_name, _) in enumerate(export_cols):
            worksheet.set_column(col_idx, col_idx, 16 if col_name != 'SUPPLIER NAME' else 26)
        
        if len(matched_df) > 0:
            worksheet.autofilter(0, 0, len(matched_df), len(export_cols) - 1)
            worksheet.conditional_format(1, 11, len(matched_df), 11, {
                'type': 'data_bar',
                'bar_color': '#66bb6a'
            })
    
    @staticmethod
    def _create_missing_sheets(writer, workbook, final_df, formats):
        """Create missing in 2B and missing in PR sheets"""
        # Missing in 2B (present in PR, missing in 2B)
        missing_2b = final_df[final_df['MATCH_STATUS'] == 'Missing in 2B'].copy()
        
        if not missing_2b.empty:
            worksheet = workbook.add_worksheet('Missing in 2B')
            worksheet.write(0, 0, 'Records in Purchase Register but missing in GSTR-2B', formats['title'])
            
            # Get columns that exist with _PR or without suffix
            pr_cols = [col for col in missing_2b.columns if col.endswith('_PR') or (not col.endswith('_2B') and col in ['SUPPLIER GSTIN', 'SUPPLIER NAME', 'DOCUMENT NUMBER', 'TAXABLE VALUE', 'DOCUMENT DATE', 'IGST', 'CGST', 'SGST', 'TOTAL TAX', 'TOTAL VALUE'])]
            if not pr_cols:
                pr_cols = [col for col in missing_2b.columns if col.endswith('_PR') or col not in ['MATCH_STATUS', 'MATCH_REASON', 'MATCH_TIER', 'CONFIDENCE', 'RISK_LEVEL', 'IS_CREDIT', 'DOC_TYPE_DISPLAY']]
            
            for col_idx, col in enumerate(pr_cols[:15]):
                display_name = col.replace('_PR', ' (Books)')
                worksheet.write(0, col_idx, display_name, formats['header_pr'])
            
            worksheet.freeze_panes(1, 0)
            worksheet.set_tab_color('#8b5cf6')
            worksheet.hide_gridlines(2)
            if len(missing_2b) > 0:
                worksheet.autofilter(0, 0, len(missing_2b), min(len(pr_cols), 15) - 1)
            
            for row_idx, (_, data_row) in enumerate(missing_2b.iterrows()):
                for col_idx, col in enumerate(pr_cols[:15]):
                    val = data_row.get(col, '')
                    if pd.isna(val):
                        val = ''
                    
                    # Ensure numeric columns are properly formatted
                    if col in ['TAXABLE VALUE', 'IGST', 'CGST', 'SGST', 'TOTAL TAX', 'TOTAL VALUE']:
                        val = safe_float_convert(val)
                    
                    worksheet.write(row_idx + 1, col_idx, val, formats['text_pr_cell'])
        else:
            worksheet = workbook.add_worksheet('Missing in 2B')
            worksheet.write(0, 0, 'No records missing in 2B', formats['title'])
        
        # Missing in PR (present in 2B, missing in PR)
        missing_pr = final_df[final_df['MATCH_STATUS'] == 'Missing in PR'].copy()
        
        if not missing_pr.empty:
            worksheet = workbook.add_worksheet('Missing in PR')
            worksheet.write(0, 0, 'Records in GSTR-2B but missing in Purchase Register', formats['title'])
            
            cols_2b = [col for col in missing_pr.columns if col.endswith('_2B') or (not col.endswith('_PR') and col in ['SUPPLIER GSTIN', 'SUPPLIER NAME', 'DOCUMENT NUMBER', 'TAXABLE VALUE', 'DOCUMENT DATE', 'IGST', 'CGST', 'SGST', 'TOTAL TAX', 'TOTAL VALUE'])]
            if not cols_2b:
                cols_2b = [col for col in missing_pr.columns if col.endswith('_2B') or col not in ['MATCH_STATUS', 'MATCH_REASON', 'MATCH_TIER', 'CONFIDENCE', 'RISK_LEVEL', 'IS_CREDIT', 'DOC_TYPE_DISPLAY']]
            
            for col_idx, col in enumerate(cols_2b[:15]):
                display_name = col.replace('_2B', ' (2B)')
                worksheet.write(0, col_idx, display_name, formats['header_2b'])
            
            worksheet.freeze_panes(1, 0)
            worksheet.set_tab_color('#f59e0b')
            worksheet.hide_gridlines(2)
            if len(missing_pr) > 0:
                worksheet.autofilter(0, 0, len(missing_pr), min(len(cols_2b), 15) - 1)
            
            for row_idx, (_, data_row) in enumerate(missing_pr.iterrows()):
                for col_idx, col in enumerate(cols_2b[:15]):
                    val = data_row.get(col, '')
                    if pd.isna(val):
                        val = ''
                    
                    # Ensure numeric columns are properly formatted
                    if col in ['TAXABLE VALUE', 'IGST', 'CGST', 'SGST', 'TOTAL TAX', 'TOTAL VALUE']:
                        val = safe_float_convert(val)
                    
                    worksheet.write(row_idx + 1, col_idx, val, formats['text_2b_cell'])
        else:
            worksheet = workbook.add_worksheet('Missing in PR')
            worksheet.write(0, 0, 'No records missing in PR', formats['title'])
    
    @staticmethod
    def _create_credit_notes_sheet(writer, workbook, final_df, formats):
        """Create credit notes summary sheet"""
        # Find credit notes
        is_credit_col = None
        if 'IS_CREDIT' in final_df.columns:
            is_credit_col = 'IS_CREDIT'
        elif 'IS_CREDIT_NOTE' in final_df.columns:
            is_credit_col = 'IS_CREDIT_NOTE'
        
        if is_credit_col:
            credit_df = final_df[final_df[is_credit_col] == True].copy()
        elif 'DOC TYPE' in final_df.columns:
            credit_df = final_df[final_df['DOC TYPE'].str.upper().str.contains('CREDIT', na=False)].copy()
        else:
            credit_df = pd.DataFrame()
        
        worksheet = workbook.add_worksheet('Credit Notes Summary')
        worksheet.set_tab_color('#c62828')
        worksheet.hide_gridlines(2)
        
        if credit_df.empty:
            worksheet.write(0, 0, 'No credit notes found', formats['title'])
            return
        
        # Title
        worksheet.merge_range(0, 0, 0, 6, 'CREDIT NOTES RECONCILIATION SUMMARY', formats['title'])
        worksheet.merge_range(1, 0, 1, 6, f'Total Credit Notes: {len(credit_df)}', formats['subtitle'])
        
        # Headers
        credit_headers = ['S.No', 'Supplier GSTIN', 'Supplier Name', 'Credit Note No', 'Date', 'Taxable Value', 'Match Status']
        for col_idx, header in enumerate(credit_headers):
            worksheet.write(3, col_idx, header, formats['section_header'])
            worksheet.set_column(col_idx, col_idx, 18)
        
        worksheet.freeze_panes(4, 0)
        if len(credit_df) > 0:
            worksheet.autofilter(3, 0, 3 + len(credit_df), len(credit_headers) - 1)
        
        # Write credit notes
        for row_idx, (_, data_row) in enumerate(credit_df.iterrows()):
            worksheet.write(row_idx + 4, 0, row_idx + 1, formats['text_cell'])
            
            # Try to get values with suffixes
            gstin = safe_string_convert(data_row.get('SUPPLIER GSTIN_2B', data_row.get('SUPPLIER GSTIN_PR', data_row.get('SUPPLIER GSTIN', ''))))
            name = safe_string_convert(data_row.get('SUPPLIER NAME_2B', data_row.get('SUPPLIER NAME_PR', data_row.get('SUPPLIER NAME', ''))))
            doc_num = safe_string_convert(data_row.get('DOCUMENT NUMBER_2B', data_row.get('DOCUMENT NUMBER_PR', data_row.get('DOCUMENT NUMBER', ''))))
            taxable = safe_float_convert(data_row.get('TAXABLE VALUE_2B', data_row.get('TAXABLE VALUE_PR', data_row.get('TAXABLE VALUE', 0))))
            status = safe_string_convert(data_row.get('MATCH_STATUS', ''))
            
            worksheet.write(row_idx + 4, 1, gstin, formats['text_cell'])
            worksheet.write(row_idx + 4, 2, name, formats['text_cell'])
            worksheet.write(row_idx + 4, 3, doc_num, formats['text_cell'])
            
            # Date
            date_val = data_row.get('DOCUMENT DATE_2B', data_row.get('DOCUMENT DATE_PR', data_row.get('DOCUMENT DATE', '')))
            worksheet.write(row_idx + 4, 4, safe_string_convert(date_val), formats['text_cell'])
            
            worksheet.write(row_idx + 4, 5, taxable, formats['credit_note'])
            worksheet.write(row_idx + 4, 6, status, formats['status_exact'])
        
        # Add totals
        total_row = len(credit_df) + 5
        worksheet.write(total_row, 0, 'TOTAL', formats['summary_label'])
        worksheet.write_formula(total_row, 5, f'=SUM(F5:F{total_row-1})', formats['credit_note'])
    
    @staticmethod
    def _create_dashboard_sheet(writer, workbook, final_df, formats, stats):
        """Create a dashboard summary sheet with embedded native Excel charts"""
        worksheet = workbook.add_worksheet('Dashboard Summary')
        worksheet.set_tab_color('#1a1a2e')
        worksheet.hide_gridlines(2)
        
        # Title
        worksheet.merge_range(0, 0, 0, 5, 'GST RECONCILIATION DASHBOARD', formats['title'])
        worksheet.merge_range(1, 0, 1, 5, f'Report Generated: {datetime.now().strftime("%d-%m-%Y %H:%M:%S")}', formats['subtitle'])
        
        # Key Metrics
        worksheet.merge_range(3, 0, 3, 5, 'KEY METRICS', formats['section_header'])
        
        metrics = [
            ('Total Records Processed', stats.get('processed_records', 0)),
            ('Match Rate', f"{stats.get('match_rate', 0):.1f}%"),
            ('Exact Matches', stats.get('exact_matches', 0)),
            ('Suggested Matches', stats.get('suggested_matches', 0)),
            ('Partial Matches', stats.get('partial_matches', 0)),
            ('Missing in 2B', stats.get('missing_in_2b', 0)),
            ('Missing in PR', stats.get('missing_in_pr', 0)),
            ('Average Confidence', f"{stats.get('avg_confidence', 0):.1%}"),
            ('Processing Time', f"{stats.get('processing_time', 0):.2f}s"),
        ]
        
        for i, (label, value) in enumerate(metrics):
            row = 4 + i
            worksheet.write(row, 0, label, formats['summary_label'])
            worksheet.write(row, 1, value, formats['summary_value'])
            worksheet.merge_range(row, 2, row, 5, '', formats['summary_value'])
        
        # Financial Summary
        financial_start = 4 + len(metrics) + 2
        worksheet.merge_range(financial_start, 0, financial_start, 5, 'FINANCIAL SUMMARY', formats['section_header'])
        
        financial_metrics = [
            ('Total Taxable Value (2B)', stats.get('total_taxable_2b', 0)),
            ('Total Taxable Value (Books)', stats.get('total_taxable_pr', 0)),
            ('Taxable Difference', stats.get('taxable_difference', 0)),
        ]
        
        for i, (label, value) in enumerate(financial_metrics):
            row = financial_start + 1 + i
            worksheet.write(row, 0, label, formats['summary_label'])
            worksheet.write(row, 1, value, formats['summary_value'])
            worksheet.merge_range(row, 2, row, 5, '', formats['summary_value'])
        
        # Credit Notes Summary
        credit_start = financial_start + 1 + len(financial_metrics) + 2
        if 'credit_notes' in stats:
            cn = stats['credit_notes']
            worksheet.merge_range(credit_start, 0, credit_start, 5, 'CREDIT NOTES SUMMARY', formats['section_header'])
            
            credit_metrics = [
                ('Credit Notes in 2B', cn.get('total_2b', 0)),
                ('Credit Notes in Books', cn.get('total_pr', 0)),
                ('Matched Credit Notes', cn.get('matched', 0)),
                ('Unmatched Credit Notes (2B)', cn.get('unmatched_2b', 0)),
                ('Unmatched Credit Notes (Books)', cn.get('unmatched_pr', 0)),
                ('Credit Amount (2B)', cn.get('total_amount_2b', 0)),
                ('Credit Amount (Books)', cn.get('total_amount_pr', 0)),
                ('Credit Amount Difference', cn.get('difference', 0)),
            ]
            
            for i, (label, value) in enumerate(credit_metrics):
                row = credit_start + 1 + i
                worksheet.write(row, 0, label, formats['summary_label'])
                worksheet.write(row, 1, value, formats['summary_value'])
                worksheet.merge_range(row, 2, row, 5, '', formats['summary_value'])
        
        # Strategy Usage
        strategy_start = credit_start + 1 + len(credit_metrics) + 2 if 'credit_notes' in stats else credit_start + 2
        if 'strategy_breakdown' in stats:
            worksheet.merge_range(strategy_start, 0, strategy_start, 5, 'STRATEGY USAGE', formats['section_header'])
            
            strategy_items = list(stats['strategy_breakdown'].items())
            for i, (strategy, count) in enumerate(strategy_items[:15]):
                row = strategy_start + 1 + i
                worksheet.write(row, 0, strategy.replace('_', ' ').title(), formats['summary_label'])
                worksheet.write(row, 1, count, formats['summary_value'])
                worksheet.merge_range(row, 2, row, 5, '', formats['summary_value'])
        
        # ------------------------------------------------------------------
        # CHART DATA TABLES (placed off to the side, used as chart sources)
        # ------------------------------------------------------------------
        chart_col = 8  # column I
        
        # Status breakdown data -> Pie chart
        status_counts = final_df['MATCH_STATUS'].value_counts()
        worksheet.write(3, chart_col, 'Status', formats['section_header'])
        worksheet.write(3, chart_col + 1, 'Count', formats['section_header'])
        for i, (status_label, count) in enumerate(status_counts.items()):
            worksheet.write(4 + i, chart_col, str(status_label), formats['text_cell'])
            worksheet.write(4 + i, chart_col + 1, int(count), formats['summary_value'])
        status_rows = len(status_counts)
        
        status_pie = workbook.add_chart({'type': 'pie'})
        status_pie.add_series({
            'name': 'Match Status Breakdown',
            'categories': ['Dashboard Summary', 4, chart_col, 3 + status_rows, chart_col],
            'values': ['Dashboard Summary', 4, chart_col + 1, 3 + status_rows, chart_col + 1],
            'data_labels': {'percentage': True, 'category': True, 'position': 'outside_end'},
            'points': [
                {'fill': {'color': '#10b981'}}, {'fill': {'color': '#3b82f6'}},
                {'fill': {'color': '#8b5cf6'}}, {'fill': {'color': '#ef4444'}},
                {'fill': {'color': '#f59e0b'}}, {'fill': {'color': '#64748b'}},
            ][:max(status_rows, 1)]
        })
        status_pie.set_title({'name': 'Match Status Breakdown'})
        status_pie.set_size({'width': 460, 'height': 300})
        worksheet.insert_chart(3, chart_col + 3, status_pie)
        
        # Financial comparison data -> Column chart
        fin_row_start = 4 + status_rows + 2
        worksheet.write(fin_row_start, chart_col, 'Metric', formats['section_header'])
        worksheet.write(fin_row_start, chart_col + 1, 'Value', formats['section_header'])
        fin_labels = ['Taxable (2B)', 'Taxable (Books)', 'Difference']
        fin_values = [stats.get('total_taxable_2b', 0), stats.get('total_taxable_pr', 0), stats.get('taxable_difference', 0)]
        for i, (lbl, val) in enumerate(zip(fin_labels, fin_values)):
            worksheet.write(fin_row_start + 1 + i, chart_col, lbl, formats['text_cell'])
            worksheet.write(fin_row_start + 1 + i, chart_col + 1, val, formats['summary_value'])
        
        fin_chart = workbook.add_chart({'type': 'column'})
        fin_chart.add_series({
            'name': 'Taxable Value Comparison (₹)',
            'categories': ['Dashboard Summary', fin_row_start + 1, chart_col, fin_row_start + 3, chart_col],
            'values': ['Dashboard Summary', fin_row_start + 1, chart_col + 1, fin_row_start + 3, chart_col + 1],
            'fill': {'color': '#1565c0'},
            'data_labels': {'value': True, 'num_format': '#,##0'}
        })
        fin_chart.set_title({'name': 'Taxable Value: 2B vs Books'})
        fin_chart.set_legend({'none': True})
        fin_chart.set_size({'width': 460, 'height': 300})
        worksheet.insert_chart(3 + 17, chart_col + 3, fin_chart)
        
        # Strategy usage -> horizontal bar chart
        if 'strategy_breakdown' in stats and stats['strategy_breakdown']:
            strat_row_start = fin_row_start + 6
            worksheet.write(strat_row_start, chart_col, 'Strategy', formats['section_header'])
            worksheet.write(strat_row_start, chart_col + 1, 'Matches', formats['section_header'])
            strat_items = list(stats['strategy_breakdown'].items())[:12]
            for i, (strat, count) in enumerate(strat_items):
                worksheet.write(strat_row_start + 1 + i, chart_col, str(strat).replace('_', ' ').title(), formats['text_cell'])
                worksheet.write(strat_row_start + 1 + i, chart_col + 1, count, formats['summary_value'])
            
            strat_chart = workbook.add_chart({'type': 'bar'})
            strat_chart.add_series({
                'name': 'Matches by Strategy',
                'categories': ['Dashboard Summary', strat_row_start + 1, chart_col, strat_row_start + len(strat_items), chart_col],
                'values': ['Dashboard Summary', strat_row_start + 1, chart_col + 1, strat_row_start + len(strat_items), chart_col + 1],
                'fill': {'color': '#6a1b9a'},
                'data_labels': {'value': True}
            })
            strat_chart.set_title({'name': 'Matches by Strategy'})
            strat_chart.set_legend({'none': True})
            strat_chart.set_size({'width': 460, 'height': 320})
            worksheet.insert_chart(3 + 34, chart_col + 3, strat_chart)
        
        # Set column widths
        worksheet.set_column(0, 0, 35)
        worksheet.set_column(1, 1, 20)
        for col in range(2, 6):
            worksheet.set_column(col, col, 15)
        worksheet.set_column(chart_col, chart_col, 24)
        worksheet.set_column(chart_col + 1, chart_col + 1, 14)

# ============================================================================
# GST RECONCILIATION ENGINE - STRICTENED FOR NO FALSE POSITIVES
# ============================================================================

class GSTReconciliationEngine:
    """Enterprise-grade GST Reconciliation Engine - Strict matching to avoid false positives"""
    
    def __init__(self, config: ReconciliationConfig):
        self.config = config
        self.logger = LoggerSetup().get_logger()
        self.stats = defaultdict(int)
        self._lock = threading.Lock()
        self.metrics = ProcessingMetrics()
        
        self.logger.info(f"Initialized GSTReconciliationEngine v{VERSION}")
    
    def _safe_series_sum(self, series: pd.Series) -> float:
        """Safely sum a pandas Series, converting string values to float"""
        if series is None or len(series) == 0:
            return 0.0
        
        numeric_series = pd.to_numeric(series, errors='coerce')
        return numeric_series.sum()
    
    def reconcile(self, df_2b: pd.DataFrame, df_pr: pd.DataFrame) -> Tuple[pd.DataFrame, Dict]:
        """Main reconciliation orchestration with strict matching"""
        start_time = time.time()
        self.metrics.start_time = datetime.now()
        self.metrics.total_records_2b = len(df_2b)
        self.metrics.total_records_pr = len(df_pr)
        
        self.logger.info(f"Starting reconciliation with {len(df_2b)} 2B records and {len(df_pr)} PR records")
        
        try:
            df_2b_processed, df_pr_processed = self._preprocess_data(df_2b, df_pr)
            self._track_credit_notes(df_2b_processed, df_pr_processed)
            
            df_2b_processed = DataProcessor.optimize_dtypes(df_2b_processed)
            df_pr_processed = DataProcessor.optimize_dtypes(df_pr_processed)
            
            self.metrics.memory_usage_mb = (
                df_2b_processed.memory_usage(deep=True).sum() + 
                df_pr_processed.memory_usage(deep=True).sum()
            ) / (1024 * 1024)
            
            matched_2b_indices = set()
            matched_pr_indices = set()
            all_matches = []
            
            strategies = self._get_strategy_pipeline()
            
            for strategy in strategies:
                strategy_start = time.time()
                self.logger.debug(f"Executing strategy: {strategy.value}")
                
                try:
                    matches = self._execute_strategy(
                        strategy, df_2b_processed, df_pr_processed,
                        matched_2b_indices, matched_pr_indices
                    )
                    
                    if matches and not matches['dataframe'].empty:
                        # Filter matches with confidence >= min_confidence
                        df_match = matches['dataframe'].copy()
                        if 'MATCH_TIER' not in df_match.columns:
                            df_match['MATCH_TIER'] = matches.get('tier', 2)
                        if 'CONFIDENCE' in df_match.columns:
                            df_match = df_match[df_match['CONFIDENCE'] >= self.config.min_confidence_for_match]
                        if df_match.empty:
                            continue
                        
                        with self._lock:
                            all_matches.append(df_match)
                            matched_2b_indices.update(matches['matched_2b'])
                            matched_pr_indices.update(matches['matched_pr'])
                            count = len(df_match)
                            self.stats[f'matches_{strategy.value}'] += count
                            self.logger.info(f"Strategy {strategy.value}: matched {count} records")
                            self.metrics.strategy_usage[strategy.value] = count
                            self.metrics.matched_records += count
                            
                            if 'IS_CREDIT_NOTE_2B' in df_match.columns:
                                cn_matches = df_match[df_match['IS_CREDIT_NOTE_2B'] == True]
                                self.metrics.matched_credit_notes += len(cn_matches)
                
                except Exception as e:
                    self.logger.error(f"Strategy {strategy.value} failed: {str(e)}")
                    self.metrics.errors.append(f"Strategy {strategy.value}: {str(e)}")
                    continue
                
                self.logger.debug(f"Strategy {strategy.value} completed in {time.time() - strategy_start:.2f}s")
            
            # Determine unmatched records
            unmatched_2b = df_2b_processed[~df_2b_processed.index.isin(matched_2b_indices)].copy()
            unmatched_pr = df_pr_processed[~df_pr_processed.index.isin(matched_pr_indices)].copy()
            
            self.metrics.unmatched_2b = len(unmatched_2b)
            self.metrics.unmatched_pr = len(unmatched_pr)
            
            if 'IS_CREDIT_NOTE' in unmatched_2b.columns:
                self.metrics.unmatched_credit_notes_2b = len(unmatched_2b[unmatched_2b['IS_CREDIT_NOTE'] == True])
            if 'IS_CREDIT_NOTE' in unmatched_pr.columns:
                self.metrics.unmatched_credit_notes_pr = len(unmatched_pr[unmatched_pr['IS_CREDIT_NOTE'] == True])
            
            final_df = self._assemble_results(all_matches, unmatched_2b, unmatched_pr)
            
            # Ensure numeric columns are properly typed
            for col in ['TAXABLE VALUE', 'IGST', 'CGST', 'SGST', 'TOTAL TAX', 'TOTAL VALUE', 'CONFIDENCE']:
                for suffix in ['', '_2B', '_PR']:
                    col_name = f"{col}{suffix}"
                    if col_name in final_df.columns:
                        final_df[col_name] = pd.to_numeric(final_df[col_name], errors='coerce').fillna(0.0)
            
            stats = self._calculate_statistics(final_df, df_2b_processed, df_pr_processed)
            stats['processing_time'] = time.time() - start_time
            stats['strategy_breakdown'] = dict(self.stats)
            stats['tier_distribution'] = self.metrics.tier_distribution
            
            stats['credit_notes'] = {
                'total_2b': self.metrics.credit_notes_2b,
                'total_pr': self.metrics.credit_notes_pr,
                'matched': self.metrics.matched_credit_notes,
                'unmatched_2b': self.metrics.unmatched_credit_notes_2b,
                'unmatched_pr': self.metrics.unmatched_credit_notes_pr,
                'total_amount_2b': float(self.metrics.total_credit_amount_2b),
                'total_amount_pr': float(self.metrics.total_credit_amount_pr),
                'difference': float(self.metrics.credit_note_difference)
            }
            
            self.metrics.end_time = datetime.now()
            self.metrics.processing_time_seconds = stats['processing_time']
            
            self.logger.info(f"Reconciliation completed in {stats['processing_time']:.2f}s")
            
            return final_df, stats
        
        except Exception as e:
            self.logger.error(f"Reconciliation failed: {str(e)}")
            self.logger.error(traceback.format_exc())
            raise
    
    def _track_credit_notes(self, df_2b: pd.DataFrame, df_pr: pd.DataFrame):
        """Track credit note metrics"""
        if 'IS_CREDIT_NOTE' in df_2b.columns:
            self.metrics.credit_notes_2b = len(df_2b[df_2b['IS_CREDIT_NOTE'] == True])
            self.metrics.total_credit_amount_2b = float(df_2b[df_2b['IS_CREDIT_NOTE'] == True]['TAXABLE VALUE'].sum() if not df_2b.empty else 0)
        else:
            self.metrics.credit_notes_2b = 0
            self.metrics.total_credit_amount_2b = 0
        
        if 'IS_CREDIT_NOTE' in df_pr.columns:
            self.metrics.credit_notes_pr = len(df_pr[df_pr['IS_CREDIT_NOTE'] == True])
            self.metrics.total_credit_amount_pr = float(df_pr[df_pr['IS_CREDIT_NOTE'] == True]['TAXABLE VALUE'].sum() if not df_pr.empty else 0)
        else:
            self.metrics.credit_notes_pr = 0
            self.metrics.total_credit_amount_pr = 0
        
        self.metrics.credit_note_difference = float(
            self.metrics.total_credit_amount_2b - self.metrics.total_credit_amount_pr
        )
    
    def _execute_strategy(self, strategy: MatchStrategy, df_2b: pd.DataFrame, df_pr: pd.DataFrame,
                          matched_2b: set, matched_pr: set) -> Optional[Dict]:
        """Execute a specific matching strategy"""
        strategy_map = {
            MatchStrategy.EXACT: self._match_exact,
            MatchStrategy.SMART: self._match_smart,
            MatchStrategy.VALUE_BASED: self._match_value_based,
            MatchStrategy.FUZZY_NAME: self._match_fuzzy_name,
            MatchStrategy.PATTERN_RECOGNITION: self._match_pattern_recognition,
            MatchStrategy.SEQUENTIAL: self._match_sequential,
            MatchStrategy.AGGREGATE: self._match_aggregate,
            MatchStrategy.PERCENTAGE: self._match_percentage,
            MatchStrategy.WILDCARD: self._match_wildcard,
            MatchStrategy.AI_ENHANCED: self._match_ai_enhanced,
            MatchStrategy.CREDIT_NOTE: self._match_credit_notes,
            MatchStrategy.NEGATIVE_VALUE: self._match_negative_values,
        }
        
        match_func = strategy_map.get(strategy)
        if match_func:
            return match_func(df_2b, df_pr, matched_2b, matched_pr)
        return None
    
    def _get_strategy_pipeline(self) -> List[MatchStrategy]:
        """Get ordered list of strategies based on configuration"""
        strategies = [MatchStrategy.EXACT, MatchStrategy.SMART]
        
        if self.config.enable_credit_note_matching:
            strategies.append(MatchStrategy.CREDIT_NOTE)
        
        if self.config.enable_negative_value_matching:
            strategies.append(MatchStrategy.NEGATIVE_VALUE)
        
        if self.config.enable_pattern_recognition:
            strategies.append(MatchStrategy.PATTERN_RECOGNITION)
        if self.config.enable_sequential_matching:
            strategies.append(MatchStrategy.SEQUENTIAL)
        if self.config.enable_aggregate_matching:
            strategies.append(MatchStrategy.AGGREGATE)
        if self.config.enable_percentage_matching:
            strategies.append(MatchStrategy.PERCENTAGE)
        if self.config.enable_wildcard_matching:
            strategies.append(MatchStrategy.WILDCARD)
        if self.config.enable_fuzzy_matching:
            strategies.append(MatchStrategy.FUZZY_NAME)
        if self.config.enable_ai_enhanced:
            strategies.append(MatchStrategy.AI_ENHANCED)
        
        strategies.append(MatchStrategy.VALUE_BASED)
        return strategies
    
    def _preprocess_data(self, df_2b: pd.DataFrame, df_pr: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """Advanced preprocessing with credit note detection"""
        self.logger.debug("Starting data preprocessing with credit note detection")
        
        processor = DataProcessor()
        
        processed_dfs = []
        for df, name in [(df_2b, '2B'), (df_pr, 'PR')]:
            df = processor.standardize_columns(df)
            df = processor.clean_data(df)
            
            required_cols = ['SUPPLIER GSTIN', 'DOCUMENT NUMBER', 'TAXABLE VALUE', 
                           'SUPPLIER NAME', 'DOCUMENT DATE']
            
            missing = processor.get_missing_columns(df, required_cols)
            if missing:
                self.logger.warning(f"{name} missing columns: {missing}")
                for col in missing:
                    if col == 'SUPPLIER GSTIN' and 'GSTIN' in df.columns:
                        df.rename(columns={'GSTIN': 'SUPPLIER GSTIN'}, inplace=True)
                    elif col == 'DOCUMENT NUMBER' and 'INVOICE NO' in df.columns:
                        df.rename(columns={'INVOICE NO': 'DOCUMENT NUMBER'}, inplace=True)
                    elif col == 'DOCUMENT DATE' and 'INVOICE DATE' in df.columns:
                        df.rename(columns={'INVOICE DATE': 'DOCUMENT DATE'}, inplace=True)
                    elif col == 'TAXABLE VALUE' and 'AMOUNT' in df.columns:
                        df.rename(columns={'AMOUNT': 'TAXABLE VALUE'}, inplace=True)
            
            # Guarantee all required columns exist so downstream code never KeyErrors
            # on a missing/empty column — fill with sensible empty defaults instead.
            for col in required_cols:
                if col not in df.columns:
                    df[col] = 0.0 if col == 'TAXABLE VALUE' else ''
            
            # Convert TAXABLE VALUE to numeric
            if 'TAXABLE VALUE' in df.columns:
                df['TAXABLE VALUE'] = pd.to_numeric(df['TAXABLE VALUE'], errors='coerce').fillna(0.0)
            
            df['PAN'] = df['SUPPLIER GSTIN'].apply(extract_pan_from_gstin)
            df['NORM_DOC'] = df['DOCUMENT NUMBER'].apply(normalize_document_number)
            df['DOC_DATE_PARSED'] = df['DOCUMENT DATE'].apply(parse_date)
            
            tax_cols = [col for col in ['IGST', 'CGST', 'SGST'] if col in df.columns]
            if tax_cols:
                df['TOTAL_TAX'] = df[tax_cols].sum(axis=1, skipna=True)
            elif 'TOTAL TAX' in df.columns:
                df['TOTAL_TAX'] = pd.to_numeric(df['TOTAL TAX'], errors='coerce').fillna(0.0)
            else:
                df['TOTAL_TAX'] = 0.0
            
            df['TAXABLE VALUE'] = df['TAXABLE VALUE'].apply(safe_float_convert)
            df['TOTAL_VALUE'] = df['TAXABLE VALUE'] + df['TOTAL_TAX']
            
            df['MONTH_NUM'] = df['DOC_DATE_PARSED'].apply(
                lambda x: x.month if pd.notna(x) else None
            )
            df['YEAR'] = df['DOC_DATE_PARSED'].apply(
                lambda x: x.year if pd.notna(x) else None
            )
            df['QUARTER'] = df['MONTH_NUM'].apply(
                lambda x: get_quarter(x) if pd.notna(x) else None
            )
            
            if 'DOC TYPE' in df.columns:
                df['DOC_TYPE_NORM'] = df['DOC TYPE'].apply(
                    lambda x: str(x).upper().strip() if pd.notna(x) else 'INVOICE'
                )
            else:
                df['DOC_TYPE_NORM'] = 'INVOICE'
            
            df = processor.detect_credit_notes(df)
            
            df['IS_CREDIT'] = df['IS_CREDIT_NOTE'] | df['NEGATIVE_AMOUNT']
            df['ABS_TAXABLE'] = df['TAXABLE VALUE'].abs()
            df['SIGNED_TAXABLE'] = df['TAXABLE VALUE']
            
            if 'REFERENCE DOCUMENT' not in df.columns:
                df['REFERENCE DOCUMENT'] = ''
            else:
                df['REFERENCE DOCUMENT'] = df['REFERENCE DOCUMENT'].fillna('')
            
            if self.config.validate_gstin:
                df['GSTIN_VALID'] = df['SUPPLIER GSTIN'].apply(validate_gstin_format)
            
            df['orig_idx'] = df.index
            
            processed_dfs.append(df)
        
        self.logger.debug("Preprocessing complete")
        return processed_dfs[0], processed_dfs[1]
    
    # ========================================================================
    # CREDIT NOTE MATCHING STRATEGIES - STRICTENED
    # ========================================================================
    
    def _match_credit_notes(self, df_2b: pd.DataFrame, df_pr: pd.DataFrame,
                           matched_2b: set, matched_pr: set) -> Optional[Dict]:
        """Credit Note specific matching strategy - strict criteria"""
        self.logger.debug("Executing credit note match strategy")
        
        if not self.config.enable_credit_note_matching:
            return None
        
        available_2b = df_2b[~df_2b.index.isin(matched_2b)].copy()
        available_pr = df_pr[~df_pr.index.isin(matched_pr)].copy()
        
        if available_2b.empty or available_pr.empty:
            return None
        
        cn_2b = available_2b[available_2b['IS_CREDIT'] == True].copy()
        cn_pr = available_pr[available_pr['IS_CREDIT'] == True].copy()
        
        if cn_2b.empty or cn_pr.empty:
            return None
        
        matches = []
        tolerance = self.config.tolerance_amount
        max_diff = tolerance  # stricter: only tolerance (not 2×)
        
        for pan, group_2b in cn_2b.groupby('PAN'):
            group_pr = cn_pr[cn_pr['PAN'] == pan]
            
            if group_pr.empty:
                continue
            
            cross = pd.merge(
                group_2b, group_pr,
                how='cross',
                suffixes=('_2B', '_PR')
            )
            
            cross['ABS_TAX_DIFF'] = (
                cross['ABS_TAXABLE_2B'] - cross['ABS_TAXABLE_PR']
            ).abs()
            
            cross['REF_MATCH'] = cross.apply(
                lambda r: self._compare_reference_documents(
                    r.get('REFERENCE DOCUMENT_2B', ''),
                    r.get('REFERENCE DOCUMENT_PR', '')
                ),
                axis=1
            )
            
            cross['DATE_DIFF'] = cross.apply(
                lambda r: self._calculate_date_diff(
                    r.get('DOC_DATE_PARSED_2B'),
                    r.get('DOC_DATE_PARSED_PR')
                ),
                axis=1
            )
            
            cross['DOC_SIM'] = cross.apply(
                lambda r: self._calculate_document_similarity(
                    r.get('NORM_DOC_2B', ''),
                    r.get('NORM_DOC_PR', '')
                ),
                axis=1
            )
            
            # Stricter: doc_sim >= 0.7, ref_match >= 0.8, date diff <= tolerance, tax diff <= max_diff
            valid = cross[
                (cross['ABS_TAX_DIFF'] <= max_diff) &
                (cross['DATE_DIFF'] <= self.config.date_tolerance_days) &
                (cross['REF_MATCH'] >= 0.8) &
                (cross['DOC_SIM'] >= 0.7)
            ]
            
            if not valid.empty:
                valid['CONFIDENCE'] = (
                    (1.0 - (valid['ABS_TAX_DIFF'] / (max_diff + 1))) * 0.3 +
                    valid['REF_MATCH'] * 0.4 +
                    valid['DOC_SIM'] * 0.3
                )
                valid['CONFIDENCE'] = valid['CONFIDENCE'].clip(0, 1)
                valid = valid.sort_values(['CONFIDENCE', 'REF_MATCH'], ascending=[False, False])
                
                valid = valid.drop_duplicates(subset=['orig_idx_2B'], keep='first')
                valid = valid.drop_duplicates(subset=['orig_idx_PR'], keep='first')
                
                matches.append(valid)
        
        if not matches:
            return None
        
        merged = pd.concat(matches, ignore_index=True)
        
        return {
            'dataframe': merged,
            'matched_2b': set(merged['orig_idx_2B']),
            'matched_pr': set(merged['orig_idx_PR']),
            'strategy': MatchStrategy.CREDIT_NOTE.value,
            'tier': MatchTier.TIER_8_CREDIT_NOTE.value,
            'confidence': merged['CONFIDENCE'].mean()
        }
    
    def _match_negative_values(self, df_2b: pd.DataFrame, df_pr: pd.DataFrame,
                               matched_2b: set, matched_pr: set) -> Optional[Dict]:
        """Negative value matching strategy - strict"""
        self.logger.debug("Executing negative value match strategy")
        
        if not self.config.enable_negative_value_matching:
            return None
        
        available_2b = df_2b[~df_2b.index.isin(matched_2b)].copy()
        available_pr = df_pr[~df_pr.index.isin(matched_pr)].copy()
        
        if available_2b.empty or available_pr.empty:
            return None
        
        neg_2b = available_2b[available_2b['TAXABLE VALUE'] < 0].copy()
        neg_pr = available_pr[available_pr['TAXABLE VALUE'] < 0].copy()
        
        if neg_2b.empty or neg_pr.empty:
            return None
        
        matches = []
        tolerance = self.config.tolerance_amount
        max_diff = tolerance
        
        for pan, group_2b in neg_2b.groupby('PAN'):
            group_pr = neg_pr[neg_pr['PAN'] == pan]
            
            if group_pr.empty:
                continue
            
            cross = pd.merge(
                group_2b, group_pr,
                how='cross',
                suffixes=('_2B', '_PR')
            )
            
            cross['ABS_TAX_DIFF'] = (
                cross['TAXABLE VALUE_2B'].abs() - cross['TAXABLE VALUE_PR'].abs()
            ).abs()
            
            cross['DATE_DIFF'] = cross.apply(
                lambda r: self._calculate_date_diff(
                    r.get('DOC_DATE_PARSED_2B'),
                    r.get('DOC_DATE_PARSED_PR')
                ),
                axis=1
            )
            
            cross['DOC_SIM'] = cross.apply(
                lambda r: self._calculate_document_similarity(
                    r.get('NORM_DOC_2B', ''),
                    r.get('NORM_DOC_PR', '')
                ),
                axis=1
            )
            
            valid = cross[
                (cross['ABS_TAX_DIFF'] <= max_diff) &
                (cross['DATE_DIFF'] <= self.config.date_tolerance_days) &
                (cross['DOC_SIM'] >= 0.7)  # stricter
            ]
            
            if not valid.empty:
                valid['CONFIDENCE'] = (
                    (1.0 - (valid['ABS_TAX_DIFF'] / (max_diff + 1))) * 0.6 +
                    valid['DOC_SIM'] * 0.4
                )
                valid['CONFIDENCE'] = valid['CONFIDENCE'].clip(0, 1)
                valid = valid.sort_values(['CONFIDENCE', 'ABS_TAX_DIFF'], ascending=[False, True])
                
                valid = valid.drop_duplicates(subset=['orig_idx_2B'], keep='first')
                valid = valid.drop_duplicates(subset=['orig_idx_PR'], keep='first')
                
                matches.append(valid)
        
        if not matches:
            return None
        
        merged = pd.concat(matches, ignore_index=True)
        
        return {
            'dataframe': merged,
            'matched_2b': set(merged['orig_idx_2B']),
            'matched_pr': set(merged['orig_idx_PR']),
            'strategy': MatchStrategy.NEGATIVE_VALUE.value,
            'tier': MatchTier.TIER_9_NEGATIVE.value,
            'confidence': merged['CONFIDENCE'].mean()
        }
    
    # ========================================================================
    # STRICTENED MATCHING STRATEGIES
    # ========================================================================
    
    def _match_exact(self, df_2b: pd.DataFrame, df_pr: pd.DataFrame,
                     matched_2b: set, matched_pr: set) -> Optional[Dict]:
        """Tier 1: Exact matching - strict criteria"""
        self.logger.debug("Executing exact match strategy")
        
        available_2b = df_2b[~df_2b.index.isin(matched_2b)].copy()
        available_pr = df_pr[~df_pr.index.isin(matched_pr)].copy()
        
        if available_2b.empty or available_pr.empty:
            return None
        
        matches = []
        for pan, group_2b in available_2b.groupby('PAN'):
            group_pr = available_pr[available_pr['PAN'] == pan]
            if group_pr.empty:
                continue
            
            cross = pd.merge(
                group_2b, group_pr,
                how='cross',
                suffixes=('_2B', '_PR')
            )
            
            cross['EXACT_KEY'] = (
                cross['PAN_2B'].astype(str) + '|' +
                cross['NORM_DOC_2B'].astype(str) + '|' +
                cross['TAXABLE VALUE_2B'].round(2).astype(str) + '|' +
                cross['IS_CREDIT_2B'].astype(str)
            )
            cross['EXACT_KEY_PR'] = (
                cross['PAN_PR'].astype(str) + '|' +
                cross['NORM_DOC_PR'].astype(str) + '|' +
                cross['TAXABLE VALUE_PR'].round(2).astype(str) + '|' +
                cross['IS_CREDIT_PR'].astype(str)
            )
            
            cross['DATE_DIFF'] = cross.apply(
                lambda r: self._calculate_date_diff(
                    r.get('DOC_DATE_PARSED_2B'),
                    r.get('DOC_DATE_PARSED_PR')
                ),
                axis=1
            )
            
            cross['RAW_DOC_SIM'] = cross.apply(
                lambda r: self._calculate_document_similarity(
                    r.get('DOCUMENT NUMBER_2B', ''),
                    r.get('DOCUMENT NUMBER_PR', '')
                ),
                axis=1
            )
            
            # Strict: exact key match, date diff <= 1, raw doc sim >= 0.95
            valid = cross[
                (cross['EXACT_KEY'] == cross['EXACT_KEY_PR']) &
                (cross['DATE_DIFF'] <= 1) &
                (cross['RAW_DOC_SIM'] >= 0.95)
            ]
            
            if not valid.empty:
                valid['CONFIDENCE'] = 1.0
                valid['TAX_DIFF'] = 0.0
                valid = valid.drop_duplicates(subset=['orig_idx_2B'], keep='first')
                valid = valid.drop_duplicates(subset=['orig_idx_PR'], keep='first')
                matches.append(valid)
        
        if not matches:
            return None
        
        merged = pd.concat(matches, ignore_index=True)
        
        return {
            'dataframe': merged,
            'matched_2b': set(merged['orig_idx_2B']),
            'matched_pr': set(merged['orig_idx_PR']),
            'strategy': MatchStrategy.EXACT.value,
            'tier': MatchTier.TIER_1_EXACT.value,
            'confidence': 1.0
        }
    
    def _match_smart(self, df_2b: pd.DataFrame, df_pr: pd.DataFrame,
                     matched_2b: set, matched_pr: set) -> Optional[Dict]:
        """Tier 2: Smart matching with stricter criteria"""
        self.logger.debug("Executing smart match strategy")
        
        available_2b = df_2b[~df_2b.index.isin(matched_2b)].copy()
        available_pr = df_pr[~df_pr.index.isin(matched_pr)].copy()
        
        if available_2b.empty or available_pr.empty:
            return None
        
        matches = []
        tolerance = self.config.tolerance_amount
        max_diff = tolerance  # stricter: use tolerance directly
        
        for pan, group_2b in available_2b.groupby('PAN'):
            group_pr = available_pr[available_pr['PAN'] == pan]
            
            if group_pr.empty:
                continue
            
            cross = pd.merge(
                group_2b, group_pr,
                how='cross',
                suffixes=('_2B', '_PR')
            )
            
            cross['TAX_DIFF'] = (
                cross['TAXABLE VALUE_2B'] - cross['TAXABLE VALUE_PR']
            ).abs()
            
            cross['DATE_DIFF'] = cross.apply(
                lambda r: self._calculate_date_diff(
                    r.get('DOC_DATE_PARSED_2B'),
                    r.get('DOC_DATE_PARSED_PR')
                ),
                axis=1
            )
            
            cross['DOC_SIM'] = cross.apply(
                lambda r: self._calculate_document_similarity(
                    r.get('NORM_DOC_2B', ''),
                    r.get('NORM_DOC_PR', '')
                ),
                axis=1
            )
            
            cross['CREDIT_TYPE_MATCH'] = (
                cross['IS_CREDIT_2B'] == cross['IS_CREDIT_PR']
            ).astype(int)
            
            valid = cross[
                (cross['TAX_DIFF'] <= max_diff) &
                (cross['DATE_DIFF'] <= self.config.date_tolerance_days) &
                (cross['CREDIT_TYPE_MATCH'] == 1) &
                (cross['DOC_SIM'] >= 0.8)  # stricter
            ]
            
            if not valid.empty:
                valid['CONFIDENCE'] = (
                    (1.0 - (valid['TAX_DIFF'] / (max_diff + 1))) * 0.4 +
                    (1.0 - (valid['DATE_DIFF'] / (self.config.date_tolerance_days + 1))) * 0.2 +
                    valid['DOC_SIM'] * 0.4
                )
                valid['CONFIDENCE'] = valid['CONFIDENCE'].clip(0, 1)
                valid = valid.sort_values(['CONFIDENCE', 'TAX_DIFF'], ascending=[False, True])
                
                valid = valid.drop_duplicates(subset=['orig_idx_2B'], keep='first')
                valid = valid.drop_duplicates(subset=['orig_idx_PR'], keep='first')
                
                matches.append(valid)
        
        if not matches:
            return None
        
        merged = pd.concat(matches, ignore_index=True)
        
        return {
            'dataframe': merged,
            'matched_2b': set(merged['orig_idx_2B']),
            'matched_pr': set(merged['orig_idx_PR']),
            'strategy': MatchStrategy.SMART.value,
            'tier': MatchTier.TIER_2_SMART.value,
            'confidence': merged['CONFIDENCE'].mean()
        }
    
    def _match_value_based(self, df_2b: pd.DataFrame, df_pr: pd.DataFrame,
                           matched_2b: set, matched_pr: set) -> Optional[Dict]:
        """Tier 3: Value-based matching with stricter criteria"""
        self.logger.debug("Executing value-based match strategy")
        
        available_2b = df_2b[~df_2b.index.isin(matched_2b)].copy()
        available_pr = df_pr[~df_pr.index.isin(matched_pr)].copy()
        
        if available_2b.empty or available_pr.empty:
            return None
        
        matches = []
        tolerance = self.config.tolerance_amount
        max_diff = tolerance
        
        for pan, group_2b in available_2b.groupby('PAN'):
            group_pr = available_pr[available_pr['PAN'] == pan]
            
            if group_pr.empty:
                continue
            
            cross = pd.merge(
                group_2b, group_pr,
                how='cross',
                suffixes=('_2B', '_PR')
            )
            
            cross['TAX_DIFF'] = (
                cross['TAXABLE VALUE_2B'] - cross['TAXABLE VALUE_PR']
            ).abs()
            
            cross['DOC_SIM'] = cross.apply(
                lambda r: self._calculate_document_similarity(
                    r.get('NORM_DOC_2B', ''),
                    r.get('NORM_DOC_PR', '')
                ),
                axis=1
            )
            
            cross['CREDIT_TYPE_MATCH'] = (
                cross['IS_CREDIT_2B'] == cross['IS_CREDIT_PR']
            ).astype(int)
            
            valid = cross[
                (cross['TAX_DIFF'] <= max_diff) &
                (cross['CREDIT_TYPE_MATCH'] == 1) &
                (cross['DOC_SIM'] >= 0.7)  # stricter
            ]
            
            if not valid.empty:
                valid['CONFIDENCE'] = (
                    (1.0 - (valid['TAX_DIFF'] / (max_diff + 1))) * 0.6 +
                    valid['DOC_SIM'] * 0.4
                )
                valid['CONFIDENCE'] = valid['CONFIDENCE'].clip(0, 1)
                valid = valid.sort_values(['CONFIDENCE', 'TAX_DIFF'], ascending=[False, True])
                
                valid = valid.drop_duplicates(subset=['orig_idx_2B'], keep='first')
                valid = valid.drop_duplicates(subset=['orig_idx_PR'], keep='first')
                
                matches.append(valid)
        
        if not matches:
            return None
        
        merged = pd.concat(matches, ignore_index=True)
        
        return {
            'dataframe': merged,
            'matched_2b': set(merged['orig_idx_2B']),
            'matched_pr': set(merged['orig_idx_PR']),
            'strategy': MatchStrategy.VALUE_BASED.value,
            'tier': MatchTier.TIER_3_VALUE.value,
            'confidence': merged['CONFIDENCE'].mean()
        }
    
    def _match_fuzzy_name(self, df_2b: pd.DataFrame, df_pr: pd.DataFrame,
                          matched_2b: set, matched_pr: set) -> Optional[Dict]:
        """Tier 4: Fuzzy name matching with strict criteria"""
        self.logger.debug("Executing fuzzy name match strategy")
        
        if not self.config.enable_fuzzy_matching:
            return None
        
        available_2b = df_2b[~df_2b.index.isin(matched_2b)].copy()
        available_pr = df_pr[~df_pr.index.isin(matched_pr)].copy()
        
        if available_2b.empty or available_pr.empty:
            return None
        
        matches = []
        tolerance = self.config.tolerance_amount
        max_diff = tolerance
        
        for pan, group_2b in available_2b.groupby('PAN'):
            group_pr = available_pr[available_pr['PAN'] == pan]
            
            if group_pr.empty:
                continue
            
            cross = pd.merge(
                group_2b, group_pr,
                how='cross',
                suffixes=('_2B', '_PR')
            )
            
            cross['FUZZY_SCORE'] = cross.apply(
                lambda r: self._calculate_fuzzy_score(
                    r.get('SUPPLIER NAME_2B', ''),
                    r.get('SUPPLIER NAME_PR', '')
                ),
                axis=1
            )
            
            cross['TAX_DIFF'] = (
                cross['TAXABLE VALUE_2B'] - cross['TAXABLE VALUE_PR']
            ).abs()
            
            cross['DOC_SIM'] = cross.apply(
                lambda r: self._calculate_document_similarity(
                    r.get('NORM_DOC_2B', ''),
                    r.get('NORM_DOC_PR', '')
                ),
                axis=1
            )
            
            cross['SAME_TYPE'] = (
                cross['IS_CREDIT_2B'] == cross['IS_CREDIT_PR']
            ).astype(int)
            
            valid = cross[
                (cross['FUZZY_SCORE'] >= self.config.fuzzy_threshold) &
                (cross['TAX_DIFF'] <= max_diff) &
                (cross['SAME_TYPE'] == 1) &
                (cross['DOC_SIM'] >= 0.7)  # stricter
            ]
            
            if not valid.empty:
                valid['CONFIDENCE'] = (
                    (cross['FUZZY_SCORE'] / 100) * 0.3 +
                    (1.0 - (valid['TAX_DIFF'] / (max_diff + 1))) * 0.3 +
                    valid['DOC_SIM'] * 0.4
                )
                valid['CONFIDENCE'] = valid['CONFIDENCE'].clip(0, 1)
                valid = valid.sort_values(['CONFIDENCE', 'FUZZY_SCORE'], ascending=[False, False])
                
                valid = valid.drop_duplicates(subset=['orig_idx_2B'], keep='first')
                valid = valid.drop_duplicates(subset=['orig_idx_PR'], keep='first')
                
                matches.append(valid)
        
        if not matches:
            return None
        
        merged = pd.concat(matches, ignore_index=True)
        
        return {
            'dataframe': merged,
            'matched_2b': set(merged['orig_idx_2B']),
            'matched_pr': set(merged['orig_idx_PR']),
            'strategy': MatchStrategy.FUZZY_NAME.value,
            'tier': MatchTier.TIER_4_FUZZY.value,
            'confidence': merged['CONFIDENCE'].mean()
        }
    
    def _match_pattern_recognition(self, df_2b: pd.DataFrame, df_pr: pd.DataFrame,
                                   matched_2b: set, matched_pr: set) -> Optional[Dict]:
        """Tier 5: Pattern-based matching with strict type check"""
        self.logger.debug("Executing pattern recognition strategy")
        
        if not self.config.enable_pattern_recognition:
            return None
        
        available_2b = df_2b[~df_2b.index.isin(matched_2b)].copy()
        available_pr = df_pr[~df_pr.index.isin(matched_pr)].copy()
        
        if available_2b.empty or available_pr.empty:
            return None
        
        available_2b['DOC_PATTERN'] = available_2b['NORM_DOC'].apply(
            lambda x: self._extract_doc_pattern(x)
        )
        available_pr['DOC_PATTERN'] = available_pr['NORM_DOC'].apply(
            lambda x: self._extract_doc_pattern(x)
        )
        
        matches = []
        tolerance = self.config.tolerance_amount
        max_diff = tolerance
        
        for pan, group_2b in available_2b.groupby('PAN'):
            group_pr = available_pr[available_pr['PAN'] == pan]
            
            if group_pr.empty:
                continue
            
            for idx_2b, row_2b in group_2b.iterrows():
                best_match = None
                best_score = 0
                
                for idx_pr, row_pr in group_pr.iterrows():
                    if row_2b['IS_CREDIT'] != row_pr['IS_CREDIT']:
                        continue
                    
                    pattern_match = self._compare_doc_patterns(
                        row_2b['DOC_PATTERN'],
                        row_pr['DOC_PATTERN']
                    )
                    
                    tax_diff = abs(row_2b['TAXABLE VALUE'] - row_pr['TAXABLE VALUE'])
                    tax_sim = 1.0 - min(tax_diff / (max_diff + 1), 1.0)
                    
                    total_score = pattern_match * 0.6 + tax_sim * 0.4
                    
                    if total_score > best_score and tax_diff <= max_diff and pattern_match >= 0.7:
                        best_score = total_score
                        best_match = (idx_2b, idx_pr, row_2b, row_pr)
                
                if best_match and best_score >= 0.7:
                    confidence = best_score
                    matches.append({
                        'row_2b': best_match[2],
                        'row_pr': best_match[3],
                        'confidence': confidence
                    })
        
        if not matches:
            return None
        
        merged = pd.DataFrame([{
            **{f"{k}_2B": v for k, v in m['row_2b'].to_dict().items()},
            **{f"{k}_PR": v for k, v in m['row_pr'].to_dict().items()},
            'CONFIDENCE': m['confidence']
        } for m in matches])
        
        return {
            'dataframe': merged,
            'matched_2b': set(merged['orig_idx_2B']),
            'matched_pr': set(merged['orig_idx_PR']),
            'strategy': MatchStrategy.PATTERN_RECOGNITION.value,
            'tier': MatchTier.TIER_5_PATTERN.value,
            'confidence': merged['CONFIDENCE'].mean()
        }
    
    def _match_sequential(self, df_2b: pd.DataFrame, df_pr: pd.DataFrame,
                          matched_2b: set, matched_pr: set) -> Optional[Dict]:
        """Tier 6: Sequential matching with type consistency"""
        self.logger.debug("Executing sequential match strategy")
        
        if not self.config.enable_sequential_matching:
            return None
        
        available_2b = df_2b[~df_2b.index.isin(matched_2b)].copy()
        available_pr = df_pr[~df_pr.index.isin(matched_pr)].copy()
        
        if available_2b.empty or available_pr.empty:
            return None
        
        available_2b['SEQ_GROUP'] = available_2b['IS_CREDIT'].astype(str)
        available_pr['SEQ_GROUP'] = available_pr['IS_CREDIT'].astype(str)
        
        matches = []
        tolerance = self.config.tolerance_amount
        max_diff = tolerance
        
        for pan, group_2b in available_2b.groupby('PAN'):
            group_pr = available_pr[available_pr['PAN'] == pan]
            
            if group_pr.empty:
                continue
            
            for seq_group, g_2b in group_2b.groupby('SEQ_GROUP'):
                g_pr = group_pr[group_pr['SEQ_GROUP'] == seq_group]
                
                if g_pr.empty:
                    continue
                
                g_2b = g_2b.sort_values('NORM_DOC')
                g_pr = g_pr.sort_values('NORM_DOC')
                
                window_size = 5
                for i in range(len(g_2b)):
                    for j in range(max(0, i - window_size), min(len(g_pr), i + window_size + 1)):
                        row_2b = g_2b.iloc[i]
                        row_pr = g_pr.iloc[j]
                        
                        seq_distance = abs(i - j)
                        if seq_distance > window_size:
                            continue
                        
                        tax_diff = abs(row_2b['TAXABLE VALUE'] - row_pr['TAXABLE VALUE'])
                        
                        doc_sim = self._calculate_document_similarity(
                            row_2b['NORM_DOC'], row_pr['NORM_DOC']
                        )
                        
                        if tax_diff <= max_diff and doc_sim >= 0.7:
                            confidence = 1.0 - (seq_distance / window_size * 0.5) - \
                                         (tax_diff / (max_diff + 1) * 0.3) + \
                                         (doc_sim * 0.2)
                            confidence = max(0.5, min(1, confidence))  # ensure at least 0.5
                            
                            matches.append({
                                'row_2b': row_2b,
                                'row_pr': row_pr,
                                'confidence': confidence
                            })
        
        if not matches:
            return None
        
        merged = pd.DataFrame([{
            **{f"{k}_2B": v for k, v in m['row_2b'].to_dict().items()},
            **{f"{k}_PR": v for k, v in m['row_pr'].to_dict().items()},
            'CONFIDENCE': m['confidence']
        } for m in matches])
        
        return {
            'dataframe': merged,
            'matched_2b': set(merged['orig_idx_2B']),
            'matched_pr': set(merged['orig_idx_PR']),
            'strategy': MatchStrategy.SEQUENTIAL.value,
            'tier': MatchTier.TIER_6_SEQUENTIAL.value,
            'confidence': merged['CONFIDENCE'].mean()
        }
    
    def _match_aggregate(self, df_2b: pd.DataFrame, df_pr: pd.DataFrame,
                         matched_2b: set, matched_pr: set) -> Optional[Dict]:
        """Tier 7: Aggregate matching - only 1:1 groups with strict criteria"""
        self.logger.debug("Executing aggregate match strategy")
        
        if not self.config.enable_aggregate_matching:
            return None
        
        available_2b = df_2b[~df_2b.index.isin(matched_2b)].copy()
        available_pr = df_pr[~df_pr.index.isin(matched_pr)].copy()
        
        if available_2b.empty or available_pr.empty:
            return None
        
        matches = []
        tolerance = self.config.tolerance_amount
        max_diff = tolerance
        
        for is_credit in [False, True]:
            agg_2b = available_2b[available_2b['IS_CREDIT'] == is_credit].groupby(['PAN', 'MONTH_NUM', 'YEAR']).agg({
                'TAXABLE VALUE': 'sum',
                'TOTAL_TAX': 'sum',
                'orig_idx': list,
                'NORM_DOC': lambda x: list(x)
            }).reset_index()
            
            agg_pr = available_pr[available_pr['IS_CREDIT'] == is_credit].groupby(['PAN', 'MONTH_NUM', 'YEAR']).agg({
                'TAXABLE VALUE': 'sum',
                'TOTAL_TAX': 'sum',
                'orig_idx': list,
                'NORM_DOC': lambda x: list(x)
            }).reset_index()
            
            if agg_2b.empty or agg_pr.empty:
                continue
            
            # Only match groups where both sides have exactly 1 record
            agg_2b['count'] = agg_2b['orig_idx'].apply(len)
            agg_pr['count'] = agg_pr['orig_idx'].apply(len)
            
            agg_2b = agg_2b[agg_2b['count'] == 1]
            agg_pr = agg_pr[agg_pr['count'] == 1]
            
            if agg_2b.empty or agg_pr.empty:
                continue
            
            merged = pd.merge(
                agg_2b, agg_pr,
                on=['PAN', 'MONTH_NUM', 'YEAR'],
                suffixes=('_2B', '_PR'),
                how='inner'
            )
            
            merged['TAX_DIFF'] = (merged['TAXABLE VALUE_2B'] - merged['TAXABLE VALUE_PR']).abs()
            merged = merged[merged['TAX_DIFF'] <= max_diff]
            
            if not merged.empty:
                for _, row in merged.iterrows():
                    idx_2b = row['orig_idx_2B'][0]
                    idx_pr = row['orig_idx_PR'][0]
                    
                    row_2b = available_2b.loc[idx_2b]
                    row_pr = available_pr.loc[idx_pr]
                    
                    doc_sim = self._calculate_document_similarity(
                        row_2b['NORM_DOC'],
                        row_pr['NORM_DOC']
                    )
                    
                    # Require doc_sim >= 0.7
                    if doc_sim >= 0.7:
                        conf = (1.0 - (row['TAX_DIFF'] / (max_diff + 1))) * 0.6 + doc_sim * 0.4
                        conf = max(0.5, min(1, conf))
                        matches.append({
                            'row_2b': row_2b,
                            'row_pr': row_pr,
                            'confidence': conf,
                            'tax_diff': row['TAX_DIFF']
                        })
        
        if not matches:
            return None
        
        merged_df = pd.DataFrame([{
            **{f"{k}_2B": v for k, v in m['row_2b'].to_dict().items()},
            **{f"{k}_PR": v for k, v in m['row_pr'].to_dict().items()},
            'CONFIDENCE': m['confidence']
        } for m in matches])
        
        return {
            'dataframe': merged_df,
            'matched_2b': set(merged_df['orig_idx_2B']),
            'matched_pr': set(merged_df['orig_idx_PR']),
            'strategy': MatchStrategy.AGGREGATE.value,
            'tier': MatchTier.TIER_7_AGGREGATE.value,
            'confidence': merged_df['CONFIDENCE'].mean()
        }
    
    def _match_percentage(self, df_2b: pd.DataFrame, df_pr: pd.DataFrame,
                         matched_2b: set, matched_pr: set) -> Optional[Dict]:
        """Percentage-based matching with type match"""
        self.logger.debug("Executing percentage match strategy")
        
        if not self.config.enable_percentage_matching:
            return None
        
        available_2b = df_2b[~df_2b.index.isin(matched_2b)].copy()
        available_pr = df_pr[~df_pr.index.isin(matched_pr)].copy()
        
        if available_2b.empty or available_pr.empty:
            return None
        
        matches = []
        percentage_threshold = self.config.percentage_tolerance
        tolerance = self.config.tolerance_amount
        max_diff = tolerance
        
        for pan, group_2b in available_2b.groupby('PAN'):
            group_pr = available_pr[available_pr['PAN'] == pan]
            
            if group_pr.empty:
                continue
            
            cross = pd.merge(
                group_2b, group_pr,
                how='cross',
                suffixes=('_2B', '_PR')
            )
            
            abs_2b = cross['TAXABLE VALUE_2B'].abs()
            abs_pr = cross['TAXABLE VALUE_PR'].abs()
            
            cross['PCT_DIFF'] = (
                (abs_2b - abs_pr).abs() /
                (abs_2b + 1) * 100
            )
            
            cross['DOC_SIM'] = cross.apply(
                lambda r: self._calculate_document_similarity(
                    r.get('NORM_DOC_2B', ''),
                    r.get('NORM_DOC_PR', '')
                ),
                axis=1
            )
            
            cross['TYPE_MATCH'] = (
                cross['IS_CREDIT_2B'] == cross['IS_CREDIT_PR']
            ).astype(int)
            
            cross['TAX_DIFF'] = (cross['TAXABLE VALUE_2B'] - cross['TAXABLE VALUE_PR']).abs()
            
            valid = cross[
                (cross['PCT_DIFF'] <= percentage_threshold) &
                (cross['TAX_DIFF'] <= max_diff) &
                (cross['TYPE_MATCH'] == 1) &
                (cross['DOC_SIM'] >= 0.7)
            ]
            
            if not valid.empty:
                valid['CONFIDENCE'] = (
                    (1.0 - (valid['PCT_DIFF'] / percentage_threshold)) * 0.4 +
                    (1.0 - (valid['TAX_DIFF'] / (max_diff + 1))) * 0.3 +
                    valid['DOC_SIM'] * 0.3
                )
                valid['CONFIDENCE'] = valid['CONFIDENCE'].clip(0, 1)
                valid = valid.sort_values(['CONFIDENCE', 'PCT_DIFF'], ascending=[False, True])
                
                valid = valid.drop_duplicates(subset=['orig_idx_2B'], keep='first')
                valid = valid.drop_duplicates(subset=['orig_idx_PR'], keep='first')
                
                matches.append(valid)
        
        if not matches:
            return None
        
        merged = pd.concat(matches, ignore_index=True)
        
        return {
            'dataframe': merged,
            'matched_2b': set(merged['orig_idx_2B']),
            'matched_pr': set(merged['orig_idx_PR']),
            'strategy': MatchStrategy.PERCENTAGE.value,
            'tier': MatchTier.TIER_3_VALUE.value,
            'confidence': merged['CONFIDENCE'].mean()
        }
    
    def _match_wildcard(self, df_2b: pd.DataFrame, df_pr: pd.DataFrame,
                       matched_2b: set, matched_pr: set) -> Optional[Dict]:
        """Wildcard matching with type check"""
        self.logger.debug("Executing wildcard match strategy")
        
        if not self.config.enable_wildcard_matching:
            return None
        
        available_2b = df_2b[~df_2b.index.isin(matched_2b)].copy()
        available_pr = df_pr[~df_pr.index.isin(matched_pr)].copy()
        
        if available_2b.empty or available_pr.empty:
            return None
        
        matches = []
        tolerance = self.config.tolerance_amount
        max_diff = tolerance
        
        for pan, group_2b in available_2b.groupby('PAN'):
            group_pr = available_pr[available_pr['PAN'] == pan]
            
            if group_pr.empty:
                continue
            
            for _, row_2b in group_2b.iterrows():
                doc_num_2b = row_2b['NORM_DOC']
                
                for _, row_pr in group_pr.iterrows():
                    if row_2b['IS_CREDIT'] != row_pr['IS_CREDIT']:
                        continue
                    
                    doc_num_pr = row_pr['NORM_DOC']
                    wildcard_score = self._wildcard_match(doc_num_2b, doc_num_pr)
                    
                    if wildcard_score >= 0.7:
                        tax_diff = abs(row_2b['TAXABLE VALUE'] - row_pr['TAXABLE VALUE'])
                        
                        if tax_diff <= max_diff:
                            confidence = wildcard_score * 0.5 + \
                                         (1.0 - tax_diff / (max_diff + 1)) * 0.5
                            confidence = max(0.5, min(1, confidence))
                            
                            matches.append({
                                'row_2b': row_2b,
                                'row_pr': row_pr,
                                'confidence': confidence
                            })
        
        if not matches:
            return None
        
        merged = pd.DataFrame([{
            **{f"{k}_2B": v for k, v in m['row_2b'].to_dict().items()},
            **{f"{k}_PR": v for k, v in m['row_pr'].to_dict().items()},
            'CONFIDENCE': m['confidence']
        } for m in matches])
        
        return {
            'dataframe': merged,
            'matched_2b': set(merged['orig_idx_2B']),
            'matched_pr': set(merged['orig_idx_PR']),
            'strategy': MatchStrategy.WILDCARD.value,
            'tier': MatchTier.TIER_5_PATTERN.value,
            'confidence': merged['CONFIDENCE'].mean()
        }
    
    def _match_ai_enhanced(self, df_2b: pd.DataFrame, df_pr: pd.DataFrame,
                          matched_2b: set, matched_pr: set) -> Optional[Dict]:
        """AI-Enhanced matching with combined scores"""
        self.logger.debug("Executing AI-enhanced match strategy")
        
        if not self.config.enable_ai_enhanced:
            return None
        
        available_2b = df_2b[~df_2b.index.isin(matched_2b)].copy()
        available_pr = df_pr[~df_pr.index.isin(matched_pr)].copy()
        
        if available_2b.empty or available_pr.empty:
            return None
        
        matches = []
        tolerance = self.config.tolerance_amount
        max_diff = tolerance
        
        for pan, group_2b in available_2b.groupby('PAN'):
            group_pr = available_pr[available_pr['PAN'] == pan]
            
            if group_pr.empty:
                continue
            
            for idx_2b, row_2b in group_2b.iterrows():
                best_match = None
                best_score = 0
                
                for idx_pr, row_pr in group_pr.iterrows():
                    if row_2b['IS_CREDIT'] != row_pr['IS_CREDIT']:
                        continue
                    
                    name_sim = self._calculate_fuzzy_score(
                        row_2b.get('SUPPLIER NAME', ''),
                        row_pr.get('SUPPLIER NAME', '')
                    )
                    
                    doc_sim = self._wildcard_match(
                        row_2b.get('NORM_DOC', ''),
                        row_pr.get('NORM_DOC', '')
                    )
                    
                    tax_diff = abs(row_2b['TAXABLE VALUE'] - row_pr['TAXABLE VALUE'])
                    tax_sim = 1.0 - min(tax_diff / (max_diff + 1), 1.0)
                    
                    ref_match = self._compare_reference_documents(
                        row_2b.get('REFERENCE DOCUMENT', ''),
                        row_pr.get('REFERENCE DOCUMENT', '')
                    )
                    
                    score = (
                        (name_sim / 100) * 0.15 +
                        doc_sim * 0.25 +
                        tax_sim * 0.35 +
                        ref_match * 0.25
                    )
                    
                    if score > best_score and tax_diff <= max_diff and score >= 0.65:
                        best_score = score
                        best_match = (row_2b, row_pr, score)
                
                if best_match and best_score >= 0.65:
                    matches.append({
                        'row_2b': best_match[0],
                        'row_pr': best_match[1],
                        'confidence': best_match[2]
                    })
        
        if not matches:
            return None
        
        merged = pd.DataFrame([{
            **{f"{k}_2B": v for k, v in m['row_2b'].to_dict().items()},
            **{f"{k}_PR": v for k, v in m['row_pr'].to_dict().items()},
            'CONFIDENCE': m['confidence']
        } for m in matches])
        
        return {
            'dataframe': merged,
            'matched_2b': set(merged['orig_idx_2B']),
            'matched_pr': set(merged['orig_idx_PR']),
            'strategy': MatchStrategy.AI_ENHANCED.value,
            'tier': MatchTier.TIER_4_FUZZY.value,
            'confidence': merged['CONFIDENCE'].mean()
        }
    
    # ========================================================================
    # RESULTS ASSEMBLY - STATUS BASED ON TIER
    # ========================================================================
    
    def _assemble_results(self, matches: List[pd.DataFrame],
                         unmatched_2b: pd.DataFrame,
                         unmatched_pr: pd.DataFrame) -> pd.DataFrame:
        """Assemble final results with credit note information and unified columns.
           Status depends on MATCH_TIER to avoid false 'Exact'.
        """
        result_dfs = []

        base_cols = ['SUPPLIER GSTIN', 'SUPPLIER NAME', 'DOCUMENT NUMBER', 'DOCUMENT DATE',
                     'TAXABLE VALUE', 'IGST', 'CGST', 'SGST', 'TOTAL TAX', 'TOTAL VALUE',
                     'DOC TYPE', 'REFERENCE DOCUMENT', 'IS_CREDIT', 'DOC_TYPE_NORM']

        def add_suffix(df, suffix, cols=None):
            if cols is None:
                cols = [c for c in df.columns if c not in ['orig_idx', 'orig_idx_2B', 'orig_idx_PR', 'MATCH_STATUS', 'MATCH_REASON', 'MATCH_TIER', 'CONFIDENCE', 'RISK_LEVEL', 'IS_CREDIT', 'DOC_TYPE_DISPLAY']]
            rename_dict = {c: f"{c}_{suffix}" for c in cols if c in df.columns}
            return df.rename(columns=rename_dict)

        if matches:
            main_df = pd.concat(matches, ignore_index=True, sort=False)

            if 'MATCH_TIER' not in main_df.columns:
                main_df['MATCH_TIER'] = 2

            def assign_status(row):
                tier = row.get('MATCH_TIER', 2)
                conf = row.get('CONFIDENCE', 0)
                if tier == MatchTier.TIER_1_EXACT.value and conf >= 0.99:
                    return 'Exact'
                elif conf >= 0.8:
                    return 'Suggested'
                elif conf >= 0.5:
                    return 'Partial'
                else:
                    return 'Partial'

            main_df['MATCH_STATUS'] = main_df.apply(assign_status, axis=1)
            main_df['MATCH_REASON'] = "Multi-strategy matching successful"

            if 'IS_CREDIT_2B' in main_df.columns:
                main_df['IS_CREDIT'] = main_df['IS_CREDIT_2B'].fillna(False).astype(bool)
            elif 'IS_CREDIT' in main_df.columns:
                main_df['IS_CREDIT'] = main_df['IS_CREDIT'].fillna(False).astype(bool)
            else:
                main_df['IS_CREDIT'] = False

            if 'TAX_DIFF' in main_df.columns:
                is_not_credit = ~main_df['IS_CREDIT'].astype(bool)
                conditions = [
                    (main_df['TAX_DIFF'] > self.config.tolerance_amount * 3) & is_not_credit,
                    (main_df['TAX_DIFF'] > self.config.tolerance_amount * 2) & is_not_credit,
                    (main_df['TAX_DIFF'] > self.config.tolerance_amount) & is_not_credit
                ]
                choices = ['Critical', 'High', 'Medium']
                main_df['RISK_LEVEL'] = np.select(conditions, choices, default='Low')
            else:
                main_df['RISK_LEVEL'] = 'Low'

            main_df['DOC_TYPE_DISPLAY'] = np.where(
                main_df['IS_CREDIT'], 'CREDIT NOTE', 'INVOICE'
            )

            result_dfs.append(main_df)

        if not unmatched_2b.empty:
            u2b = unmatched_2b.copy()
            cols_2b = [c for c in u2b.columns if c not in ['orig_idx', 'MATCH_STATUS', 'MATCH_REASON', 'MATCH_TIER', 'CONFIDENCE', 'RISK_LEVEL', 'IS_CREDIT', 'DOC_TYPE_DISPLAY']]
            u2b = add_suffix(u2b, '2B', cols_2b)
            for col in base_cols:
                if col not in u2b.columns:
                    u2b[f"{col}_PR"] = ''
            for col in ['IGST', 'CGST', 'SGST', 'TOTAL TAX', 'TOTAL VALUE']:
                if f"{col}_PR" not in u2b.columns:
                    u2b[f"{col}_PR"] = 0.0
            u2b['MATCH_STATUS'] = 'Missing in PR'
            u2b['MATCH_REASON'] = 'Present in GSTR-2B but not found in Purchase Register'
            u2b['MATCH_TIER'] = 0
            u2b['CONFIDENCE'] = 0.0
            u2b['RISK_LEVEL'] = 'Critical'
            if 'IS_CREDIT' not in u2b.columns:
                u2b['IS_CREDIT'] = u2b.get('IS_CREDIT_2B', False)
            u2b['DOC_TYPE_DISPLAY'] = np.where(
                u2b['IS_CREDIT'].astype(bool), 'CREDIT NOTE', 'INVOICE'
            )
            result_dfs.append(u2b)

        if not unmatched_pr.empty:
            upr = unmatched_pr.copy()
            cols_pr = [c for c in upr.columns if c not in ['orig_idx', 'MATCH_STATUS', 'MATCH_REASON', 'MATCH_TIER', 'CONFIDENCE', 'RISK_LEVEL', 'IS_CREDIT', 'DOC_TYPE_DISPLAY']]
            upr = add_suffix(upr, 'PR', cols_pr)
            for col in base_cols:
                if col not in upr.columns:
                    upr[f"{col}_2B"] = ''
            for col in ['IGST', 'CGST', 'SGST', 'TOTAL TAX', 'TOTAL VALUE']:
                if f"{col}_2B" not in upr.columns:
                    upr[f"{col}_2B"] = 0.0
            upr['MATCH_STATUS'] = 'Missing in 2B'
            upr['MATCH_REASON'] = 'Present in Purchase Register but not found in GSTR-2B'
            upr['MATCH_TIER'] = 0
            upr['CONFIDENCE'] = 0.0
            upr['RISK_LEVEL'] = 'Critical'
            if 'IS_CREDIT' not in upr.columns:
                upr['IS_CREDIT'] = upr.get('IS_CREDIT_PR', False)
            upr['DOC_TYPE_DISPLAY'] = np.where(
                upr['IS_CREDIT'].astype(bool), 'CREDIT NOTE', 'INVOICE'
            )
            result_dfs.append(upr)

        final_df = pd.concat(result_dfs, ignore_index=True, sort=False) if result_dfs else pd.DataFrame()

        temp_cols = [col for col in final_df.columns if col.startswith('_') or 'EXACT_KEY' in col or 'SEQ_GROUP' in col or 'DOC_PATTERN' in col]
        final_df = final_df.drop(columns=[c for c in temp_cols if c in final_df.columns], errors='ignore')

        all_cols = set()
        for df in result_dfs:
            all_cols.update(df.columns)
        for col in all_cols:
            if col not in final_df.columns:
                final_df[col] = np.nan

        return final_df
    
    def _calculate_statistics(self, final_df: pd.DataFrame,
                             df_2b: pd.DataFrame,
                             df_pr: pd.DataFrame) -> Dict:
        """Calculate comprehensive reconciliation statistics"""
        stats = {
            'total_2b_records': len(df_2b),
            'total_pr_records': len(df_pr),
            'processed_records': len(final_df),
        }
        
        if final_df.empty:
            stats.update({
                'exact_matches': 0,
                'suggested_matches': 0,
                'partial_matches': 0,
                'missing_in_2b': 0,
                'missing_in_pr': 0,
                'match_rate': 0.0,
                'avg_confidence': 0.0,
                'taxable_difference': 0.0,
                'total_taxable_2b': 0.0,
                'total_taxable_pr': 0.0,
                'risk_breakdown': {'Low': 0, 'Medium': 0, 'High': 0, 'Critical': 0},
                'credit_notes': {
                    'total_2b': 0, 'total_pr': 0, 'matched': 0,
                    'unmatched_2b': 0, 'unmatched_pr': 0,
                    'total_amount_2b': 0, 'total_amount_pr': 0, 'difference': 0
                }
            })
            return stats
        
        status_counts = final_df['MATCH_STATUS'].value_counts().to_dict()
        stats.update({
            'exact_matches': status_counts.get('Exact', 0),
            'suggested_matches': status_counts.get('Suggested', 0),
            'partial_matches': status_counts.get('Partial', 0),
            'missing_in_2b': status_counts.get('Missing in 2B', 0),
            'missing_in_pr': status_counts.get('Missing in PR', 0),
        })
        
        if 'RISK_LEVEL' in final_df.columns:
            risk_counts = final_df['RISK_LEVEL'].value_counts().to_dict()
            stats['risk_breakdown'] = {
                'Low': risk_counts.get('Low', 0),
                'Medium': risk_counts.get('Medium', 0),
                'High': risk_counts.get('High', 0),
                'Critical': risk_counts.get('Critical', 0)
            }
        
        if 'TAXABLE VALUE_2B' in final_df.columns:
            stats['total_taxable_2b'] = float(self._safe_series_sum(final_df['TAXABLE VALUE_2B']))
        elif 'TAXABLE VALUE' in final_df.columns:
            stats['total_taxable_2b'] = float(self._safe_series_sum(final_df[final_df['MATCH_STATUS'] != 'Missing in PR']['TAXABLE VALUE']))
        else:
            stats['total_taxable_2b'] = 0.0
        
        if 'TAXABLE VALUE_PR' in final_df.columns:
            stats['total_taxable_pr'] = float(self._safe_series_sum(final_df['TAXABLE VALUE_PR']))
        elif 'TAXABLE VALUE' in final_df.columns:
            stats['total_taxable_pr'] = float(self._safe_series_sum(final_df[final_df['MATCH_STATUS'] != 'Missing in 2B']['TAXABLE VALUE']))
        else:
            stats['total_taxable_pr'] = 0.0
        
        stats['taxable_difference'] = float(stats.get('total_taxable_2b', 0) - stats.get('total_taxable_pr', 0))
        
        total_matched = stats.get('exact_matches', 0) + stats.get('suggested_matches', 0) + stats.get('partial_matches', 0)
        stats['match_rate'] = (
            total_matched / stats['processed_records'] * 100
            if stats['processed_records'] > 0 else 0
        )
        
        if 'CONFIDENCE' in final_df.columns:
            conf_series = pd.to_numeric(final_df['CONFIDENCE'], errors='coerce').fillna(0)
            stats['avg_confidence'] = float(conf_series.mean())
            stats['confidence_std'] = float(conf_series.std())
            stats['confidence_high'] = int((conf_series >= 0.8).sum())
            stats['confidence_medium'] = int(((conf_series >= 0.5) & (conf_series < 0.8)).sum())
            stats['confidence_low'] = int(((conf_series > 0) & (conf_series < 0.5)).sum())
        
        if 'MATCH_TIER' in final_df.columns:
            tier_counts = final_df[final_df['MATCH_TIER'] > 0]['MATCH_TIER'].value_counts().to_dict()
            stats['tier_distribution'] = tier_counts
        
        if 'IS_CREDIT' in final_df.columns:
            credit_df = final_df[final_df['IS_CREDIT'] == True]
            matched_credit = credit_df[credit_df['MATCH_STATUS'].isin(['Exact', 'Suggested', 'Partial'])]
            
            stats['credit_notes'] = {
                'total_2b': self.metrics.credit_notes_2b,
                'total_pr': self.metrics.credit_notes_pr,
                'matched': len(matched_credit),
                'unmatched_2b': self.metrics.unmatched_credit_notes_2b,
                'unmatched_pr': self.metrics.unmatched_credit_notes_pr,
                'total_amount_2b': float(self.metrics.total_credit_amount_2b),
                'total_amount_pr': float(self.metrics.total_credit_amount_pr),
                'difference': float(self.metrics.credit_note_difference)
            }
        else:
            stats['credit_notes'] = {
                'total_2b': 0, 'total_pr': 0, 'matched': 0,
                'unmatched_2b': 0, 'unmatched_pr': 0,
                'total_amount_2b': 0, 'total_amount_pr': 0, 'difference': 0
            }
        
        return stats
    
    # ========================================================================
    # UTILITY METHODS
    # ========================================================================
    
    @staticmethod
    def _calculate_date_diff(date1, date2) -> int:
        """Calculate absolute date difference in days"""
        if pd.isna(date1) or pd.isna(date2):
            return 9999
        try:
            if isinstance(date1, pd.Timestamp):
                date1 = date1.to_pydatetime()
            if isinstance(date2, pd.Timestamp):
                date2 = date2.to_pydatetime()
            return abs((date2 - date1).days)
        except:
            return 9999
    
    @staticmethod
    def _calculate_fuzzy_score(name1: str, name2: str) -> float:
        """Calculate fuzzy name similarity score"""
        if pd.isna(name1) or pd.isna(name2):
            return 0.0
        
        n1 = str(name1).upper().strip()
        n2 = str(name2).upper().strip()
        
        suffixes = [
            'PVT LTD', 'PVT. LTD.', 'PRIVATE LIMITED', 'LTD', 'LIMITED',
            'LLP', 'AND SONS', '& SONS', 'COMPANY', 'CO',
            'INC', 'INC.', 'CORPORATION', 'CORP', 'CORP.',
            'ENTERPRISES', 'ENTERPRISE', 'INDUSTRIES', 'INDUSTRY'
        ]
        for suffix in suffixes:
            n1 = re.sub(r'\b' + re.escape(suffix) + r'\b', '', n1).strip()
            n2 = re.sub(r'\b' + re.escape(suffix) + r'\b', '', n2).strip()
        
        n1 = re.sub(r'\s+', ' ', n1).strip()
        n2 = re.sub(r'\s+', ' ', n2).strip()
        
        return SequenceMatcher(None, n1, n2).ratio() * 100
    
    @staticmethod
    def _extract_doc_pattern(doc_num: str) -> Dict:
        """Extract pattern from document number"""
        if pd.isna(doc_num) or not doc_num:
            return {'prefix': '', 'number': '', 'suffix': '', 'length': 0}
        
        doc_str = str(doc_num).strip()
        numbers = re.findall(r'\d+', doc_str)
        letters = re.findall(r'[A-Za-z]+', doc_str)
        
        return {
            'prefix': letters[0] if letters else '',
            'number': numbers[0] if numbers else '',
            'suffix': letters[-1] if len(letters) > 1 else '',
            'length': len(doc_str),
            'num_count': len(numbers),
            'letter_count': len(letters)
        }
    
    @staticmethod
    def _compare_doc_patterns(pattern1: Dict, pattern2: Dict) -> float:
        """Compare two document number patterns"""
        if not pattern1 or not pattern2:
            return 0.0
        
        score = 0.0
        total_weight = 0.0
        
        if pattern1.get('prefix') and pattern2.get('prefix'):
            if pattern1['prefix'] == pattern2['prefix']:
                score += 0.3
            total_weight += 0.3
        
        if pattern1.get('number') and pattern2.get('number'):
            len_diff = abs(len(pattern1['number']) - len(pattern2['number']))
            if len_diff <= 2:
                score += 0.2 * (1.0 - len_diff / 3.0)
            total_weight += 0.2
        
        if pattern1.get('suffix') and pattern2.get('suffix'):
            if pattern1['suffix'] == pattern2['suffix']:
                score += 0.2
            total_weight += 0.2
        
        if pattern1.get('length') and pattern2.get('length'):
            length_ratio = min(pattern1['length'], pattern2['length']) / max(pattern1['length'], pattern2['length'])
            score += 0.3 * length_ratio
            total_weight += 0.3
        
        return score / total_weight if total_weight > 0 else 0.0
    
    @staticmethod
    def _wildcard_match(doc1: str, doc2: str) -> float:
        """Wildcard pattern matching for document numbers"""
        if pd.isna(doc1) or pd.isna(doc2):
            return 0.0
        
        doc1 = str(doc1).upper().strip()
        doc2 = str(doc2).upper().strip()
        
        seq1 = re.findall(r'[A-Z0-9]+', doc1)
        seq2 = re.findall(r'[A-Z0-9]+', doc2)
        
        matches = 0
        total = max(len(seq1), len(seq2))
        
        used_indices = set()
        for s1 in seq1:
            for idx, s2 in enumerate(seq2):
                if idx in used_indices:
                    continue
                if s1 == s2 or (len(s1) > 2 and len(s2) > 2 and (s1 in s2 or s2 in s1)):
                    matches += 1
                    used_indices.add(idx)
                    break
        
        return matches / total if total > 0 else 0.0
    
    @staticmethod
    def _compare_reference_documents(ref1: str, ref2: str) -> float:
        """Compare reference documents for credit notes"""
        if pd.isna(ref1) or pd.isna(ref2):
            return 0.0
        
        ref1 = str(ref1).upper().strip()
        ref2 = str(ref2).upper().strip()
        
        if not ref1 or not ref2:
            return 0.0
        
        ref1 = re.sub(r'[^A-Z0-9]', '', ref1)
        ref2 = re.sub(r'[^A-Z0-9]', '', ref2)
        
        if ref1 == ref2:
            return 1.0
        
        if len(ref1) > 3 and len(ref2) > 3:
            if ref1 in ref2 or ref2 in ref1:
                return 0.8
        
        similarity = SequenceMatcher(None, ref1, ref2).ratio()
        return similarity
    
    @staticmethod
    def _calculate_document_similarity(doc1: str, doc2: str) -> float:
        """Calculate document number similarity"""
        if pd.isna(doc1) or pd.isna(doc2):
            return 0.0
        
        doc1 = str(doc1).upper().strip()
        doc2 = str(doc2).upper().strip()
        
        return SequenceMatcher(None, doc1, doc2).ratio()

# ============================================================================
# UI COMPONENTS (unchanged from original)
# ============================================================================

class UIComponents:
    """Enterprise UI Components with Credit Note Support"""
    
    @staticmethod
    def render_global_css():
        """Inject a polished, consistent visual theme across the whole app"""
        st.markdown("""
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
            
            html, body, [class*="css"] { font-family: 'Inter', -apple-system, sans-serif; }
            
            .main .block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 1400px; }
            
            /* Metric cards */
            div[data-testid="stMetric"] {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 14px;
                padding: 1rem 1.1rem;
                box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04);
                transition: box-shadow 0.15s ease, transform 0.15s ease;
            }
            div[data-testid="stMetric"]:hover {
                box-shadow: 0 6px 16px rgba(15, 23, 42, 0.08);
                transform: translateY(-1px);
            }
            div[data-testid="stMetricLabel"] { font-weight: 600; color: #64748b; }
            div[data-testid="stMetricValue"] { font-weight: 800; color: #0f172a; }
            
            /* Buttons */
            .stButton > button, .stDownloadButton > button {
                border-radius: 10px;
                font-weight: 700;
                border: 1px solid #cbd5e1;
                transition: all 0.15s ease;
            }
            .stButton > button:hover, .stDownloadButton > button:hover {
                border-color: #1565c0;
                color: #1565c0;
                box-shadow: 0 4px 10px rgba(21, 101, 192, 0.15);
            }
            div[data-testid="stDownloadButton"] > button {
                background: linear-gradient(135deg, #1565c0, #0d47a1);
                color: white;
                border: none;
            }
            div[data-testid="stDownloadButton"] > button:hover {
                background: linear-gradient(135deg, #1976d2, #1565c0);
                color: white;
            }
            
            /* Tabs */
            .stTabs [data-baseweb="tab-list"] { gap: 4px; border-bottom: 2px solid #e2e8f0; }
            .stTabs [data-baseweb="tab"] {
                border-radius: 10px 10px 0 0;
                padding: 0.6rem 1.2rem;
                font-weight: 600;
                color: #64748b;
            }
            .stTabs [aria-selected="true"] {
                background: #eff6ff;
                color: #1565c0 !important;
                border-bottom: 3px solid #1565c0;
            }
            
            /* Expanders */
            details {
                border-radius: 12px !important;
                border: 1px solid #e2e8f0 !important;
                overflow: hidden;
            }
            summary { font-weight: 700 !important; }
            
            /* Dataframes */
            div[data-testid="stDataFrame"] {
                border-radius: 12px;
                overflow: hidden;
                border: 1px solid #e2e8f0;
            }
            
            /* Progress bar */
            div[data-testid="stProgress"] > div > div { background: linear-gradient(90deg, #1565c0, #10b981); }
            
            /* Sidebar */
            section[data-testid="stSidebar"] { background: #f8fafc; border-right: 1px solid #e2e8f0; }
            
            /* Alerts */
            div[data-testid="stAlert"] { border-radius: 12px; }
        </style>
        """, unsafe_allow_html=True)
    
    @staticmethod
    def render_header():
        """Render enterprise header with version info"""
        st.markdown(f"""
        <div style="
            background: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #0f172a 100%);
            padding: 2.5rem 2rem;
            border-radius: 20px;
            margin-bottom: 2rem;
            border: 1px solid #334155;
            box-shadow: 0 10px 30px rgba(0,0,0,0.3);
        ">
            <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap;">
                <div>
                    <div style="display: flex; align-items: center; gap: 1rem;">
                        <div style="font-size: 3rem;">🏢</div>
                        <div>
                            <h1 style="color: #f8fafc; margin: 0; font-size: 2.8rem; font-weight: 800; letter-spacing: -0.5px;">
                                GST Recon Pro
                            </h1>
                            <p style="color: #94a3b8; margin: 0.25rem 0 0 0; font-size: 1.2rem; font-weight: 300;">
                                Enterprise Multi-Strategy GST Reconciliation Engine
                            </p>
                        </div>
                    </div>
                </div>
                <div style="display: flex; gap: 0.75rem; align-items: center; flex-wrap: wrap; margin-top: 0.5rem;">
                    <span style="
                        background: linear-gradient(135deg, #10b981, #059669);
                        color: white;
                        padding: 0.35rem 1.2rem;
                        border-radius: 9999px;
                        font-size: 0.7rem;
                        font-weight: 700;
                        letter-spacing: 0.5px;
                        text-transform: uppercase;
                    ">● ENTERPRISE v{VERSION}</span>
                    <span style="
                        background: linear-gradient(135deg, #3b82f6, #2563eb);
                        color: white;
                        padding: 0.35rem 1.2rem;
                        border-radius: 9999px;
                        font-size: 0.7rem;
                        font-weight: 700;
                        letter-spacing: 0.5px;
                        text-transform: uppercase;
                    ">● CREDIT NOTE SUPPORT</span>
                    <span style="
                        background: linear-gradient(135deg, #8b5cf6, #7c3aed);
                        color: white;
                        padding: 0.35rem 1.2rem;
                        border-radius: 9999px;
                        font-size: 0.7rem;
                        font-weight: 700;
                        letter-spacing: 0.5px;
                        text-transform: uppercase;
                    ">● 11 STRATEGIES</span>
                    <span style="
                        background: linear-gradient(135deg, #f59e0b, #d97706);
                        color: white;
                        padding: 0.35rem 1.2rem;
                        border-radius: 9999px;
                        font-size: 0.7rem;
                        font-weight: 700;
                        letter-spacing: 0.5px;
                        text-transform: uppercase;
                    ">● SIDE-BY-SIDE EXPORT</span>
                </div>
            </div>
            <div style="
                margin-top: 1rem;
                padding-top: 1rem;
                border-top: 1px solid #334155;
                display: flex;
                gap: 2rem;
                flex-wrap: wrap;
                color: #94a3b8;
                font-size: 0.85rem;
            ">
                <span>👨‍💻 <strong>Author:</strong> Abhishek Jakkula</span>
                <span>📧 <strong>Email:</strong> jakkulaabhishek5@gmail.com</span>
                <span>⚡ <strong>Status:</strong> <span style="color: #10b981;">● Operational</span></span>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    @staticmethod
    def render_sidebar():
        """Render enterprise sidebar with credit note sample data"""
        with st.sidebar:
            st.markdown(f"""
            <div style="
                text-align: center;
                padding: 1.5rem 0 1rem 0;
                border-bottom: 2px solid #e2e8f0;
                margin-bottom: 1.5rem;
            ">
                <div style="font-size: 3.5rem; margin-bottom: 0.25rem;">🏢</div>
                <h3 style="margin: 0; color: #1e293b; font-weight: 800;">GST Recon Pro</h3>
                <p style="margin: 0; color: #64748b; font-size: 0.8rem; font-weight: 500;">
                    Enterprise Edition v{VERSION}
                </p>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown("### 📥 Download Sample Files")
            st.markdown("Download sample Excel files with credit notes:")
            
            col1, col2 = st.columns(2)
            with col1:
                sample_2b, sample_pr = SampleDataGenerator.create_sample_excel_files()
                
                st.download_button(
                    label="📄 GSTR-2B Sample",
                    data=sample_2b,
                    file_name="GSTR_2B_Sample.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            with col2:
                st.download_button(
                    label="📄 Purchase Register Sample",
                    data=sample_pr,
                    file_name="Purchase_Register_Sample.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            st.markdown("---")
            st.markdown("### 📂 Data Upload")
            
            use_sample = st.checkbox("Use Generated Sample Data", value=False)
            
            if use_sample:
                st.info("🔄 Sample data with credit notes will be generated")
                col1, col2 = st.columns(2)
                with col1:
                    sample_size = st.slider("Sample Size", 50, 500, 100, 50)
                with col2:
                    credit_ratio = st.slider("Credit Note Ratio", 0.05, 0.35, 0.15, 0.05)
                
                return {
                    'file_2b': None,
                    'file_pr': None,
                    'use_sample': True,
                    'sample_size': sample_size,
                    'credit_ratio': credit_ratio
                }
            
            return {
                'file_2b': st.file_uploader("📄 GSTR-2B File", type=['xlsx', 'xls', 'csv']),
                'file_pr': st.file_uploader("📄 Purchase Register", type=['xlsx', 'xls', 'csv']),
                'use_sample': False,
                'sample_size': 0,
                'credit_ratio': 0.15
            }
    
    @staticmethod
    def render_strategy_config():
        """Render strategy configuration with credit note options"""
        st.markdown("---")
        st.markdown("### 🎯 Strategy Configuration")
        
        with st.expander("Matching Strategies", expanded=True):
            st.markdown("**Enable matching strategies:**")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.checkbox("Exact Match", value=True, disabled=True, key='enable_exact')
                st.checkbox("Smart Match", value=True, key='enable_smart')
                st.checkbox("Value-Based Match", value=True, key='enable_value')
                st.checkbox("Fuzzy Name Match", value=True, key='enable_fuzzy')
            
            with col2:
                st.checkbox("Pattern Recognition", value=True, key='enable_pattern')
                st.checkbox("Sequential Match", value=True, key='enable_sequential')
                st.checkbox("Aggregate Match", value=True, key='enable_aggregate')
                st.checkbox("Percentage Match", value=True, key='enable_percentage')
            
            with col3:
                st.checkbox("Wildcard Match", value=True, key='enable_wildcard')
                st.checkbox("AI-Enhanced Match", value=False, key='enable_ai')
                st.checkbox("Credit Note Match", value=True, key='enable_credit')
                st.checkbox("Negative Value Match", value=True, key='enable_negative')
        
        with st.expander("⚙️ Parameters"):
            col1, col2 = st.columns(2)
            with col1:
                tolerance = st.number_input("Value Tolerance (₹)", min_value=0.0, max_value=100000.0, value=20.0, step=5.0)
                date_tolerance = st.number_input("Date Tolerance (Days)", min_value=0, max_value=365, value=7, step=1)
            
            with col2:
                fuzzy_threshold = st.slider("Fuzzy Match Threshold", min_value=50, max_value=100, value=85, step=5)
                percentage_tolerance = st.slider("Percentage Tolerance (%)", min_value=1, max_value=20, value=5, step=1)
            
            min_confidence = st.slider("Minimum Confidence for Match", min_value=0.0, max_value=1.0, value=0.5, step=0.05, help="Matches with confidence below this threshold will be treated as unmatched.")
        
        with st.expander("🔧 Advanced Settings"):
            st.markdown("**Processing Options:**")
            col1, col2 = st.columns(2)
            with col1:
                include_rc = st.checkbox("Include Reverse Charge", value=True, key='include_rc')
                auto_claim = st.checkbox("Auto-claim ITC for Exact Matches", value=True, key='auto_claim')
                validate_gstin = st.checkbox("Validate GSTIN Format", value=True, key='validate_gstin')
                treat_negative = st.checkbox("Treat Negative Values as Credit Notes", value=True, key='treat_negative')
            
            with col2:
                strict_fy = st.checkbox("Strict Financial Year Matching", value=False, key='strict_fy')
                parallel_processing = st.checkbox("Parallel Processing", value=True, key='parallel_processing')
                enable_caching = st.checkbox("Enable Caching", value=True, key='enable_caching')
                sep_credit = st.checkbox("Match Credit Notes Separately", value=True, key='sep_credit')
            
            st.markdown("**Performance:**")
            max_workers = st.number_input("Parallel Workers", min_value=1, max_value=8, value=4, step=1)
        
        return {
            'enable_exact': st.session_state.get('enable_exact', True),
            'enable_smart': st.session_state.get('enable_smart', True),
            'enable_value': st.session_state.get('enable_value', True),
            'enable_fuzzy': st.session_state.get('enable_fuzzy', True),
            'enable_pattern': st.session_state.get('enable_pattern', True),
            'enable_sequential': st.session_state.get('enable_sequential', True),
            'enable_aggregate': st.session_state.get('enable_aggregate', True),
            'enable_percentage': st.session_state.get('enable_percentage', True),
            'enable_wildcard': st.session_state.get('enable_wildcard', True),
            'enable_ai': st.session_state.get('enable_ai', False),
            'enable_credit': st.session_state.get('enable_credit', True),
            'enable_negative': st.session_state.get('enable_negative', True),
            'tolerance': tolerance,
            'date_tolerance': date_tolerance,
            'fuzzy_threshold': fuzzy_threshold,
            'percentage_tolerance': percentage_tolerance,
            'include_rc': include_rc,
            'auto_claim': auto_claim,
            'validate_gstin': validate_gstin,
            'strict_fy': strict_fy,
            'max_workers': max_workers,
            'parallel_processing': parallel_processing,
            'enable_caching': enable_caching,
            'treat_negative': treat_negative,
            'sep_credit': sep_credit,
            'min_confidence': min_confidence
        }

# ============================================================================
# MAIN APPLICATION (unchanged from original)
# ============================================================================

class ConfigManager:
    """Save/load ReconciliationConfig to/from JSON so users can reuse settings"""

    @staticmethod
    def to_dict(config: ReconciliationConfig) -> Dict[str, Any]:
        d = {}
        for field_name in config.__dataclass_fields__:
            d[field_name] = getattr(config, field_name)
        return d

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> ReconciliationConfig:
        valid_fields = set(ReconciliationConfig.__dataclass_fields__.keys())
        filtered = {k: v for k, v in d.items() if k in valid_fields}
        return ReconciliationConfig(**filtered)

    @staticmethod
    def save_to_json(config: ReconciliationConfig, path: str) -> None:
        with open(path, 'w') as f:
            json.dump(ConfigManager.to_dict(config), f, indent=2)

    @staticmethod
    def load_from_json(path: str) -> ReconciliationConfig:
        with open(path, 'r') as f:
            d = json.load(f)
        return ConfigManager.from_dict(d)

    @staticmethod
    def save_to_json_bytes(config: ReconciliationConfig) -> bytes:
        """For Streamlit download_button usage"""
        return json.dumps(ConfigManager.to_dict(config), indent=2).encode('utf-8')


class VendorSummaryAnalyzer:
    """Vendor-wise and month-wise breakdown of reconciliation results —
    helps prioritize which suppliers need follow-up first."""

    @staticmethod
    def vendor_summary(final_df: pd.DataFrame) -> pd.DataFrame:
        """
        Aggregate by supplier GSTIN: total records, match rate, and
        total ₹ value of unresolved (missing) items — sorted so the
        biggest problem vendors surface first.
        """
        if final_df.empty:
            return pd.DataFrame(columns=[
                'SUPPLIER GSTIN', 'SUPPLIER NAME', 'TOTAL RECORDS', 'MATCHED',
                'MISSING IN 2B', 'MISSING IN PR', 'MATCH RATE %', 'UNRESOLVED VALUE'
            ])

        df = final_df.copy()
        df['_GSTIN'] = df.get('SUPPLIER GSTIN_2B', pd.Series(dtype=str)).fillna('')
        mask_blank = df['_GSTIN'] == ''
        if 'SUPPLIER GSTIN_PR' in df.columns:
            df.loc[mask_blank, '_GSTIN'] = df.loc[mask_blank, 'SUPPLIER GSTIN_PR'].fillna('')

        df['_NAME'] = df.get('SUPPLIER NAME_2B', pd.Series(dtype=str)).fillna('')
        mask_blank_name = df['_NAME'] == ''
        if 'SUPPLIER NAME_PR' in df.columns:
            df.loc[mask_blank_name, '_NAME'] = df.loc[mask_blank_name, 'SUPPLIER NAME_PR'].fillna('')

        df['_UNRESOLVED_VALUE'] = 0.0
        missing_2b_mask = df['MATCH_STATUS'] == 'Missing in 2B'
        missing_pr_mask = df['MATCH_STATUS'] == 'Missing in PR'
        if 'TAXABLE VALUE_PR' in df.columns:
            df.loc[missing_2b_mask, '_UNRESOLVED_VALUE'] = pd.to_numeric(
                df.loc[missing_2b_mask, 'TAXABLE VALUE_PR'], errors='coerce'
            ).fillna(0).abs()
        if 'TAXABLE VALUE_2B' in df.columns:
            df.loc[missing_pr_mask, '_UNRESOLVED_VALUE'] = pd.to_numeric(
                df.loc[missing_pr_mask, 'TAXABLE VALUE_2B'], errors='coerce'
            ).fillna(0).abs()

        grouped = df.groupby('_GSTIN').agg(
            SUPPLIER_NAME=('_NAME', 'first'),
            TOTAL_RECORDS=('MATCH_STATUS', 'count'),
            MATCHED=('MATCH_STATUS', lambda s: s.isin(['Exact', 'Suggested', 'Partial']).sum()),
            MISSING_IN_2B=('MATCH_STATUS', lambda s: (s == 'Missing in 2B').sum()),
            MISSING_IN_PR=('MATCH_STATUS', lambda s: (s == 'Missing in PR').sum()),
            UNRESOLVED_VALUE=('_UNRESOLVED_VALUE', 'sum'),
        ).reset_index()

        grouped['MATCH_RATE_PCT'] = (grouped['MATCHED'] / grouped['TOTAL_RECORDS'] * 100).round(1)
        grouped = grouped.rename(columns={'_GSTIN': 'SUPPLIER GSTIN'})
        grouped = grouped.sort_values('UNRESOLVED_VALUE', ascending=False).reset_index(drop=True)

        grouped.columns = [
            'SUPPLIER GSTIN', 'SUPPLIER NAME', 'TOTAL RECORDS', 'MATCHED',
            'MISSING IN 2B', 'MISSING IN PR', 'UNRESOLVED VALUE', 'MATCH RATE %'
        ]
        return grouped[[
            'SUPPLIER GSTIN', 'SUPPLIER NAME', 'TOTAL RECORDS', 'MATCHED',
            'MISSING IN 2B', 'MISSING IN PR', 'MATCH RATE %', 'UNRESOLVED VALUE'
        ]]

    @staticmethod
    def month_summary(final_df: pd.DataFrame) -> pd.DataFrame:
        """Aggregate match rate and unresolved value by month/year"""
        if final_df.empty:
            return pd.DataFrame(columns=['PERIOD', 'TOTAL RECORDS', 'MATCHED', 'MATCH RATE %'])

        df = final_df.copy()
        date_col = None
        for candidate in ['DOCUMENT DATE_2B', 'DOCUMENT DATE_PR', 'DOCUMENT DATE']:
            if candidate in df.columns:
                date_col = candidate
                break
        if date_col is None:
            return pd.DataFrame(columns=['PERIOD', 'TOTAL RECORDS', 'MATCHED', 'MATCH RATE %'])

        parsed = df[date_col].apply(lambda x: parse_date(x) if not isinstance(x, datetime) else x)
        df['_PERIOD'] = parsed.apply(lambda d: d.strftime('%Y-%m') if pd.notna(d) and d else 'Unknown')

        grouped = df.groupby('_PERIOD').agg(
            TOTAL_RECORDS=('MATCH_STATUS', 'count'),
            MATCHED=('MATCH_STATUS', lambda s: s.isin(['Exact', 'Suggested', 'Partial']).sum()),
        ).reset_index()
        grouped['MATCH_RATE_PCT'] = (grouped['MATCHED'] / grouped['TOTAL_RECORDS'] * 100).round(1)
        grouped = grouped.rename(columns={'_PERIOD': 'PERIOD'})
        grouped.columns = ['PERIOD', 'TOTAL RECORDS', 'MATCHED', 'MATCH RATE %']
        return grouped.sort_values('PERIOD').reset_index(drop=True)


def run_headless(path_2b: str, path_pr: str, output_path: str,
                  config: Optional[ReconciliationConfig] = None,
                  config_json_path: Optional[str] = None,
                  run_quality_check: bool = True) -> Dict[str, Any]:
    """
    Run a full reconciliation without Streamlit — for CLI/automation/CI use.
    Returns the stats dict; writes the side-by-side Excel report to output_path.
    """
    logger = LoggerSetup().get_logger()

    def _read_any(path: str) -> pd.DataFrame:
        if path.lower().endswith('.csv'):
            return pd.read_csv(path)
        return pd.read_excel(path)

    logger.info(f"Reading 2B file: {path_2b}")
    df_2b = _read_any(path_2b)
    logger.info(f"Reading PR file: {path_pr}")
    df_pr = _read_any(path_pr)

    processor = DataProcessor()
    df_2b = processor.standardize_columns(df_2b)
    df_pr = processor.standardize_columns(df_pr)
    df_2b, df_pr = processor.normalize_column_mappings(df_2b, df_pr)
    df_2b = processor.detect_credit_notes(df_2b)
    df_pr = processor.detect_credit_notes(df_pr)

    if run_quality_check:
        q2b = DataQualityAnalyzer.analyze(df_2b, 'GSTR-2B')
        qpr = DataQualityAnalyzer.analyze(df_pr, 'Purchase Register')
        print(DataQualityAnalyzer.format_report_text(q2b))
        print()
        print(DataQualityAnalyzer.format_report_text(qpr))
        print()

    if config_json_path:
        config = ConfigManager.load_from_json(config_json_path)
    elif config is None:
        config = ReconciliationConfig()

    engine = GSTReconciliationEngine(config)
    final_df, stats = engine.reconcile(df_2b, df_pr)

    excel_engine = ExcelExportEngine()
    output = excel_engine.create_comparison_export(final_df, stats)
    with open(output_path, 'wb') as f:
        f.write(output.getvalue())

    logger.info(f"Reconciliation complete. Match rate: {stats['match_rate']:.1f}%. Report saved to {output_path}")
    print(f"\n✅ Done. {stats['processed_records']} records processed, "
          f"match rate {stats['match_rate']:.1f}%. Report: {output_path}")

    return stats


def _build_cli_parser() -> 'argparse.ArgumentParser':
    import argparse
    parser = argparse.ArgumentParser(
        description=f'GST Recon Pro v{VERSION} — headless/CLI mode'
    )
    parser.add_argument('--2b', dest='file_2b', required=True, help='Path to GSTR-2B file (xlsx/csv)')
    parser.add_argument('--pr', dest='file_pr', required=True, help='Path to Purchase Register file (xlsx/csv)')
    parser.add_argument('--out', dest='output', required=True, help='Path to write the output Excel report')
    parser.add_argument('--config', dest='config_json', default=None, help='Path to a saved config JSON file')
    parser.add_argument('--no-quality-check', dest='no_quality_check', action='store_true',
                         help='Skip the pre-flight data quality report')
    return parser


def main():
    """Enterprise GST Reconciliation Application with Credit Note Support"""
    
    st.set_page_config(
        page_title=f"GST Recon Pro Enterprise v{VERSION}",
        page_icon="🏢",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    logger = LoggerSetup().get_logger()
    
    ui = UIComponents()
    ui.render_global_css()
    ui.render_header()
    
    uploaded_files = ui.render_sidebar()
    config_params = ui.render_strategy_config()
    
    if uploaded_files.get('use_sample', False):
        st.info("🔍 Generating sample data with credit notes...")
        sample_size = uploaded_files.get('sample_size', 100)
        credit_ratio = uploaded_files.get('credit_ratio', 0.15)
        
        df_2b_sample = SampleDataGenerator.generate_gstr_2b_data(sample_size, credit_note_ratio=credit_ratio)
        df_pr_sample = SampleDataGenerator.generate_purchase_register_data(df_2b_sample, match_rate=0.85, credit_note_ratio=credit_ratio)
        
        cn_summary_2b = SampleDataGenerator.get_credit_note_summary(df_2b_sample)
        cn_summary_pr = SampleDataGenerator.get_credit_note_summary(df_pr_sample)
        
        st.success(f"✅ Generated {len(df_2b_sample)} GSTR-2B records and {len(df_pr_sample)} Purchase Register records")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("GSTR-2B Total", len(df_2b_sample))
        with col2:
            st.metric("Purchase Register Total", len(df_pr_sample))
        with col3:
            st.metric("Credit Notes (2B)", cn_summary_2b['total_credit_notes'])
        with col4:
            st.metric("Credit Notes (PR)", cn_summary_pr['total_credit_notes'])
        
        with st.expander("📊 Sample Data Preview with Credit Notes"):
            col1, col2 = st.columns(2)
            with col1:
                st.write("GSTR-2B Sample:")
                st.dataframe(df_2b_sample.head(10), use_container_width=True)
            with col2:
                st.write("Purchase Register Sample:")
                st.dataframe(df_pr_sample.head(10), use_container_width=True)
        
        if st.button("🚀 Run Reconciliation with Credit Note Support", use_container_width=True):
            process_reconciliation(df_2b_sample, df_pr_sample, config_params)
    
    elif uploaded_files['file_2b'] and uploaded_files['file_pr']:
        try:
            file_2b = uploaded_files['file_2b']
            file_pr = uploaded_files['file_pr']
            
            if file_2b.name.endswith('.csv'):
                df_2b = pd.read_csv(io.BytesIO(file_2b.getvalue()))
            else:
                df_2b = pd.read_excel(io.BytesIO(file_2b.getvalue()))
            
            if file_pr.name.endswith('.csv'):
                df_pr = pd.read_csv(io.BytesIO(file_pr.getvalue()))
            else:
                df_pr = pd.read_excel(io.BytesIO(file_pr.getvalue()))
            
            process_reconciliation(df_2b, df_pr, config_params)
        
        except Exception as e:
            st.error(f"❌ Error reading files: {str(e)}")
            with st.expander("🔍 Technical Details"):
                st.code(traceback.format_exc())
    
    else:
        display_welcome()

def process_reconciliation(df_2b: pd.DataFrame, df_pr: pd.DataFrame, config_params: Dict):
    """Process reconciliation with credit note support"""
    try:
        with st.spinner("🚀 Processing with Enterprise Multi-Strategy Engine (Credit Note Support)..."):
            progress_bar = st.progress(0, text="Initializing...")
            
            progress_bar.progress(10, text="Reading and standardizing data...")
            
            processor = DataProcessor()
            df_2b = processor.standardize_columns(df_2b)
            df_pr = processor.standardize_columns(df_pr)
            
            required_cols = ['SUPPLIER GSTIN', 'DOCUMENT NUMBER', 'TAXABLE VALUE', 
                           'SUPPLIER NAME', 'DOCUMENT DATE']
            
            missing_2b = processor.get_missing_columns(df_2b, required_cols)
            missing_pr = processor.get_missing_columns(df_pr, required_cols)
            
            if missing_2b or missing_pr:
                st.warning("⚠️ Some required columns are missing. Attempting to find alternatives...")
                df_2b, df_pr = processor.normalize_column_mappings(df_2b, df_pr)
                
                missing_2b = processor.get_missing_columns(df_2b, required_cols)
                missing_pr = processor.get_missing_columns(df_pr, required_cols)
                
                if missing_2b:
                    st.error(f"❌ GSTR-2B file missing required columns: {missing_2b}")
                    st.stop()
                
                if missing_pr:
                    st.error(f"❌ Purchase Register file missing required columns: {missing_pr}")
                    st.stop()
            
            df_2b = processor.detect_credit_notes(df_2b)
            df_pr = processor.detect_credit_notes(df_pr)
            
            config = ReconciliationConfig(
                tolerance_amount=config_params['tolerance'],
                date_tolerance_days=config_params['date_tolerance'],
                fuzzy_threshold=config_params['fuzzy_threshold'],
                enable_reverse_charge=config_params['include_rc'],
                enable_auto_claim=config_params['auto_claim'],
                enable_fuzzy_matching=config_params['enable_fuzzy'],
                enable_pattern_recognition=config_params['enable_pattern'],
                enable_sequential_matching=config_params['enable_sequential'],
                enable_aggregate_matching=config_params['enable_aggregate'],
                enable_percentage_matching=config_params['enable_percentage'],
                enable_wildcard_matching=config_params['enable_wildcard'],
                enable_ai_enhanced=config_params['enable_ai'],
                validate_gstin=config_params['validate_gstin'],
                strict_financial_year=config_params['strict_fy'],
                max_workers=config_params['max_workers'],
                percentage_tolerance=config_params['percentage_tolerance'],
                enable_parallel_processing=config_params.get('parallel_processing', True),
                enable_caching=config_params.get('enable_caching', True),
                enable_credit_note_matching=config_params.get('enable_credit', True),
                enable_negative_value_matching=config_params.get('enable_negative', True),
                match_credit_notes_separately=config_params.get('sep_credit', True),
                treat_negative_as_credit=config_params.get('treat_negative', True),
                min_confidence_for_match=config_params.get('min_confidence', 0.5)
            )
            
            progress_bar.progress(30, text="Initializing matching engine...")
            engine = GSTReconciliationEngine(config)
            
            progress_bar.progress(50, text="Executing multi-strategy matching with credit note support...")
            final_df, stats = engine.reconcile(df_2b, df_pr)
            
            progress_bar.progress(100, text="Complete!")
            time.sleep(0.5)
            progress_bar.empty()
            
            # Store in session state for export
            st.session_state['recon_data'] = {
                'final_df': final_df,
                'stats': stats
            }
            
            display_results(final_df, stats)
            
    except Exception as e:
        st.error(f"❌ Error during processing: {str(e)}")
        with st.expander("🔍 Technical Details"):
            st.exception(e)
            st.code(traceback.format_exc())

def _confidence_color(value) -> str:
    """Manual red->yellow->green gradient for confidence scores (0-1), no matplotlib needed."""
    try:
        v = float(value)
    except (TypeError, ValueError):
        return ''
    v = max(0.0, min(1.0, v))
    
    # Two-stage interpolation: red(#ef4444) -> yellow(#facc15) -> green(#22c55e)
    def lerp(a, b, t):
        return int(a + (b - a) * t)
    
    if v < 0.5:
        t = v / 0.5
        r, g, b = lerp(239, 250, t), lerp(68, 204, t), lerp(68, 21, t)
    else:
        t = (v - 0.5) / 0.5
        r, g, b = lerp(250, 34, t), lerp(204, 197, t), lerp(21, 94, t)
    
    text_color = '#0f172a' if v > 0.35 else '#ffffff'
    return f'background-color: rgb({r},{g},{b}); color: {text_color}; font-weight:600;'

def _status_badge_html(status: str) -> str:
    colors = {
        'Exact': ('#e8f5e9', '#1b5e20'),
        'Suggested': ('#e3f2fd', '#1565c0'),
        'Partial': ('#f3e5f5', '#6a1b9a'),
        'Missing in 2B': ('#ffebee', '#b71c1c'),
        'Missing in PR': ('#fff3e0', '#e65100'),
    }
    bg, fg = colors.get(status, ('#f1f5f9', '#475569'))
    return f'background-color:{bg}; color:{fg}; font-weight:700; border-radius:6px; padding:2px 6px;'

def display_results(final_df: pd.DataFrame, stats: Dict):
    """Display reconciliation results with credit note visualizations"""
    
    st.markdown("---")
    st.markdown("## 📊 Reconciliation Results")
    
    tab_overview, tab_charts, tab_data, tab_export = st.tabs(
        ["📈 Overview", "🎯 Match Analytics", "📄 Data Preview", "💾 Export"]
    )
    
    status_counts = final_df['MATCH_STATUS'].value_counts().reset_index()
    status_counts.columns = ['Status', 'Count']
    status_colors = {
        'Exact': '#10b981',
        'Suggested': '#3b82f6',
        'Partial': '#8b5cf6',
        'Missing in 2B': '#ef4444',
        'Missing in PR': '#f59e0b'
    }
    
    with tab_overview:
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            st.metric("Total Records", f"{stats['processed_records']:,}")
        with col2:
            st.metric("Match Rate", f"{stats['match_rate']:.1f}%")
        with col3:
            financial_impact = stats.get('taxable_difference', 0)
            st.metric("Taxable Difference", f"₹{abs(financial_impact):,.2f}")
        with col4:
            avg_confidence = stats.get('avg_confidence', 0)
            st.metric("Avg Confidence", f"{avg_confidence:.1%}")
        with col5:
            processing_time = stats.get('processing_time', 0)
            st.metric("Processing Time", f"{processing_time:.2f}s")
        
        if 'credit_notes' in stats:
            cn_stats = stats['credit_notes']
            st.markdown("### 📋 Credit Note Summary")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Credit Notes (2B)", f"{cn_stats['total_2b']:,}")
            with col2:
                st.metric("Credit Notes (PR)", f"{cn_stats['total_pr']:,}")
            with col3:
                st.metric("Matched Credit Notes", f"{cn_stats['matched']:,}")
        
        st.markdown("### 📋 Status Breakdown")
        cols = st.columns(min(len(status_counts), 5))
        for idx, (_, row) in enumerate(status_counts.iterrows()):
            if idx < 5:
                with cols[idx]:
                    color = status_colors.get(row['Status'], '#64748b')
                    st.markdown(f"""
                    <div style="
                        background: {color}15;
                        border: 1px solid {color}40;
                        border-radius: 12px;
                        padding: 1rem;
                        text-align: center;
                    ">
                        <div style="font-size: 1.75rem; font-weight: 800; color: {color};">
                            {row['Count']:,}
                        </div>
                        <div style="color: #475569; font-weight: 600; font-size: 0.9rem;">{row['Status']}</div>
                    </div>
                    """, unsafe_allow_html=True)
        
        st.success(f"✅ Processing complete! {stats['processed_records']:,} records processed in {stats['processing_time']:.2f} seconds.")
    
    with tab_charts:
        chart_col1, chart_col2 = st.columns(2)
        
        with chart_col1:
            donut = go.Figure(data=[go.Pie(
                labels=status_counts['Status'],
                values=status_counts['Count'],
                hole=0.55,
                marker=dict(colors=[status_colors.get(s, '#64748b') for s in status_counts['Status']]),
                textinfo='label+percent',
                textfont=dict(size=12)
            )])
            donut.update_layout(
                title="Match Status Distribution",
                showlegend=True,
                height=380,
                margin=dict(t=50, b=10, l=10, r=10),
                font=dict(family='Inter, sans-serif')
            )
            st.plotly_chart(donut, use_container_width=True)
        
        with chart_col2:
            fin_labels = ['Taxable (2B)', 'Taxable (Books)']
            fin_values = [stats.get('total_taxable_2b', 0), stats.get('total_taxable_pr', 0)]
            fin_bar = go.Figure(data=[go.Bar(
                x=fin_labels, y=fin_values,
                marker_color=['#1565c0', '#6a1b9a'],
                text=[f"₹{v:,.0f}" for v in fin_values],
                textposition='outside'
            )])
            fin_bar.update_layout(
                title="Taxable Value: 2B vs Books",
                height=380,
                margin=dict(t=50, b=10, l=10, r=10),
                font=dict(family='Inter, sans-serif'),
                yaxis_title="₹"
            )
            st.plotly_chart(fin_bar, use_container_width=True)
        
        if 'strategy_breakdown' in stats and stats['strategy_breakdown']:
            strat_items = sorted(stats['strategy_breakdown'].items(), key=lambda x: x[1], reverse=True)[:12]
            strat_labels = [s.replace('_', ' ').title() for s, _ in strat_items]
            strat_values = [c for _, c in strat_items]
            
            strat_bar = go.Figure(data=[go.Bar(
                x=strat_values, y=strat_labels,
                orientation='h',
                marker_color='#8b5cf6',
                text=strat_values,
                textposition='outside'
            )])
            strat_bar.update_layout(
                title="Matches by Strategy (Matching Style Used)",
                height=max(320, 32 * len(strat_labels)),
                margin=dict(t=50, b=10, l=10, r=40),
                font=dict(family='Inter, sans-serif'),
                yaxis=dict(autorange='reversed')
            )
            st.plotly_chart(strat_bar, use_container_width=True)
        
        conf_col = None
        for cand in ['CONFIDENCE']:
            if cand in final_df.columns:
                conf_col = cand
        if conf_col:
            conf_hist = px.histogram(
                final_df, x=conf_col, nbins=20,
                color_discrete_sequence=['#1565c0'],
                title="Confidence Score Distribution"
            )
            conf_hist.update_layout(height=340, margin=dict(t=50, b=10, l=10, r=10), font=dict(family='Inter, sans-serif'))
            st.plotly_chart(conf_hist, use_container_width=True)
    
    with tab_data:
        st.markdown("### 📄 Data Preview")
        st.caption("Match status and confidence are color-coded for quick scanning. Showing first 100 rows.")
        
        preview_df = final_df.head(100).copy()
        style_cols = {}
        if 'MATCH_STATUS' in preview_df.columns:
            style_cols['MATCH_STATUS'] = lambda v: _status_badge_html(v)
        
        styler = preview_df.style
        if 'MATCH_STATUS' in preview_df.columns:
            styler = styler.map(lambda v: _status_badge_html(v), subset=['MATCH_STATUS'])
        if 'CONFIDENCE' in preview_df.columns:
            styler = styler.map(_confidence_color, subset=['CONFIDENCE'])
        
        st.dataframe(styler, use_container_width=True, hide_index=True, height=420)
    
    with tab_export:
        st.markdown("### 💾 Export Results")
        st.caption("Excel exports include native charts, conditional formatting, autofilters, and frozen headers.")
        
        col_export1, col_export2, col_export3 = st.columns(3)
        
        with col_export1:
            if st.button("📥 Export Side-by-Side Excel", use_container_width=True):
                export_side_by_side_excel(final_df, stats)
        
        with col_export2:
            if st.button("📄 Export CSV", use_container_width=True):
                csv = final_df.to_csv(index=False)
                st.download_button(
                    label="Download CSV",
                    data=csv,
                    file_name=f"GST_Recon_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
        
        with col_export3:
            if st.button("📊 Export Simple Excel", use_container_width=True):
                export_simple_excel(final_df, stats)

def export_side_by_side_excel(final_df: pd.DataFrame, stats: Dict):
    """Export enhanced side-by-side comparison Excel"""
    try:
        excel_engine = ExcelExportEngine()
        output = excel_engine.create_comparison_export(final_df, stats)
        
        st.download_button(
            label="📥 Download Side-by-Side Comparison Report",
            data=output.getvalue(),
            file_name=f"GST_Recon_Side_by_Side_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.success("✅ Enhanced side-by-side comparison report generated successfully!")
        st.info("📋 Report includes: Side-by-Side Comparison, Matched Records, Missing Records, Credit Notes Summary, and Dashboard")
    except Exception as e:
        st.error(f"❌ Error generating Excel report: {str(e)}")
        with st.expander("🔍 Technical Details"):
            st.code(traceback.format_exc())

def export_simple_excel(final_df: pd.DataFrame, stats: Dict):
    """Export simple Excel report"""
    try:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, sheet_name='Reconciliation', index=False)
            
            summary_data = [
                ['Metric', 'Value'],
                ['Total Records', stats['processed_records']],
                ['Match Rate', f"{stats['match_rate']:.1f}%"],
                ['Exact Matches', stats['exact_matches']],
                ['Suggested Matches', stats['suggested_matches']],
                ['Partial Matches', stats['partial_matches']],
                ['Missing in 2B', stats['missing_in_2b']],
                ['Missing in PR', stats['missing_in_pr']],
                ['Processing Time', f"{stats['processing_time']:.2f}s"]
            ]
            
            pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False, header=False)
        
        st.download_button(
            label="📥 Download Simple Excel Report",
            data=output.getvalue(),
            file_name=f"GST_Recon_Simple_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.success("✅ Simple Excel report generated successfully!")
    except Exception as e:
        st.error(f"❌ Error generating Excel report: {str(e)}")

def display_welcome():
    """Display welcome screen"""
    st.markdown("""
    <div style="
        text-align: center;
        padding: 4rem 2rem;
        background: linear-gradient(180deg, #f8fafc 0%, #ffffff 100%);
        border-radius: 20px;
        border: 1px solid #e2e8f0;
    ">
        <div style="font-size: 5rem; margin-bottom: 1.5rem;">🏢</div>
        <h2 style="
            color: #1e293b;
            font-size: 2.8rem;
            font-weight: 800;
            letter-spacing: -0.5px;
            margin: 0;
        ">
            Welcome to GST Recon Pro
        </h2>
        <p style="
            color: #64748b;
            font-size: 1.2rem;
            max-width: 600px;
            margin: 1rem auto;
            line-height: 1.6;
        ">
            Upload your GSTR-2B and Purchase Register files to start intelligent reconciliation 
            with 11 advanced matching strategies including full credit note support.
        </p>
        <div style="
            display: flex;
            gap: 1rem;
            justify-content: center;
            margin-top: 2rem;
            flex-wrap: wrap;
        ">
            <span style="
                background: #e8f5e9;
                color: #2e7d32;
                padding: 0.5rem 1rem;
                border-radius: 20px;
                font-size: 0.85rem;
            ">✅ Side-by-Side Excel Export</span>
            <span style="
                background: #e3f2fd;
                color: #1565c0;
                padding: 0.5rem 1rem;
                border-radius: 20px;
                font-size: 0.85rem;
            ">✅ Formula-Based Calculations</span>
            <span style="
                background: #f3e5f5;
                color: #6a1b9a;
                padding: 0.5rem 1rem;
                border-radius: 20px;
                font-size: 0.85rem;
            ">✅ Colorful Formatting</span>
            <span style="
                background: #fce4ec;
                color: #c62828;
                padding: 0.5rem 1rem;
                border-radius: 20px;
                font-size: 0.85rem;
            ">✅ Credit Note Support</span>
        </div>
        <p style="margin-top: 2rem;">💡 Download sample files from sidebar or use generated data for testing</p>
    </div>
    """, unsafe_allow_html=True)

# ============================================================================
# APPLICATION ENTRY POINT
# ============================================================================
if __name__ == "__main__":
    # If invoked with CLI flags (e.g. `python gst_recon_pro.py --2b a.xlsx --pr b.xlsx --out r.xlsx`),
    # run in headless mode instead of launching the Streamlit UI. This lets the
    # engine be used in scripts/CI without `streamlit run`.
    if any(arg.startswith('--2b') or arg.startswith('--pr') for arg in sys.argv[1:]):
        cli_parser = _build_cli_parser()
        cli_args = cli_parser.parse_args()
        cfg = None
        run_headless(
            path_2b=cli_args.file_2b,
            path_pr=cli_args.file_pr,
            output_path=cli_args.output,
            config_json_path=cli_args.config_json,
            run_quality_check=not cli_args.no_quality_check
        )
    else:
        try:
            main()
        except Exception as e:
            st.error(f"❌ Application Error: {str(e)}")
            st.code(traceback.format_exc())
