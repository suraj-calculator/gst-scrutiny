#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
AUTOMATIC E-WAY BILL MERGER - SEPARATE INWARD & OUTWARD
"""

import os
import glob
import time
import pandas as pd
from datetime import datetime
import json
import shutil

CONFIG_FILE = "processed_files.json"
MERGE_FOLDER = "merge_ewb"
INWARD_FOLDER = "inward_eway bill"
OUTWARD_FOLDER = "outward_eway bill"

def load_processed_files():
    """Load list of already processed files"""
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            return json.load(f)
    return []

def save_processed_files(processed):
    """Save list of processed files"""
    with open(CONFIG_FILE, 'w') as f:
        json.dump(processed, f, indent=2)

def merge_inward_files():
    """Merge only inward e-way bill files"""
    print("\n📥 Merging INWARD E-Way Bills...")
    
    all_dataframes = []
    
    if os.path.exists(INWARD_FOLDER):
        # Look for both .xls and .xlsx files
        files = []
        for ext in ['*.xlsx', '*.xls']:
            files.extend(glob.glob(os.path.join(INWARD_FOLDER, ext)))
        
        for file_path in files:
            try:
                if file_path.endswith('.xls'):
                    # Try to read .xls file
                    try:
                        df = pd.read_excel(file_path, engine='xlrd')
                    except:
                        df = pd.read_excel(file_path)
                else:
                    df = pd.read_excel(file_path, engine='openpyxl')
                
                all_dataframes.append(df)
                print(f"  ✓ Loaded: {os.path.basename(file_path)}")
            except Exception as e:
                print(f"  ✗ Error: {os.path.basename(file_path)} - {str(e)[:50]}")
    
    if all_dataframes:
        merged_df = pd.concat(all_dataframes, ignore_index=True, sort=False)
        merged_df = merged_df.drop_duplicates()
        
        os.makedirs(MERGE_FOLDER, exist_ok=True)
        
        # Save inward merged file
        inward_output = os.path.join(MERGE_FOLDER, "inward_eway_bill_merged.xlsx")
        merged_df.to_excel(inward_output, index=False)
        
        # Save with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        inward_backup = os.path.join(MERGE_FOLDER, f"inward_eway_bill_{timestamp}.xlsx")
        merged_df.to_excel(inward_backup, index=False)
        
        print(f"\n✅ INWARD Merge Complete!")
        print(f"   Files merged: {len(all_dataframes)}")
        print(f"   Total rows: {len(merged_df)}")
        print(f"   Output: inward_eway_bill_merged.xlsx")
        return True
    else:
        print(f"  ⚠️  No inward files found")
        return False

def merge_outward_files():
    """Merge only outward e-way bill files"""
    print("\n📤 Merging OUTWARD E-Way Bills...")
    
    all_dataframes = []
    
    if os.path.exists(OUTWARD_FOLDER):
        # Look for both .xls and .xlsx files
        files = []
        for ext in ['*.xlsx', '*.xls']:
            files.extend(glob.glob(os.path.join(OUTWARD_FOLDER, ext)))
        
        for file_path in files:
            try:
                if file_path.endswith('.xls'):
                    # Try to read .xls file
                    try:
                        df = pd.read_excel(file_path, engine='xlrd')
                    except:
                        df = pd.read_excel(file_path)
                else:
                    df = pd.read_excel(file_path, engine='openpyxl')
                
                all_dataframes.append(df)
                print(f"  ✓ Loaded: {os.path.basename(file_path)}")
            except Exception as e:
                print(f"  ✗ Error: {os.path.basename(file_path)} - {str(e)[:50]}")
    
    if all_dataframes:
        merged_df = pd.concat(all_dataframes, ignore_index=True, sort=False)
        merged_df = merged_df.drop_duplicates()
        
        os.makedirs(MERGE_FOLDER, exist_ok=True)
        
        # Save outward merged file
        outward_output = os.path.join(MERGE_FOLDER, "outward_eway_bill_merged.xlsx")
        merged_df.to_excel(outward_output, index=False)
        
        # Save with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        outward_backup = os.path.join(MERGE_FOLDER, f"outward_eway_bill_{timestamp}.xlsx")
        merged_df.to_excel(outward_backup, index=False)
        
        print(f"\n✅ OUTWARD Merge Complete!")
        print(f"   Files merged: {len(all_dataframes)}")
        print(f"   Total rows: {len(merged_df)}")
        print(f"   Output: outward_eway_bill_merged.xlsx")
        return True
    else:
        print(f"  ⚠️  No outward files found")
        return False

def merge_all_files():
    """Merge inward and outward files separately"""
    print("\n" + "="*60)
    print("🔄 STARTING MERGE PROCESS")
    print("="*60)
    
    inward_success = merge_inward_files()
    outward_success = merge_outward_files()
    
    print("\n" + "="*60)
    if inward_success:
        print("✅ INWARD files merged successfully")
    else:
        print("⚠️  No INWARD files found")
    
    if outward_success:
        print("✅ OUTWARD files merged successfully")
    else:
        print("⚠️  No OUTWARD files found")
    print("="*60)
    
    return inward_success or outward_success

def monitor_and_merge():
    """Monitor folder and merge when new files appear"""
    print("="*60)
    print("AUTOMATIC E-WAY BILL MERGER - RUNNING")
    print("="*60)
    print("\n📁 Monitoring folders:")
    print(f"   📥 INWARD:  {INWARD_FOLDER}")
    print(f"   📤 OUTWARD: {OUTWARD_FOLDER}")
    print("\n💡 What will happen:")
    print("   - Files in 'inward_eway bill' → inward_eway_bill_merged.xlsx")
    print("   - Files in 'outward_eway bill' → outward_eway_bill_merged.xlsx")
    print("\n⚠️  Close this window to stop")
    print("="*60)
    
    processed_files = load_processed_files()
    last_merge_time = 0
    merge_interval = 10  # Wait 10 seconds after detecting new files
    
    while True:
        try:
            current_files = []
            for folder in [INWARD_FOLDER, OUTWARD_FOLDER]:
                if os.path.exists(folder):
                    for ext in ['*.xls', '*.xlsx']:
                        current_files.extend(glob.glob(os.path.join(folder, ext)))
            
            new_files = [f for f in current_files if f not in processed_files]
            
            if new_files and (time.time() - last_merge_time) >= merge_interval:
                print(f"\n{'='*60}")
                print(f"[{datetime.now().strftime('%H:%M:%S')}] New files detected!")
                
                # Separate inward and outward new files
                inward_new = [f for f in new_files if INWARD_FOLDER in f]
                outward_new = [f for f in new_files if OUTWARD_FOLDER in f]
                
                if inward_new:
                    print(f"\n📥 New INWARD files:")
                    for file in inward_new:
                        print(f"  • {os.path.basename(file)}")
                
                if outward_new:
                    print(f"\n📤 New OUTWARD files:")
                    for file in outward_new:
                        print(f"  • {os.path.basename(file)}")
                
                print(f"\n🔄 Merging...")
                merge_all_files()
                
                # Mark all as processed
                processed_files.extend(new_files)
                save_processed_files(processed_files)
                last_merge_time = time.time()
                print(f"✅ Auto-merge complete!")
            
            time.sleep(5)  # Check every 5 seconds
            
        except KeyboardInterrupt:
            print("\n\n👋 Monitoring stopped by user")
            break
        except Exception as e:
            print(f"Error: {e}")
            time.sleep(10)

def run_once():
    """Run merge once and exit"""
    print("="*60)
    print("ONE-TIME E-WAY BILL MERGER")
    print("="*60)
    merge_all_files()
    print("\n✅ Done!")
    input("Press Enter to exit...")

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1 and sys.argv[1] == "--once":
        run_once()
    else:
        monitor_and_merge()