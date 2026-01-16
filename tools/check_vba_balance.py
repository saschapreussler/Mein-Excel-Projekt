#!/usr/bin/env python3
"""
VBA Module Balance Checker
Verifies that all Sub/Function/Property declarations have matching End statements
"""
import os
import re
import sys
from pathlib import Path

def analyze_vba_file(filepath):
    """Analyze a VBA file for Sub/Function/Property balance"""
    try:
        with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
            content = f.read()
    except Exception as e:
        return {'error': str(e), 'filepath': filepath}
    
    # Count procedures
    sub_starts = len(re.findall(r'^\s*(Public|Private|Friend)?\s*Sub\s+\w+', content, re.MULTILINE | re.IGNORECASE))
    function_starts = len(re.findall(r'^\s*(Public|Private|Friend)?\s*Function\s+\w+', content, re.MULTILINE | re.IGNORECASE))
    property_starts = len(re.findall(r'^\s*(Public|Private|Friend)?\s*Property\s+(Get|Let|Set)\s+\w+', content, re.MULTILINE | re.IGNORECASE))
    
    end_subs = len(re.findall(r'^\s*End\s+Sub\s*$', content, re.MULTILINE | re.IGNORECASE))
    end_functions = len(re.findall(r'^\s*End\s+Function\s*$', content, re.MULTILINE | re.IGNORECASE))
    end_properties = len(re.findall(r'^\s*End\s+Property\s*$', content, re.MULTILINE | re.IGNORECASE))
    
    return {
        'filepath': filepath,
        'filename': os.path.basename(filepath),
        'sub_starts': sub_starts,
        'function_starts': function_starts,
        'property_starts': property_starts,
        'end_subs': end_subs,
        'end_functions': end_functions,
        'end_properties': end_properties,
    }

def main():
    # Find project_exports directory
    script_dir = Path(__file__).parent.parent
    project_exports = script_dir / 'project_exports'
    
    if not project_exports.exists():
        print(f"Error: project_exports directory not found at {project_exports}")
        sys.exit(1)
    
    results = []
    
    # Analyze all VBA files
    for ext in ['*.bas', '*.cls', '*.frm']:
        for file in sorted(project_exports.glob(ext)):
            results.append(analyze_vba_file(file))
    
    if not results:
        print("No VBA files found in project_exports/")
        sys.exit(1)
    
    print("VBA MODULE BALANCE CHECK")
    print("=" * 80)
    print()
    
    all_balanced = True
    errors = []
    
    for r in results:
        if 'error' in r:
            errors.append(r)
            continue
            
        total_starts = r['sub_starts'] + r['function_starts'] + r['property_starts']
        total_ends = r['end_subs'] + r['end_functions'] + r['end_properties']
        
        if total_starts == total_ends:
            status = "✓ BALANCED"
        else:
            status = "✗ IMBALANCED"
            all_balanced = False
        
        print(f"{r['filename']:40} {status}")
        
        if total_starts != total_ends:
            print(f"  Sub: {r['sub_starts']} starts, {r['end_subs']} ends")
            print(f"  Function: {r['function_starts']} starts, {r['end_functions']} ends")
            print(f"  Property: {r['property_starts']} starts, {r['end_properties']} ends")
            print(f"  Total: {total_starts} starts, {total_ends} ends (diff: {total_ends - total_starts})")
    
    print()
    
    if errors:
        print("ERRORS:")
        print("-" * 80)
        for e in errors:
            print(f"  {e['filename']}: {e['error']}")
        print()
    
    if all_balanced and not errors:
        print("✓ All modules are properly balanced!")
        print()
        sys.exit(0)
    else:
        print("✗ Some modules have issues. Please review and fix.")
        print()
        sys.exit(1)

if __name__ == '__main__':
    main()
