#!/usr/bin/env python3
import re
import sys

def replace_debug_with_log(file_path):
    """Replace Debug.WriteLine with LogHelper.Log in a C# file"""
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Replace Debug.WriteLine( with LogHelper.Log(
    modified = re.sub(r'Debug\.WriteLine\(', 'LogHelper.Log(', content)
    
    # Count changes
    original_count = content.count('Debug.WriteLine(')
    new_count = modified.count('LogHelper.Log(')
    changes = new_count - content.count('LogHelper.Log(')
    
    if changes > 0:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(modified)
        print(f"✓ {file_path}: Replaced {changes} occurrences")
        return changes
    else:
        print(f"- {file_path}: No changes needed")
        return 0

if __name__ == "__main__":
    files = sys.argv[1:]
    total = 0
    for file_path in files:
        total += replace_debug_with_log(file_path)
    print(f"\nTotal: {total} replacements")
