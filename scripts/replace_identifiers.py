# Python 3 script
# Usage: python3 replace_identifiers.py --root project_exports --old "_SetShapeBackgroundColor" --new "SetShapeBackgroundColor"
import argparse
import re
import os
from pathlib import Path

def replace_in_line(line, old, new):
    # Process a single VB line, skipping content inside quotes and after a comment apostrophe (')
    res = []
    i = 0
    n = len(line)
    in_string = False
    while i < n:
        ch = line[i]
        if ch == '"' :
            # entering or leaving string. VB escapes quotes as ""
            res.append(ch)
            i += 1
            # consume pair "" inside string
            while i < n:
                if line[i] == '"' and i+1 < n and line[i+1] == '"':
                    # escaped quote
                    res.append('""')
                    i += 2
                    continue
                elif line[i] == '"':
                    res.append('"')
                    i += 1
                    break
                else:
                    res.append(line[i])
                    i += 1
            continue
        if ch == "'":
            # start of comment: append rest and break
            res.append(line[i:])
            break
        # non-string, non-comment char, build segment until next " or '
        seg_start = i
        while i < n and line[i] != '"' and line[i] != "'":
            i += 1
        segment = line[seg_start:i]
        # replace only whole-word tokens using regex \b
        # VB identifier boundary: letters/digits/underscore -> use \b which should work for ASCII
        segment = re.sub(r'\b' + re.escape(old) + r'\b', new, segment)
        res.append(segment)
    return ''.join(res)

def process_file(path, old, new):
    changed = False
    with open(path, 'r', encoding='utf-8', errors='replace') as f:
        lines = f.readlines()
    out_lines = []
    for line in lines:
        new_line = replace_in_line(line, old, new)
        if new_line != line:
            changed = True
        out_lines.append(new_line)
    if changed:
        # write back as utf-8 without BOM
        with open(path, 'w', encoding='utf-8', newline='\n') as f:
            f.writelines(out_lines)
    return changed

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--root', required=True, help='Root folder to scan (e.g. project_exports)')
    parser.add_argument('--old', required=True)
    parser.add_argument('--new', required=True)
    parser.add_argument('--exts', default='.bas,.cls,.frm', help='Comma separated extensions')
    args = parser.parse_args()
    exts = [e.strip() for e in args.exts.split(',')]
    root = Path(args.root)
    if not root.exists():
        print("Root path not found:", root)
        return
    changed_files = []
    for p in root.rglob('*'):
        if p.is_file() and any(str(p).lower().endswith(ext) for ext in exts):
            if process_file(p, args.old, args.new):
                changed_files.append(str(p))
                print("Updated:", p)
    print("Done. Files changed:", len(changed_files))
    for f in changed_files:
        print(" -", f)

if __name__ == '__main__':
    main()