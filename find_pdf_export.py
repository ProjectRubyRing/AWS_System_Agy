# -*- coding: utf-8 -*-
import json

path = r'C:\Users\taka_\.gemini\antigravity-cli\brain\f526ce98-612a-433c-b09b-eee6eef6420b\.system_generated\logs\transcript_full.jsonl'
with open(path, 'r', encoding='utf-8') as f:
    for line in f:
        if 'ExportAsFixedFormat' in line or 'output_preview.pdf' in line:
            d = json.loads(line)
            print(f"Step {d.get('step_index')}: {d.get('type')}")
            for c in d.get('tool_calls', []):
                print("  CALL:", c.get('name'))
                for k, v in c.get('args', {}).items():
                    if 'output_preview' in str(v) or 'ExportAsFixedFormat' in str(v):
                        print(f"    {k}: {str(v)[:300]}")
