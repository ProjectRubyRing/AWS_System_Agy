# -*- coding: utf-8 -*-
import json

path = r'C:\Users\taka_\.gemini\antigravity-cli\brain\f526ce98-612a-433c-b09b-eee6eef6420b\.system_generated\logs\transcript_full.jsonl'
with open(path, 'r', encoding='utf-8') as f:
    lines = [json.loads(l) for l in f]

for d in lines[-6:]:
    print(f"=== Step {d.get('step_index')}: {d.get('source')} / {d.get('type')} ===")
    if d.get('tool_calls'):
        for c in d['tool_calls']:
            print("CALL:", c.get('name'))
            print("ARGS:", c.get('args'))
    if d.get('content'):
        print("CONTENT:", d.get('content')[:500])
