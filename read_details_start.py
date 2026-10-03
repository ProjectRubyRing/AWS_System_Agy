# -*- coding: utf-8 -*-
import json

path = r'C:\Users\taka_\.gemini\antigravity-cli\brain\104e81ab-fef3-447c-97d8-a7422eebbd0c\.system_generated\logs\transcript.jsonl'
with open(path, 'r', encoding='utf-8') as f:
    lines = [json.loads(l) for l in f]

for i in range(min(25, len(lines))):
    d = lines[i]
    stype = d.get('type')
    src = d.get('source')
    sidx = d.get('step_index')
    content = d.get('content', '')
    calls = d.get('tool_calls', [])
    print(f"[{sidx}] {src}:{stype}")
    if calls:
        for c in calls:
            print(f"   CALL: {c['name']} -> {str(c.get('args'))[:150]}")
    if content and len(content) < 500:
        print(f"   CONTENT: {content.strip()[:300]}")
