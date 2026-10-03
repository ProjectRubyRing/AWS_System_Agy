# -*- coding: utf-8 -*-
import json

path = r'C:\Users\taka_\.gemini\antigravity-cli\brain\104e81ab-fef3-447c-97d8-a7422eebbd0c\.system_generated\logs\transcript.jsonl'
with open(path, 'r', encoding='utf-8') as f:
    lines = [json.loads(l) for l in f]

with open('104e81ab_summary.txt', 'w', encoding='utf-8') as out:
    for d in lines:
        sidx = d.get('step_index')
        stype = d.get('type')
        thinking = d.get('thinking', '')
        content = d.get('content', '')
        calls = d.get('tool_calls', [])
        
        if thinking:
            out.write(f"=== STEP {sidx} THINKING ===\n{thinking}\n\n")
        if calls:
            for c in calls:
                out.write(f"=== STEP {sidx} CALL: {c.get('name')} ===\n{json.dumps(c.get('args'), ensure_ascii=False)[:300]}\n\n")
        if content and ("error" in content.lower() or "output" in content.lower() or "exception" in content.lower()):
            out.write(f"=== STEP {sidx} CONTENT ===\n{content[:500]}\n\n")
