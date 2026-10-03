# -*- coding: utf-8 -*-
import json

def read_last_steps(cid, n=15):
    path = f'C:\\Users\\taka_\\.gemini\\antigravity-cli\\brain\\{cid}\\.system_generated\\logs\\transcript.jsonl'
    try:
        with open(path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        print(f'=== Session {cid} (total lines: {len(lines)}) ===')
        for l in lines[-n:]:
            data = json.loads(l)
            src = data.get('source')
            stype = data.get('type')
            sidx = data.get('step_index')
            content = data.get('content')
            calls = data.get('tool_calls')
            if content:
                print(f'[{sidx}:{src}:{stype}] {content[:150]}...')
            if calls:
                print(f'[{sidx}:CALLS] {[c.get("name") for c in calls]}')
    except Exception as e:
        print(e)

read_last_steps('f526ce98-612a-433c-b09b-eee6eef6420b', 15)
print('-----------------------------------------')
read_last_steps('104e81ab-fef3-447c-97d8-a7422eebbd0c', 20)
