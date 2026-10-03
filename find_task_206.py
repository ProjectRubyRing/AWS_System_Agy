# -*- coding: utf-8 -*-
import json

path = r'C:\Users\taka_\.gemini\antigravity-cli\brain\f526ce98-612a-433c-b09b-eee6eef6420b\.system_generated\logs\transcript_full.jsonl'
with open(path, 'r', encoding='utf-8') as f:
    for line in f:
        if 'task-206' in line:
            d = json.loads(line)
            print("FOUND LINE with task-206:")
            print(f"Step {d.get('step_index')}: {d.get('content')}")
            for c in d.get('tool_calls', []):
                print(c)
