# -*- coding: utf-8 -*-
import json

path = r'C:\Users\taka_\.gemini\antigravity-cli\brain\f526ce98-612a-433c-b09b-eee6eef6420b\.system_generated\logs\transcript.jsonl'
with open(path, 'r', encoding='utf-8') as f:
    lines = [json.loads(l) for l in f]

with open('f526_user_req.txt', 'w', encoding='utf-8') as out:
    for d in lines:
        if d.get('type') == 'USER_INPUT':
            out.write("USER_INPUT:\n" + d.get('content', '') + "\n\n")
