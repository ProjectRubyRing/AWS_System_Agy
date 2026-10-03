# -*- coding: utf-8 -*-
import json

def inspect_session_details(cid, start=0, count=50):
    path = f'C:\\Users\\taka_\\.gemini\\antigravity-cli\\brain\\{cid}\\.system_generated\\logs\\transcript.jsonl'
    try:
        with open(path, 'r', encoding='utf-8') as f:
            lines = [json.loads(l) for l in f]
        print(f"Total lines: {len(lines)}")
        for i in range(start, min(start+count, len(lines))):
            d = lines[i]
            stype = d.get('type')
            calls = d.get('tool_calls', [])
            content = d.get('content', '')
            error = d.get('error', '')
            call_info = ""
            if calls:
                call_info = " -> " + ", ".join([f"{c['name']}({list(c.get('args', {}).keys())})" for c in calls])
                for c in calls:
                    if 'CommandLine' in c.get('args', {}):
                        call_info += f" cmd: {c['args']['CommandLine'][:80]}"
            print(f"[{d.get('step_index')}] {d.get('source')}:{stype}{call_info}")
            if error:
                print(f"    ERROR: {error[:100]}")
            if content and len(content) < 200 and ("error" in content.lower() or "exception" in content.lower() or "warning" in content.lower()):
                print(f"    CONTENT: {content.strip()}")
    except Exception as e:
        print(e)

print("=== Session 104e81ab (last 40 steps) ===")
inspect_session_details('104e81ab-fef3-447c-97d8-a7422eebbd0c', start=100, count=55)
