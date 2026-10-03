import json

p = r'C:\Users\taka_\.gemini\antigravity-cli\brain\f526ce98-612a-433c-b09b-eee6eef6420b\.system_generated\logs\transcript.jsonl'
with open(p, 'r', encoding='utf-8') as f:
    for line in f:
        data = json.loads(line)
        idx = data.get('step_index')
        if idx is not None and idx >= 200:
            print(f"Step {idx}: type={data.get('type')}, source={data.get('source')}")
            tc = data.get('tool_calls')
            if tc:
                for c in tc:
                    print("  Tool call:", c.get('name'), str(c.get('args'))[:200])
            content = data.get('content')
            if content and data.get('type') in ('SYSTEM_MESSAGE', 'USER_INPUT'):
                print("  Content:", content[:200].replace('\n', ' '))
