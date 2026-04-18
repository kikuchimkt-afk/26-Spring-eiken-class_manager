import urllib.request, json, sys
sys.stdout.reconfigure(encoding='utf-8')

data = json.loads(urllib.request.urlopen('https://script.google.com/macros/s/AKfycbz2lwmv77SBY8kDSwJK5wlawVUg0CAOKmQxTpX-GNLi5DjnXYRn2jbYqI-U2I1__ofC/exec').read().decode('utf-8'))
d1 = data['appData']['day1']['participants']

# Expected option format
def expected_option(year, num):
    return f"{year}年度 第{num}回"

for pid, v in d1.items():
    rp = v.get('rpContent', '')
    ap = v.get('apContent', '')
    rem = v.get('remarks', '')
    if rp or ap or rem:
        print(f"{pid}:")
        print(f"  rpContent: [{rp}]")
        print(f"  apContent: [{ap}]")
        print(f"  remarks:   [{rem}]")
        # Check if it matches expected format
        matched = False
        for y in range(2018, 2026):
            for n in range(1, 4):
                if rp == expected_option(y, n):
                    matched = True
                    print(f"  -> MATCHES: {expected_option(y, n)}")
        if not matched and rp:
            print(f"  -> NO MATCH (data may need re-entry)")
        print()
