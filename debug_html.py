import requests
from bs4 import BeautifulSoup

url = "https://kodepos.nomor.net/_kodepos.php?_i=cari-kodepos&jobs=36.03.22.2011"
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}

print(f"Fetching: {url}\n")
response = requests.get(url, headers=headers)

print(f"Status: {response.status_code}\n")

soup = BeautifulSoup(response.content, 'html.parser')

tables = soup.find_all('table')
print(f"Total tables found: {len(tables)}\n")

for idx, table in enumerate(tables[:3]):
    print(f"=== TABLE {idx + 1} ===")
    rows = table.find_all('tr')
    print(f"Total rows: {len(rows)}\n")
    
    for i, row in enumerate(rows[:5]):
        cols = row.find_all(['td', 'th'])
        print(f"Row {i}: {len(cols)} columns")
        for j, col in enumerate(cols):
            text = col.get_text(strip=True)[:50]
            print(f"  Col {j}: {text}")
        print()
    
    print()
