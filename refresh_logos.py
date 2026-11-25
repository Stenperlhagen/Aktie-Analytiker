import ast, json, base64, io, pathlib, concurrent.futures
import requests
from PIL import Image
import yfinance as yf
from urllib.parse import urlparse
tree=ast.parse(pathlib.Path('app.py').read_text())
env={}
for node in tree.body:
    if isinstance(node,ast.Assign) and any(isinstance(t,ast.Name) and t.id in {'GROUPS','CAT','ADDITIONAL_COMPANIES','ALL_COMPANIES'} for t in node.targets):
        exec(compile(ast.Module(body=[node],type_ignores=[]),'app.py','exec'),env)
companies=env['ALL_COMPANIES']
path=pathlib.Path('company_logos.json')
logos=json.loads(path.read_text()) if path.exists() else {}
def fetch(item):
    name,symbol=item
    if symbol in logos: return symbol,logos[symbol]
    try:
        website=yf.Ticker(symbol).get_info().get('website','')
        host=urlparse(website if '://' in website else 'https://'+website).hostname
        if not host: return symbol,None
        for url in [f'https://www.google.com/s2/favicons?domain={host}&sz=128',f'https://icons.duckduckgo.com/ip3/{host}.ico',f'https://{host}/favicon.ico']:
            try:
                r=requests.get(url,timeout=8);r.raise_for_status()
                im=Image.open(io.BytesIO(r.content)).convert('RGBA')
                if min(im.size)<16: continue
                im.thumbnail((128,128)); out=io.BytesIO();im.save(out,format='PNG')
                return symbol,'data:image/png;base64,'+base64.b64encode(out.getvalue()).decode()
            except Exception: pass
    except Exception: pass
    return symbol,None
with concurrent.futures.ThreadPoolExecutor(max_workers=4) as pool:
    for symbol,logo in pool.map(fetch,companies.items()):
        if logo: logos[symbol]=logo
path.write_text(json.dumps(logos,sort_keys=True))
print(f'Logo coverage: {sum(s in logos for s in companies.values())}/{len(companies)}')
