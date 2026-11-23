import io, re, datetime as dt, urllib.parse
import streamlit as st
import pandas as pd
import plotly.express as px
import requests, feedparser, yfinance as yf
from pypdf import PdfReader
from docx import Document

st.set_page_config(page_title='EquityLab | Svenska aktier', page_icon='📈', layout='wide')
st.markdown('''<style>.block-container{max-width:1300px;padding-top:2rem}div[data-testid="stMetric"]{background:#f3f5f9;border:1px solid #e3e8f0;border-radius:15px;padding:15px}h1,h2,h3{letter-spacing:-.03em}</style>''', unsafe_allow_html=True)

# Swedish large-cap-oriented starter universe. Entries are not asserted to be an exhaustive, live Nasdaq classification.
GROUPS={
'Industri': {'ABB':'ABB.ST','Addtech B':'ADDT-B.ST','AFRY':'AFRY.ST','Alfa Laval':'ALFA.ST','Assa Abloy B':'ASSA-B.ST','Atlas Copco A':'ATCO-A.ST','Atlas Copco B':'ATCO-B.ST','Beijer Ref B':'BEIJ-B.ST','Bravida':'BRAV.ST','Bufab':'BUFAB.ST','Epiroc A':'EPI-A.ST','Indutrade':'INDT.ST','Lifco B':'LIFCO-B.ST','Nibe B':'NIBE-B.ST','Saab B':'SAAB-B.ST','Sandvik':'SAND.ST','Securitas B':'SECU-B.ST','Skanska B':'SKA-B.ST','Sweco B':'SWEC-B.ST','Trelleborg B':'TREL-B.ST','Volvo B':'VOLV-B.ST','Volvo A':'VOLV-A.ST','Munters':'MTRS.ST','Mycronic':'MYCR.ST','Loomis':'LOOMIS.ST','NCC B':'NCC-B.ST','Peab B':'PEAB-B.ST','SKF B':'SKF-B.ST','Systemair':'SYSR.ST','Lindab':'LIAB.ST','Alimak':'ALIG.ST','Clas Ohlson B':'CLAS-B.ST','AQ Group':'AQ.ST','Beijer Alma B':'BEIA-B.ST'},
'Bank och finans': {'Investor B':'INVE-B.ST','Investor A':'INVE-A.ST','Industrivärden C':'INDU-C.ST','SEB A':'SEB-A.ST','Handelsbanken A':'SHB-A.ST','Swedbank A':'SWED-A.ST','Nordea':'NDA-SE.ST','Avanza':'AZA.ST','Nordnet':'SAVE.ST','EQT':'EQT.ST','Latour B':'LATO-B.ST','Lundbergföretagen B':'LUND-B.ST','Kinnevik B':'KINV-B.ST','Bure Equity':'BURE.ST','Sampo':'SAMPO-SDB.ST','Storskogen B':'STOR-B.ST'},
'Fastigheter': {'Balder B':'BALD-B.ST','Castellum':'CAST.ST','Catena':'CATE.ST','Fabege':'FABG.ST','Sagax B':'SAGA-B.ST','Wallenstam B':'WALL-B.ST','Wihlborgs':'WIHL.ST','Hufvudstaden A':'HUFV-A.ST','Atrium Ljungberg B':'ATRLJ-B.ST','Pandox B':'PNDX-B.ST','NP3 Fastigheter':'NP3.ST','Hemnet':'HEM.ST','Corem B':'CORE-B.ST'},
'Teknik och telekom': {'Ericsson B':'ERIC-B.ST','Hexagon B':'HEXA-B.ST','Tele2 B':'TEL2-B.ST','Telia':'TELIA.ST','Sinch':'SINCH.ST','Truecaller B':'TRUE-B.ST','Vitec B':'VIT-B.ST','Addnode B':'ANOD-B.ST','Lagercrantz B':'LAGR-B.ST','HMS Networks':'HMS.ST','Yubico':'YUBICO.ST'},
'Hälsovård': {'AstraZeneca':'AZN.ST','Getinge B':'GETI-B.ST','Sobi':'SOBI.ST','Elekta B':'EKTA-B.ST','Sectra B':'SECT-B.ST','BioArctic B':'BIOA-B.ST','Camurus':'CAMX.ST','Bonesupport':'BONEX.ST','Vitrolife':'VITR.ST','Medicover B':'MCOV-B.ST','AddLife B':'ALIF-B.ST'},
'Konsument och övrigt': {'H&M B':'HM-B.ST','Evolution':'EVO.ST','Essity B':'ESSITY-B.ST','Axfood':'AXFO.ST','AAK':'AAK.ST','Electrolux B':'ELUX-B.ST','Husqvarna B':'HUSQ-B.ST','Thule':'THULE.ST','Dometic':'DOM.ST','Scandic Hotels':'SHOT.ST','SkiStar B':'SKIS-B.ST','Betsson B':'BETS-B.ST','Embracer B':'EMBRAC-B.ST','Autoliv':'ALIV-SDB.ST','Bilia A':'BILI-A.ST','New Wave B':'NEWA-B.ST','Mips':'MIPS.ST'},
'Material och skog': {'Boliden':'BOL.ST','SSAB B':'SSAB-B.ST','SCA B':'SCA-B.ST','Holmen B':'HOLM-B.ST','Billerud':'BILL.ST','Stora Enso R':'STE-R.ST','Gränges':'GRNG.ST','Hexpol B':'HPOL-B.ST','Alleima':'ALLEI.ST','Lundin Mining':'LUMI.ST'},
'Övriga bolag': {'Volati AB':'VOLO.ST','Salix Group':'SALIX.ST','Röko':'ROKO.ST','Asker Healthcare':'ASKER.ST','Verisure':'VSURE.ST','Paradox Interactive':'PDX.ST'}
}
CAT={n:(t,g) for g,items in GROUPS.items() for n,t in items.items()}

def numeric(v):
    try:
        f=float(v)
        return f if pd.notna(f) else None
    except (ValueError,TypeError): return None

def show(v, suffix='', digits=1): return 'Saknas' if v is None else f'{v:,.{digits}f}'.replace(',',' ').replace('.',',')+suffix

def val(info,*keys):
    for k in keys:
        v=numeric(info.get(k))
        if v is not None:return v
    return None

# ROIC AI: valfri andra källa, kräver personlig API-nyckel i Streamlit Secrets.
# Använd aldrig hårdkodade API-nycklar eller påhittade finansdata.
def roic_key():
    try: return str(st.secrets.get('ROIC_API_KEY','')).strip()
    except Exception: return ''

@st.cache_data(ttl=86400,show_spinner=False)
def roic_company_data(ticker, api_key):
    """Identifiera instrument via ROIC:s sök-API utan manuella bolagsmappningar."""
    if not api_key: return {}, 'ROIC API-nyckel saknas.'
    root='https://api.roic.ai/v3.0.0'
    symbol=ticker.removesuffix('.ST')
    queries=list(dict.fromkeys([ticker,symbol,symbol.replace('-','.'),symbol.split('-')[0]]))
    headers={'Authorization':'Bearer '+api_key}
    matches=[]
    errors=[]
    try:
        for query in queries:
            response=requests.get(root+'/tickers/search',params={'query':query},headers=headers,timeout=12)
            if not response.ok:
                errors.append(str(response.status_code))
                if response.status_code in (401,403,429): break
                continue
            payload=response.json()
            if isinstance(payload,list): items=payload
            elif isinstance(payload,dict):
                items=next((payload[k] for k in ('data','results','tickers','items') if isinstance(payload.get(k),list)),[])
            else: items=[]
            matches.extend(m for m in items if isinstance(m,dict))
            if matches: break
        if not matches:
            detail=(' HTTP '+','.join(errors) if errors else '')
            return {},'ROIC sökning gav inga instrument för '+ticker+'.'+detail
        def parts(value):
            raw=str(value or '').upper().strip()
            exchange,_,code=raw.partition(':')
            if not code: exchange,code='',exchange
            code=code.removesuffix('.ST').replace('-','').replace('_','').replace('.','').replace(' ','')
            return exchange,code
        _,desired=parts(symbol)
        candidates=[]
        for m in matches:
            raw_symbol=m.get('symbol') or m.get('ticker') or m.get('code')
            exchange,code=parts(raw_symbol)
            if code!=desired: continue
            venue=str(m.get('exchange') or m.get('exchange_code') or m.get('mic') or exchange).upper()
            if venue and not any(tag in venue for tag in ('STO','XSTO','STOCKHOLM')):
                continue
            candidates.append(m)
        if not candidates:
            examples=[str(m.get('symbol') or m.get('ticker') or m.get('code') or '?')[:35] for m in matches[:5]]
            return {},'ROIC hittade instrument men ingen säker Stockholmsträff för '+ticker+'. Exempel: '+', '.join(examples)
        unique={str(m.get('symbol') or m.get('ticker') or m.get('code')):m for m in candidates}
        if len(unique)!=1:
            return {},'ROIC gav flera möjliga Stockholmsträffar för '+ticker+'; välj inte automatiskt.'
        selected=list(unique.values())
        match=selected[0]
        identifier=str(match.get('symbol') or match.get('ticker') or match.get('code') or '')
        if ':' not in identifier:
            return {'search_match':match},'ROIC hittade instrumentet men saknar börskvalificerad symbol.'
        def get_periods(path):
            resp=requests.get(root+path+'/'+urllib.parse.quote(identifier,safe=':'),params={'period_type':'annual','limit':3,'order':'desc'},headers=headers,timeout=12)
            resp.raise_for_status()
            body=resp.json()
            rows=body.get('data',[]) if isinstance(body,dict) else []
            return sorted([v for v in rows if isinstance(v,dict) and v.get('period_type')=='annual' and v.get('symbol')==identifier],key=lambda v:v.get('period_end_date',''),reverse=True)
        income=get_periods('/fundamental/income-statement')
        if not income: return {'search_match':match,'symbol':identifier},'ROIC identifierade aktien men årsrapporter saknas i API-svaret.'
        balance_rows=[]
        try:
            balance_rows=get_periods('/fundamental/balance-sheet')
        except (requests.RequestException,ValueError,TypeError):
            pass
        return {'search_match':match,'symbol':identifier,'income':income,'balance':balance_rows},None
    except requests.RequestException as exc:
        return {},'ROIC nätverksfel: '+type(exc).__name__
    except (ValueError,TypeError) as exc:
        return {},'ROIC svarade i oväntat format: '+type(exc).__name__

@st.cache_data(ttl=1800,show_spinner=False)
def fetch(ticker):
    """Hämta varje datatyp oberoende så ett fel inte slår ut hela analysen."""
    obj = yf.Ticker(ticker)
    errors = {}
    def attempt(label, getter, fallback):
        try:
            result = getter()
            if result is None:
                errors[label] = 'Datakällan lämnade inget svar.'
                return fallback
            return result
        except Exception as exc:
            errors[label] = f'{type(exc).__name__}: {str(exc)[:250]}'
            return fallback
    info = attempt('Bolagsprofil', lambda: obj.info, {})
    history = attempt('Kurshistorik', lambda: obj.history(period='5y', auto_adjust=False), pd.DataFrame())
    fin = attempt('Resultaträkning', lambda: obj.income_stmt, pd.DataFrame())
    cash = attempt('Kassaflöde', lambda: obj.cashflow, pd.DataFrame())
    balance = attempt('Balansräkning', lambda: obj.balance_sheet, pd.DataFrame())
    # Yahoo fast_info kan fungera även när den större bolagsprofilen blockeras.
    quick = attempt('Snabbkursdata', lambda: dict(obj.fast_info), {})
    return info or {}, history, fin, cash, balance, quick or {}, errors

@st.cache_data(ttl=3600,show_spinner=False)
def fx_rate(from_currency,to_currency):
    if not from_currency or not to_currency: return None
    if from_currency==to_currency: return 1.0
    try:
        pair=f"{from_currency}{to_currency}=X"
        frame=yf.Ticker(pair).history(period='5d',auto_adjust=False)
        if frame is None or frame.empty: return None
        rate=numeric(frame['Close'].dropna().iloc[-1])
        return rate if rate and rate>0 else None
    except Exception:
        return None

@st.cache_data(ttl=86400,show_spinner=False)
def trailing_dividend_yield(ticker, price):
    """Faktiska kontantutdelningar senaste 365 dagarna, i aktiens handelsvaluta."""
    if not isinstance(price,(int,float)) or price<=0: return None
    try:
        series=yf.Ticker(ticker).dividends
        if series is None or series.empty: return None
        dates=pd.to_datetime(series.index,utc=True)
        cutoff=pd.Timestamp.now(tz='UTC')-pd.Timedelta(days=365)
        paid=pd.to_numeric(series.loc[dates>=cutoff],errors='coerce')
        paid=paid[(paid>0)&paid.notna()]
        if paid.empty: return None
        value=float(paid.sum())/price*100
        return value if 0<value<=30 else None
    except Exception:
        return None

@st.cache_data(ttl=7200, show_spinner=False)
def extended_financials(ticker):
    """Ytterligare offentligt tillgängliga serier, hämtade separat och cachade."""
    obj = yf.Ticker(ticker)
    result, errors = {}, {}
    sources = {
        'quarterly_income': lambda: obj.quarterly_income_stmt,
        'quarterly_cash': lambda: obj.quarterly_cashflow,
        'quarterly_balance': lambda: obj.quarterly_balance_sheet,
        'valuation_history': lambda: obj.get_valuation_measures(freq='quarterly', periods=12),
    }
    for key, getter in sources.items():
        try:
            frame = getter()
            result[key] = frame if isinstance(frame, pd.DataFrame) else pd.DataFrame()
        except Exception as exc:
            result[key] = pd.DataFrame()
            errors[key] = f'{type(exc).__name__}: {str(exc)[:120]}'
    return result, errors

def assessment(fin, cash, balance, info, price, netdebt, ebitda, history, extended, quick=None):
    """Bedömning utan prognosantaganden eller uppfunna datapunkter."""
    def series(frame, keys):
        if frame is None or frame.empty: return pd.Series(dtype='float64')
        for key in keys:
            if key in frame.index:
                x = pd.to_numeric(frame.loc[key], errors='coerce').dropna()
                if len(x): return x.sort_index()
        return pd.Series(dtype='float64')
    def cagr(x):
        if len(x)<3 or x.iloc[0]<=0 or x.iloc[-1]<=0: return None
        return (float(x.iloc[-1])/float(x.iloc[0]))**(1/(len(x)-1))-1
    rev = series(fin,['Total Revenue','Operating Revenue'])
    ebit = series(fin,['Operating Income','EBIT'])
    fcf = series(cash,['Free Cash Flow'])
    ocf = series(cash,['Operating Cash Flow','Cash Flow From Continuing Operating Activities'])
    capex = series(cash,['Capital Expenditure'])
    debt = series(balance,['Total Debt'])
    cash_bal = series(balance,['Cash And Cash Equivalents','Cash Cash Equivalents And Short Term Investments'])
    net_income = series(fin,['Net Income','Net Income Common Stockholders'])
    equity = series(balance,['Stockholders Equity','Common Stock Equity'])
    sector = str(info.get('sector') or '').lower()
    industry = str(info.get('industry') or '').lower()
    specialized = any(x in sector+' '+industry for x in ['financial','bank','insurance','real estate','reit','investment management'])
    evidence = []
    def record(group, factor, metric, pass_test, note):
        if metric is not None and pd.notna(metric):
            evidence.append({'Område':group,'Faktor':factor,'Värde':round(float(metric),3),'Utfall':'Positivt' if pass_test else 'Negativt','Metod':note})
    sales_cagr = cagr(rev)
    record('Tillväxt','Omsättning CAGR',sales_cagr,sales_cagr is not None and sales_cagr>0.03,'Minst tre års observationer; 3 % är en metodgräns')
    net_cagr = cagr(net_income)
    record('Tillväxt','Nettovinst CAGR',net_cagr,net_cagr is not None and net_cagr>0,'Kräver positiva start- och slutvärden')
    margins = [float(ebit[d]/rev[d]) for d in rev.index.intersection(ebit.index) if rev[d]>0]
    if margins:
        record('Lönsamhet','Senaste EBIT-marginal',margins[-1],margins[-1]>0.08,'8 % är metodgräns, ej branschnorm')
    if len(margins)>=3:
        record('Lönsamhet','Marginalförändring',margins[-1]-margins[0],margins[-1]-margins[0]>0,'Senaste minus äldsta årsperiod')
    if len(ocf)>=3:
        record('Kassaflöde','Andel positiva kassaflödesår',float((ocf>0).mean()),float((ocf>0).mean())>=0.75,'Minst 75 % positiva år')
    if len(fcf)>=3:
        record('Kassaflöde','Andel positiva FCF-år',float((fcf>0).mean()),float((fcf>0).mean())>=0.75,'Minst 75 % positiva år')
    if not specialized and netdebt is not None and ebitda is not None and ebitda>0:
        record('Risk','Nettoskuld/EBITDA',netdebt/ebitda,netdebt/ebitda<2.5,'Metodgräns 2,5×, inte generell kreditrating')
    if not specialized and len(equity)>0 and len(net_income)>0:
        common=equity.index.intersection(net_income.index)
        if len(common)>0 and equity[common[-1]]>0:
            roe=float(net_income[common[-1]]/equity[common[-1]])
            record('Lönsamhet','ROE senaste år',roe,roe>0.10,'Metodgräns 10 %')
    quick = quick or {}
    shares = info.get('sharesOutstanding') or quick.get('shares')
    marketcap = info.get('marketCap') or quick.get('market_cap')
    if not isinstance(marketcap,(int,float)) and price and isinstance(shares,(int,float)):
        marketcap=price*shares
    if isinstance(marketcap,(int,float)) and marketcap>0 and len(fcf)>0:
        fy=float(fcf.iloc[-1]/marketcap)
        record('Värdering','FCF-yield senaste år',fy,fy>0.05,'5 % metodgräns; senaste års FCF kan vara cykliskt')
    pe = info.get('trailingPE')
    # Härledning från EPS är säker bara när EPS och kurs uttrycks i samma valuta.
    quote_currency = info.get('currency') or quick.get('currency')
    earnings_currency = info.get('financialCurrency')
    eps = info.get('trailingEps')
    if not isinstance(pe,(int,float)) and quote_currency and earnings_currency and quote_currency==earnings_currency and isinstance(eps,(int,float)) and eps>0 and price:
        pe=price/eps
    if isinstance(pe,(int,float)) and 0<pe<500:
        record('Värdering','P/E',pe,pe<20,'20× är en generell metodgräns, inte historisk jämförelse')
    vh=extended.get('valuation_history',pd.DataFrame())
    if not vh.empty:
        for row_name in ['TrailingPeRatio','TrailingPE','PeRatio','EnterpriseValueEbitda','EnterpriseValueEbitdaRatio']:
            if row_name in vh.index:
                v=pd.to_numeric(vh.loc[row_name],errors='coerce').dropna()
                dated=v.drop(labels=['Current'],errors='ignore')
                current=v.get('Current')
                if pd.notna(current) and len(dated)>=4:
                    median=float(dated.median())
                    if median>0:
                        record('Värdering',f'{row_name} relativt historisk median',float(current)/median,float(current)/median<0.95,'Minst fyra historiska observationer; lägre multipel är inte alltid undervärdering')
                break
    groups = {'Värdering':0.35,'Tillväxt':0.15,'Lönsamhet':0.15,'Kassaflöde':0.20,'Risk':0.15}
    present = {g:[x for x in evidence if x['Område']==g] for g in groups}
    # En fullständig värderingsbedömning kräver minst två värderingsmått.
    valuation_count=len(present['Värdering'])
    missing=[g for g,items in present.items() if not items]
    if specialized:
        return 'GRÅ – Särskild branschmodell krävs', '#64748B', evidence, None, 'Banker, finans- och fastighetsbolag kräver branschspecifika värderingsmått.'
    if len(evidence)<7 or valuation_count<2 or missing:
        valuation_detail = ('Värderingsdata kunde inte beräknas: kontrollera börsvärde, aktiekurs, EPS, rapportvaluta och kassaflöde.' if valuation_count==0 else f'Endast {valuation_count} verifierbart värderingsmått tillgängligt; minst två krävs.')
        return 'GRÅ – Värderingsdata saknas i anslutna källor', '#64748B', evidence, None, valuation_detail+' Övriga finansiella uppgifter kan ändå finnas i bolagets rapporter.'
    score=sum(groups[g]*sum(x['Utfall']=='Positivt' for x in items)/len(items) for g,items in present.items())
    # Negativt FCF och allvarlig belåning kan inte kompenseras av billig P/E.
    risk_flag=(netdebt is not None and ebitda is not None and ebitda>0 and netdebt/ebitda>=4)
    if score>=0.75 and not risk_flag: status,color='GRÖN – Attraktiv enligt modellen','#16803C'
    elif score<0.40 or risk_flag: status,color='RÖD – Förhöjd risk enligt modellen','#C53030'
    else: status,color='GUL – Blandad investeringsbild','#B7791F'
    return status,color,evidence,score,'Modellpoäng bygger på dokumenterade metodgränser, inte kalibrerad kursprognos.'

# Officiella rapportkällor – byggs ut bolag för bolag.
# Dessa är källor för fördjupad kontroll, inte automatiskt kvalitetssäkrade siffror.
OFFICIAL_REPORT_SOURCES = {
    'BETS-B.ST': {
        'name': 'Betsson AB',
        'report': 'https://www.betssonab.com/investors/reports-and-presentations/betsson-ab-interim-report-january-september-2026',
        'archive': 'https://www.betssonab.com/investors/reports-and-presentations',
        'period': 'Januari–september 2026'
    }
}

@st.cache_data(ttl=86400, show_spinner=False)
def official_report_excerpt(ticker):
    """Hämta text från officiell IR-sida, utan att gissa belopp eller enheter."""
    src = OFFICIAL_REPORT_SOURCES.get(ticker)
    if not src:
        return [], None
    try:
        from bs4 import BeautifulSoup
        response = requests.get(src['report'],timeout=12,headers={'User-Agent':'Mozilla/5.0'})
        response.raise_for_status()
        soup = BeautifulSoup(response.text,'html.parser')
        for tag in soup(['script','style','nav','footer','header']):
            tag.decompose()
        lines = [re.sub(r'\\s+',' ',x).strip() for x in soup.stripped_strings]
        terms = ('revenue','ebitda','operating income','net income','operating cash flow','net debt','organic increase')
        matched=[]
        for line in lines:
            if any(term in line.lower() for term in terms) and 25<=len(line)<=600 and line not in matched:
                matched.append(line)
        return matched[:16], None
    except Exception as exc:
        return [], f'{type(exc).__name__}: {str(exc)[:180]}'

# Faktagranskade svenska bolagsbeskrivningar med direktlänk till originalkällan.
VERIFIED_COMPANY_DESCRIPTIONS = {
    'BETS-B.ST': (
        'Betsson AB är en svensk börsnoterad koncern inom onlinespel. Genom sina operativa dotterbolag erbjuder koncernen kasino och sportspel på internationella marknader, bland annat under varumärkena Betsson, Betsafe och NordicBet. Verksamheten bedrivs huvudsakligen på egenutvecklade tekniska plattformar. Moderbolaget ansvarar för koncernstrategi, ägarstyrning och förvärv.',
        'https://www.betssonab.com/sv/om-oss'
    )
}

@st.cache_data(ttl=86400, show_spinner=False)
def company_description_fallback(company):
    """Hämta faktatext med källhänvisning från Wikipedia om bolaget kan matchas."""
    clean = re.sub(r'\\s+(?:AB|A|B)$','',company,flags=re.I).strip()
    def normalized(x):
        return re.sub(r'[^a-z0-9]','',x.casefold().replace('å','a').replace('ä','a').replace('ö','o'))
    wanted=normalized(clean)
    for lang in ('sv','en'):
        endpoint=f'https://{lang}.wikipedia.org/w/api.php'
        try:
            response=requests.get(endpoint,params={
                'action':'query','format':'json','generator':'search',
                'gsrsearch':clean,'gsrlimit':10,
                'prop':'extracts|info','exintro':1,'explaintext':1,
                'inprop':'url','exchars':1400
            },headers={'User-Agent':'EquityLab/1.0 (financial company profiles)'},timeout=8)
            response.raise_for_status()
            pages=response.json().get('query',{}).get('pages',{})
            for page in pages.values():
                title=page.get('title','')
                title_base=re.sub(r'\\s*\\([^)]*\\)$','',title).strip()
                title_normalized=normalized(title_base)
                if title_normalized not in (wanted,normalized(clean+' AB'),normalized(company)):
                    continue
                extract=re.sub(r'\\s+',' ',page.get('extract') or '').strip()
                if len(extract)>=65 and not any(word in extract.lower() for word in ['may refer to','kan syfta på','disambiguation']):
                    return extract[:1300],page.get('fullurl'),f'Wikipedia ({lang})'
        except Exception:
            continue
    return None,None,None

@st.cache_data(ttl=3600, show_spinner=False)
def news(company):
    """Breda nyhetssökningar med flera RSS-källor och verifierad datumfiltrering."""
    since = dt.datetime.now(dt.timezone.utc) - dt.timedelta(days=183)
    search_name = re.sub(r'\\s+[AB]$', '', company, flags=re.I).strip()
    searches = [
        ('Google News', 'https://news.google.com/rss/search?' + urllib.parse.urlencode({
            'q': f'"{search_name}"', 'hl': 'sv', 'gl': 'SE', 'ceid': 'SE:sv'})),
        ('Google News bred', 'https://news.google.com/rss/search?' + urllib.parse.urlencode({
            'q': f'{search_name} börs', 'hl': 'sv', 'gl': 'SE', 'ceid': 'SE:sv'})),
        ('Bing News', 'https://www.bing.com/news/search?' + urllib.parse.urlencode({
            'q': f'"{search_name}"', 'format': 'rss', 'setlang': 'sv'})),
    ]
    out, seen, diagnostics = [], set(), []
    for source, url in searches:
        try:
            response = requests.get(url, timeout=15, headers={'User-Agent': 'Mozilla/5.0'})
            response.raise_for_status()
            feed = feedparser.parse(response.content)
            diagnostics.append(f'{source}: HTTP {response.status_code}, {len(response.content)} byte, {len(feed.entries)} poster, typ {response.headers.get("content-type", "okänd")}')
            if feed.bozo and not feed.entries:
                diagnostics.append(f'{source} – parserfel: {str(feed.bozo_exception)[:180]}')
            for entry in feed.entries:
                published = entry.get('published_parsed') or entry.get('updated_parsed')
                if not published:
                    continue
                date = dt.datetime(*published[:6], tzinfo=dt.timezone.utc)
                if not since <= date <= dt.datetime.now(dt.timezone.utc):
                    continue
                title, link = entry.get('title', '').strip(), entry.get('link', '').strip()
                if not title or not link or link in seen:
                    continue
                seen.add(link)
                out.append({'Datum': date.date().isoformat(), 'Rubrik': title, 'Länk': link, 'Källa': source})
        except Exception as exc:
            diagnostics.append(f'{source} – fel: {type(exc).__name__}: {str(exc)[:180]}')
    out.sort(key=lambda x: x['Datum'], reverse=True)
    return out[:35], diagnostics

def parse_file(upload):
    raw=upload.getvalue(); name=upload.name.lower()
    if name.endswith('.pdf'):
        reader=PdfReader(io.BytesIO(raw))
        return '\n'.join(p.extract_text() or '' for p in reader.pages[:200])
    if name.endswith('.docx'):
        doc=Document(io.BytesIO(raw));return '\n'.join(p.text for p in doc.paragraphs)
    if name.endswith('.txt'):return raw.decode('utf-8',errors='replace')
    if name.endswith('.csv'):
        return pd.read_csv(io.BytesIO(raw)).head(1000).to_string(index=False)
    if name.endswith('.xlsx'):
        sheets=pd.read_excel(io.BytesIO(raw),sheet_name=None)
        return '\n'.join(f'BLAD {k}\n{v.head(1000).to_string(index=False)}' for k,v in sheets.items())
    return ''

st.title('EquityLab')
st.caption('Automatisk bolagsöversikt · Svenska börsbolag · Egna rapporter · Scenariovärdering')
st.info('Bolagslistan är ett brett urval och inte en verifierad fullständig Nasdaq-segmentlista. Data hämtas vid val av bolag via Yahoo Finance (inofficiell källa, kan vara fördröjd eller ofullständig). Saknade fält visas som ”Saknas” – inga värden hittas på. Källorna bör kontrolleras mot bolagets rapporter.')
# Kompletterande svenska bolag. Börssegment kan ändras vid Nasdaqs omklassificeringar.
# Listan är ett sökbart urval, inte en garanti för en fullständig aktuell segmentförteckning.
ADDITIONAL_COMPANIES = {
    'Bokusgruppen AB':'BOKUS.ST',
    'Attendo':'ATT.ST',
    'Humana':'HUM.ST',
    'Byggmax':'BMAX.ST',
    'Rusta':'RUSTA.ST',
    'RevolutionRace':'RVRC.ST',
    'Diös Fastigheter':'DIOS.ST',
    'Platzer Fastigheter B':'PLAZ-B.ST',
    'Stendörren Fastigheter B':'STEF-B.ST',
    'Swedish Logistic Property B':'SLP-B.ST',
    'AcadeMedia':'ACAD.ST','Addvise B':'ADDV-B.ST','Arjo B':'ARJO-B.ST',
    'Ascelia Pharma':'ACE.ST','Bactiguard B':'BACTI-B.ST','Bergman & Beving B':'BERG-B.ST',
    'Betsson B':'BETS-B.ST','BHG Group':'BHG.ST','Boozt':'BOOZT.ST',
    'Cibus Nordic Real Estate':'CIBUS.ST','Cloetta B':'CLA-B.ST','Coala-Life':'COALA.ST',
    'CTT Systems':'CTT.ST','Duni':'DUNI.ST','Eastnine':'EAST.ST',
    'Elanders B':'ELAN-B.ST','Engcon B':'ENGCON-B.ST','Fenix Outdoor B':'FOI-B.ST',
    'Fortnox':'FNOX.ST','Garo':'GARO.ST','Hansa Biopharma':'HNSA.ST',
    'Heba B':'HEBA-B.ST','Hexpol B':'HPOL-B.ST','Instalco':'INSTAL.ST',
    'Inwido':'INWI.ST','Karnov Group':'KAR.ST','Knowit':'KNOW.ST',
    'Lime Technologies':'LIME.ST','Maha Energy':'MAHA-A.ST','Medcap':'MCAP.ST',
    'Midsona B':'MSON-B.ST','MilDef Group':'MILDEF.ST','Momentum Group B':'MMGR-B.ST',
    'Nelly Group':'NELLY.ST','Nolato B':'NOLA-B.ST','OEM International B':'OEM-B.ST',
    'Scandi Standard':'SCST.ST','Scandic Hotels':'SHOT.ST','Scandinavian Enviro Systems':'SES.ST',
    'Sdiptech B':'SDIP-B.ST','Senzime':'SEZI.ST','Scandinavian Tobacco Group':'STG.CO',
    'SinterCast':'SINT.ST','Sleep Cycle':'SLEEP.ST','Softronic B':'SOF-B.ST',
    'Solid Försäkring':'SFAB.ST','Stillfront Group':'SF.ST','Storytel B':'STORY-B.ST',
    'Studsvik':'SVIK.ST','Surgical Science':'SUS.ST','SynAct Pharma':'SYNACT.ST',
    'TF Bank':'TFBANK.ST','Tobii Dynavox':'TDVOX.ST','Troax Group':'TROAX.ST',
    'Vestum':'VESTUM.ST','VBG Group B':'VBG-B.ST','Volati AB':'VOLO.ST',
    'Wästbygg Gruppen B':'WBGR-B.ST','XANO Industri B':'XANO-B.ST'
}
ALL_COMPANIES = {**{name:ticker for name,(ticker,_) in CAT.items()}, **ADDITIONAL_COMPANIES}

with st.sidebar:
    st.header('Välj bolag')
    st.caption('Sök genom att skriva i listan eller öppna den och skrolla bland alla bolag.')
    options = sorted(ALL_COMPANIES, key=str.casefold)
    name = st.selectbox(
        'Sök eller välj bolag',
        options,
        index=None,
        placeholder='Skriv bolagsnamn eller klicka för att välja…',
        key='company_picker',
        help='Klicka i fältet för att öppna hela listan. Börja skriva för att filtrera.'
    )
    ticker = ALL_COMPANIES[name] if name else ''
    st.caption(f'Ticker: {ticker or "–"}')
    if st.button('Uppdatera data', use_container_width=True):
        st.cache_data.clear()
        st.rerun()
if not ticker:st.stop()
with st.spinner(f'Hämtar data för {name or ticker}…'):
    try: info,hist,fin,cash,balance,quick,fetch_errors=fetch(ticker)
    except Exception as ex:
        st.error(f'Datakällan kunde inte hämtas: {ex}. Försök igen senare eller välj annan ticker.');st.stop()
if not info and hist.empty and all(f.empty for f in (fin,cash,balance)) and not quick:
    st.error('Ingen data kunde hämtas från Yahoo Finance. Det kan vara ett tillfälligt API-fel eller en begränsning på Streamlit-servern.')
    with st.expander('Visa teknisk felsökning'):
        st.json(fetch_errors)
    st.stop()
if fetch_errors:
    with st.expander('Datakällans status – visa eventuella fel'):
        st.json(fetch_errors)
company=info.get('longName') or name or ticker
st.subheader(company)
st.caption(f'{ticker} · Valuta: {info.get("currency","okänd")} · Börsdata hämtad {dt.datetime.now().strftime("%Y-%m-%d %H:%M")} · Yahoo Finance')

price=val(info,'currentPrice','regularMarketPrice','previousClose')
if price is None: price=val(quick,'last_price','previous_close')
if price is None and not hist.empty:price=numeric(hist['Close'].iloc[-1])
revenue=val(info,'totalRevenue');ebitda=val(info,'ebitda');margin=val(info,'operatingMargins');debt=val(info,'totalDebt');cash_val=val(info,'totalCash');opcf=val(info,'operatingCashflow');eps=val(info,'trailingEps');pe=val(info,'trailingPE')
netdebt=debt-cash_val if debt is not None and cash_val is not None else None

# Fallback: derive missing fundamentals from reported annual statements.
def statement_value(frame, labels):
    if frame is None or frame.empty: return None
    for label in labels:
        if label in frame.index:
            for item in frame.loc[label].tolist():
                n=numeric(item)
                if n is not None: return n
    return None

revenue_source='Yahoo Finance – bolagsprofil'
if revenue is None:
    revenue=statement_value(fin,['Total Revenue','Operating Revenue'])
    if revenue is not None: revenue_source='Yahoo Finance – senaste tillgängliga årsredovisningsperiod'
if ebitda is None:
    ebitda=statement_value(fin,['EBITDA','Normalized EBITDA'])
if opcf is None:
    opcf=statement_value(cash,['Operating Cash Flow','Cash Flow From Continuing Operating Activities'])
if debt is None:
    debt=statement_value(balance,['Total Debt','Long Term Debt And Capital Lease Obligation'])
if cash_val is None:
    cash_val=statement_value(balance,['Cash And Cash Equivalents','Cash Cash Equivalents And Short Term Investments'])
if netdebt is None and debt is not None and cash_val is not None: netdebt=debt-cash_val
if margin is None and revenue:
    operating_income=statement_value(fin,['Operating Income','EBIT'])
    if operating_income is not None: margin=operating_income/revenue
if eps is None:
    eps=statement_value(fin,['Diluted EPS','Basic EPS'])
# Beräkna inte P/E från EPS utan säkerställd valuta.


# Bedömningen använder tillgängliga års-, kvartals- och värderingsserier.
# Saknad data ger grå markering, inte ett godtyckligt neutralt eller positivt betyg.
with st.spinner('Kontrollerar fördjupade rapport- och värderingsdata…'):
    extended, extended_errors = extended_financials(ticker)
_signal_title,_signal_color,_assessment_rows,_assessment_score,_signal_detail = assessment(
    fin,cash,balance,info,price,netdebt,ebitda,hist,extended,quick
)
st.markdown(
    f'<div style="border:1px solid {_signal_color};border-left:8px solid {_signal_color};border-radius:12px;padding:15px 20px;margin:16px 0;">'
    f'<div style="font-size:21px;font-weight:700;color:{_signal_color}">{_signal_title}</div>'
    f'<div style="margin-top:5px;">{_signal_detail}</div></div>',unsafe_allow_html=True
)
st.caption('Modellbedömning, inte statistiskt fastställd sannolikhet för framtida kursuppgång. Se Analys om bolaget för datatäckning och metod.')

# Finansiell översikt: jämförbara nyckeltal utan konstruerade EBITA-värden.
ebita = statement_value(fin, ['EBITA', 'Adjusted EBITA', 'Normalized EBITA'])
sales_history = _annual_series(fin, ['Total Revenue', 'Operating Revenue']) if '_annual_series' in globals() else pd.Series(dtype='float64')
sales_growth_pct = _annual_growth(sales_history)*100 if len(sales_history)>=2 and _annual_growth(sales_history) is not None else None
netdebt_ebitda = netdebt/ebitda if netdebt is not None and ebitda is not None and ebitda>0 else None
shares_count = val(info,'sharesOutstanding')
if shares_count is None: shares_count=val(quick,'shares')
market_cap = val(info,'marketCap')
if market_cap is None: market_cap=val(quick,'market_cap')
if market_cap is None and shares_count is not None and price is not None: market_cap=shares_count*price
quote_ccy = info.get('currency') or quick.get('currency')
report_ccy = info.get('financialCurrency')
_roic_data, _roic_error = {}, None
if roic_key():
    _roic_data,_roic_error = roic_company_data(ticker,roic_key())
_roic_income=_roic_data.get('income',[])
# P/E: använd års-EPS från samma rapportår som ROIC:s verifierade rapportvaluta.
pe_source = 'Yahoo Finance trailingPE' if pe is not None else 'Saknas'
eps_source_currency = None
eps_report_period = None
if pe is None and price is not None and price>0 and quote_ccy:
    if _roic_income and isinstance(_roic_income[0].get('currency'),str):
        report = _roic_income[0]
        report_end = str(report.get('period_end_date') or '')[:10]
        report_year = report_end[:4]
        eps_value = None
        eps_origin = None
        # Prova rapporterad EPS från ROIC, utan att anta ett enskilt fältnamn.
        for key in ('eps_diluted','eps_basic','eps','is_eps_diluted','is_eps_basic','is_earnings_per_share_diluted','is_earnings_per_share_basic'):
            candidate = numeric(report.get(key))
            if candidate is not None and candidate>0:
                eps_value,eps_origin=candidate,'ROIC'
                break
        # Yahoo års-EPS får bara användas när årsperioden överensstämmer.
        if eps_value is None and report_year and fin is not None and not fin.empty:
            for column in fin.columns:
                try: column_year=str(pd.Timestamp(column).year)
                except (TypeError,ValueError): continue
                if column_year!=report_year: continue
                for row in ('Diluted EPS','Basic EPS'):
                    if row in fin.index:
                        candidate=numeric(fin.loc[row,column])
                        if candidate is not None and candidate>0:
                            eps_value,eps_origin=candidate,'Yahoo årsresultat'
                            break
                if eps_value is not None: break
        if eps_value is not None:
            eps_source_currency=report['currency']
            eps_report_period=report_end
            rate=fx_rate(eps_source_currency,quote_ccy)
            if rate is not None and rate>0:
                pe=price/(eps_value*rate)
                pe_source='Beräknat: '+eps_origin+' års-EPS '+report_year+', '+eps_source_currency+'/'+quote_ccy+' valutakurs'
    elif report_ccy and report_ccy==quote_ccy and eps is not None and eps>0:
        pe=price/eps
        pe_source='Beräknat från Yahoo EPS (samma valuta)'
# EV/EBITA: börsvärde i handelsvalutan + nettoskuld i rapportvalutan,
# endast om nettoskuld och EBITA avser jämförbar period och valuta.
ev_ebita = None
ev_ebita_source = 'Saknar jämförbar nettoskuld och EBITA'
roic_balance = _roic_data.get('balance',[])
if _roic_income and roic_balance and market_cap is not None and quote_ccy:
    inc=_roic_income[0]
    report_date=str(inc.get('period_end_date') or '')[:10]
    report_currency=inc.get('currency')
    balance_match=next((x for x in roic_balance if str(x.get('period_end_date') or '')[:10]==report_date and x.get('currency')==report_currency),None)
    def roic_field(record,keys):
        if not record: return None
        for key in keys:
            value=numeric(record.get(key))
            if value is not None: return value
        return None
    if balance_match and report_currency:
        debt_roic=roic_field(balance_match,('bs_total_debt','total_debt','totalDebt'))
        cash_roic=roic_field(balance_match,('bs_cash_and_cash_equivalents','cash_and_cash_equivalents','cash_and_short_term_investments'))
        ebita_roic=roic_field(inc,('ebita',))
        rate=fx_rate(quote_ccy,report_currency)
        if all(v is not None for v in (debt_roic,cash_roic,ebita_roic,rate)) and ebita_roic>0 and rate>0:
            ev_ebita=(market_cap*rate+debt_roic-cash_roic)/ebita_roic
            ev_ebita_source='ROIC årsbalans och EBITA '+report_date+'; börsvärde omräknat '+quote_ccy+' till '+report_currency
# Tre kompletterande nyckeltal: kassagenerering, kapitalavkastning och utdelning.
free_cash_flow = statement_value(cash,['Free Cash Flow'])
if free_cash_flow is None:
    operating_cash_reported = statement_value(cash,['Operating Cash Flow','Cash Flow From Continuing Operating Activities'])
    capital_expenditure = statement_value(cash,['Capital Expenditure'])
    if operating_cash_reported is not None and capital_expenditure is not None:
        free_cash_flow = operating_cash_reported - abs(capital_expenditure)
net_income_reported = statement_value(fin,['Net Income','Net Income Common Stockholders'])
equity_reported = statement_value(balance,['Stockholders Equity','Common Stock Equity'])
roe_overview = net_income_reported/equity_reported*100 if net_income_reported is not None and equity_reported is not None and equity_reported>0 else None
dividend_yield = val(info,'dividendYield')
# Yahoo kan ange yield som procenttal eller decimal; avstå om formatet inte kan säkerställas.
if dividend_yield is not None:
    dividend_yield = dividend_yield*100 if 0<=dividend_yield<=1 else (dividend_yield if 1<dividend_yield<=30 else None)
if dividend_yield is None:
    dividend_yield = trailing_dividend_yield(ticker, price)
# ROIC kompletterar endast verifierade, saknade värden.
if _roic_income:
    latest=_roic_income[0]
    roic_currency=latest.get('currency')
    if ebita is None and isinstance(latest.get('ebita'),(int,float)):
        ebita=latest['ebita']
    if sales_growth_pct is None and len(_roic_income)>=2:
        current_rev,prior_rev=latest.get('is_sales_revenue_turnover'),_roic_income[1].get('is_sales_revenue_turnover')
        if isinstance(current_rev,(int,float)) and isinstance(prior_rev,(int,float)) and prior_rev>0:
            sales_growth_pct=(current_rev/prior_rev-1)*100
    # EV/EBITA kräver samma period och valuta; blandning med Yahoo-värden undviks.
st.markdown('### Finansiell översikt')
overview_items = [
    ('Börsvärde',show(market_cap/1e9,' mdkr',2) if market_cap is not None and quote_ccy=='SEK' else (show(market_cap/1e9,' md '+(quote_ccy or ''),2) if market_cap is not None else 'Saknas')),
    ('Aktiekurs',show(price,' '+(quote_ccy or ''),2)),
    ('Omsättning (LTM/helår)',show(revenue/1e6,' M',0) if revenue is not None else 'Saknas'),
    ('Omsättningstillväxt',show(sales_growth_pct,' %/år',1) if sales_growth_pct is not None else 'Saknas'),
    ('Rörelsemarginal',show(margin*100,' %') if margin is not None else 'Saknas'),
    ('EBITA (helår)',show(ebita/1e6,' M',0) if ebita is not None else 'Saknas'),
    ('Avkastning på eget kapital (ROE)',show(roe_overview,' %',1) if roe_overview is not None else 'Saknas'),
    ('Operativt kassaflöde',show(opcf/1e6,' M',0) if opcf is not None else 'Saknas'),
    ('P/E',show(pe,'×') if pe is not None else 'Saknas'),
    ('EV/EBITA',show(ev_ebita,'×',1) if ev_ebita is not None else 'Saknas'),
    ('Nettoskuld/EBITDA',show(netdebt_ebitda,'×',2) if netdebt_ebitda is not None else 'Saknas'),
    ('Direktavkastning',show(dividend_yield,' %',1) if dividend_yield is not None else 'Saknas')
]
for start in range(0,len(overview_items),4):
    cols=st.columns(4)
    for col,(label,value) in zip(cols,overview_items[start:start+4]): col.metric(label,value)
if not roic_key():
    st.caption('ROIC AI är inte anslutet ännu. Lägg ROIC_API_KEY i Streamlit Secrets för automatisk komplettering av rapportdata.')
elif _roic_error:
    st.caption('ROIC AI: '+_roic_error)
elif _roic_income:
    st.caption('Kompletterande rapportdata från ROIC AI (årsperiod: '+str(_roic_income[0].get('period_end_date','okänd'))+', valuta: '+str(_roic_income[0].get('currency','okänd'))+').')
st.caption('Omsättning och kassaflöde kan avse olika perioder. Omsättningstillväxt = CAGR över tillgängliga rapporterade årsperioder. EBITA visas bara om det finns som separat rapporterad post; EBIT och EBITDA används aldrig som ersättning. EV/EBITA visas endast när rapport- och handelsvaluta matchar. ROE bygger på senaste rapporterade nettoresultat och eget kapital och är en förenklad kvot, inte ett genomsnittligt eget kapital. M = miljoner.')

one,two,three,four,five=st.tabs(['Kurs och historik','Rapporterade data','Nyheter – 6 månader','Dokument och scenarier','Analys om bolaget'])
with one:
    if not hist.empty:
        h=hist.reset_index();h['Datum']=pd.to_datetime(h['Date']).dt.tz_localize(None)
        def mark_high_low(fig, data):
            # Samma stängningskurser som linjen; intradagshögsta/-lägsta avses inte.
            clean = data[['Datum','Close']].copy()
            clean['Close'] = pd.to_numeric(clean['Close'],errors='coerce')
            clean = clean.dropna(subset=['Datum','Close'])
            if clean.empty: return fig
            high = clean.loc[clean['Close'].idxmax()]
            low = clean.loc[clean['Close'].idxmin()]
            for row, name, color, position in [
                (high,'Högsta stängningskurs','#16803C','top center'),
                (low,'Lägsta stängningskurs','#C53030','bottom center')
            ]:
                fig.add_scatter(
                    x=[row['Datum']],y=[row['Close']],mode='markers+text',
                    marker=dict(size=12,color=color,line=dict(color='white',width=2)),
                    text=[f"{name}: {row['Close']:.2f}"],
                    textposition=position,textfont=dict(color=color,size=12),
                    name=name,hovertemplate='%{x|%Y-%m-%d}<br>%{y:.2f}<extra>'+name+'</extra>'
                )
            return fig
        h_one_year = h[h['Datum'] >= h['Datum'].max() - pd.DateOffset(years=1)].copy()
        if len(h_one_year) >= 2:
            chart_one_year = px.line(h_one_year,x='Datum',y='Close',title='Aktiekurs – senaste året')
            chart_one_year.update_layout(xaxis_title='Datum',yaxis_title='Aktiekurs')
            st.plotly_chart(mark_high_low(chart_one_year,h_one_year),use_container_width=True)
        else:
            st.info('Det finns inte tillräckligt med kurshistorik för att visa ett år.')
        chart_five_years = px.line(h,x='Datum',y='Close',title='Aktiekurs – senaste fem åren')
        chart_five_years.update_layout(xaxis_title='Datum',yaxis_title='Aktiekurs')
        st.plotly_chart(mark_high_low(chart_five_years,h),use_container_width=True)
    else:st.warning('Kursserie saknas.')
    st.markdown('#### Bolagsbeskrivning')
    verified = VERIFIED_COMPANY_DESCRIPTIONS.get(ticker)
    description = (info.get('longBusinessSummary') or '').strip()
    if verified:
        st.write(verified[0])
        st.markdown(f'[Källa: Bolagets officiella webbplats]({verified[1]})')
    elif description:
        st.write(description)
        st.caption('Källa: Yahoo Finance – bolagsprofil.')
    else:
        description, description_url, description_source = company_description_fallback(name or company)
        if description:
            st.write(description)
            st.caption('Källa: '+description_source+' – uppslagsverkstext, kontrollera mot bolagets egen webbplats.')
            if description_url: st.markdown(f'[Läs originalbeskrivningen]({description_url})')
        else:
            st.info('Ingen verifierad bolagsbeskrivning kunde hämtas automatiskt från Yahoo Finance eller Wikipedia för detta bolag. Vi visar inte en påhittad beskrivning.')
    if info.get('website'):
        st.markdown(f"[Bolagets webbplats]({info['website']})")
with two:
    st.caption('Finansiella rapportdata från Yahoo Finance. Belopp visas i miljoner (M) av bolagets rapporteringsvaluta. År avser rapportperiodens slutdatum, inte nödvändigtvis kalenderår.')

    # Explicit mapping only: do not substitute EBIT for EBITA or infer missing figures.
    report_sections = [
        ('Resultaträkning', fin, [
            ('Omsättning', ['Total Revenue', 'Operating Revenue'], 'money'),
            ('Bruttoresultat', ['Gross Profit'], 'money'),
            ('EBITDA', ['EBITDA'], 'money'),
            ('Rörelseresultat (EBIT)', ['Operating Income', 'EBIT'], 'money'),
            ('Resultat före skatt', ['Pretax Income'], 'money'),
            ('Årets resultat', ['Net Income', 'Net Income Common Stockholders'], 'money'),
            ('Resultat per aktie (EPS)', ['Diluted EPS', 'Basic EPS'], 'per_share'),
        ]),
        ('Kassaflöde', cash, [
            ('Kassaflöde från löpande verksamheten', ['Operating Cash Flow', 'Cash Flow From Continuing Operating Activities'], 'money'),
            ('Investeringar i anläggningstillgångar', ['Capital Expenditure'], 'money'),
            ('Fritt kassaflöde', ['Free Cash Flow'], 'money'),
            ('Kassaflöde från investeringar', ['Investing Cash Flow'], 'money'),
            ('Kassaflöde från finansiering', ['Financing Cash Flow'], 'money'),
        ]),
        ('Balansräkning', balance, [
            ('Totala tillgångar', ['Total Assets'], 'money'),
            ('Eget kapital', ['Stockholders Equity', 'Total Equity Gross Minority Interest'], 'money'),
            ('Totala skulder', ['Total Liabilities Net Minority Interest'], 'money'),
            ('Räntebärande skulder', ['Total Debt'], 'money'),
            ('Likvida medel', ['Cash And Cash Equivalents', 'Cash Cash Equivalents And Short Term Investments'], 'money'),
        ]),
    ]
    for section_name, frame, metrics in report_sections:
        st.markdown(f'#### {section_name}')
        if frame is None or frame.empty:
            st.info('Rapportdata saknas för denna del.')
            continue
        cols = sorted(frame.columns, reverse=True)[:5]
        periods = [str(c.date()) if hasattr(c, 'date') else str(c)[:10] for c in cols]
        rows = []
        for label, candidates, unit in metrics:
            found = next((key for key in candidates if key in frame.index), None)
            values = []
            for col in cols:
                raw = numeric(frame.at[found, col]) if found else None
                if raw is None:
                    values.append('Saknas')
                elif unit == 'money':
                    values.append(show(raw / 1e6, digits=1))
                else:
                    values.append(show(raw, digits=2))
            rows.append([label] + values)
        # Compute only directly comparable ratios from the same reporting period.
        if section_name == 'Resultaträkning':
            ratio_values = []
            for col in cols:
                revenue_key = next((k for k in ['Total Revenue', 'Operating Revenue'] if k in frame.index), None)
                ebit_key = next((k for k in ['Operating Income', 'EBIT'] if k in frame.index), None)
                rev = numeric(frame.at[revenue_key, col]) if revenue_key else None
                ebit = numeric(frame.at[ebit_key, col]) if ebit_key else None
                ratio_values.append(show(100 * ebit / rev, ' %', 1) if rev and ebit is not None else 'Saknas')
            rows.insert(4, ['Rörelsemarginal (EBIT)'] + ratio_values)
        table = pd.DataFrame(rows, columns=['Nyckeltal'] + periods)
        st.dataframe(table, use_container_width=True, hide_index=True)
        st.caption('Belopp i M (miljoner), utom EPS och procenttal. ”Saknas” betyder att dataleverantören inte har lämnat uppgiften.')
        if section_name == 'Resultaträkning':
            st.caption('EBITA och EBITA-marginal visas inte eftersom de inte kan fastställas säkert från dessa standardiserade Yahoo-fält.')
        with st.expander(f'Visa samtliga originalposter – {section_name.lower()}'):
            raw_table = frame.copy()
            raw_table.columns = [str(c.date()) if hasattr(c, 'date') else str(c) for c in raw_table.columns]
            st.dataframe(raw_table, use_container_width=True)
with three:
    st.caption('Nyheter från de senaste sex månaderna. Sökningen körs först när du begär den, så bolagsbyten går snabbare.')
    if st.button('Hämta nyheter för valt bolag', key='refresh_news'):
        news.clear()
        st.session_state['news_requested_for'] = ticker
    if st.session_state.get('news_requested_for') == ticker:
        with st.spinner('Hämtar nyheter…'):
            items, news_diagnostics = news(company)
        if items:
            st.success(f'{len(items)} artiklar hittades.')
            for item in items:
                st.markdown(f'**{item["Datum"]}** · [{item["Rubrik"]}]({item["Länk"]})')
        else:
            st.warning('Inga artiklar kunde hämtas. Det betyder inte att bolaget saknar nyheter.')
        with st.expander('Visa nyhetskällans diagnostik'):
            for line in news_diagnostics:
                st.text(line)
    else:
        st.info('Klicka på Hämta nyheter för valt bolag för att starta sökningen.')
    search_link = 'https://www.google.com/search?' + urllib.parse.urlencode({
        'q': f'{company} pressmeddelanden senaste sex månaderna'
    })
    st.markdown(f'[Sök pressmeddelanden i webbläsaren]({search_link})')
with four:
    st.markdown('#### Dokumentbibliotek för denna analys')
    uploads=st.file_uploader('Ladda upp årsredovisningar, kvartalsrapporter och egna kalkyler',type=['pdf','docx','xlsx','csv','txt'],accept_multiple_files=True)
    documents=[]
    for u in uploads or []:
        try:
            content=parse_file(u)
            documents.append((u.name,content))
            st.success(f'{u.name}: {len(content):,} tecken extraherade')
        except Exception as ex:st.warning(f'{u.name} kunde inte läsas: {ex}')
    if documents:
        term=st.text_input('Sök i samtliga uppladdade filer',placeholder='Exempel: förvärv, kassaflöde, EBITA')
        if term:
            for filename,content in documents:
                hits=list(re.finditer(re.escape(term),content,re.I))[:15]
                if hits:
                    with st.expander(f'{filename} – {len(hits)} träffar (max 15 visas)'):
                        for m in hits:st.text(content[max(0,m.start()-180):min(len(content),m.end()+250)].replace('\n',' '))
        st.download_button('Ladda ner extraherad dokumenttext',data='\n\n'.join(f'## {n}\n{c}' for n,c in documents),file_name='dokumentunderlag.txt')
    st.markdown('#### Värderingsscenario (illustrativ modell)')
    if eps is None or eps<=0:st.warning('EPS är saknat eller negativt. Ange ett eget positivt antagande om du vill använda P/E-modellen.')
    a,b,c=st.columns(3)
    base_eps=a.number_input('EPS – utgångsvärde',value=float(eps) if eps and eps>0 else 1.0,min_value=0.01)
    growth=b.slider('Årlig vinsttillväxt (%)',-20,30,7)
    target_pe=c.slider('P/E om fem år',3,40,15)
    rate=st.slider('Diskonteringsränta (%)',0,25,10)
    rows=[]
    for scenario,g,m in [('Negativt',growth-5,target_pe*.75),('Bas',growth,target_pe),('Positivt',growth+5,target_pe*1.25)]:
        future=base_eps*(1+g/100)**5*m;today=future/(1+rate/100)**5
        rows.append({'Scenario':scenario,'Tillväxt %':g,'P/E år 5':round(m,1),'Riktkurs år 5':round(future,2),'Nuvärde':round(today,2),'Potential %':round((today/price-1)*100,1) if price and price>0 else None})
    st.dataframe(pd.DataFrame(rows),hide_index=True,use_container_width=True)
    st.caption('Modellen är en förenklad känslighetsanalys, inte investeringsrådgivning. Den passar sämre för banker, investmentbolag, fastighetsbolag och bolag med förlust.')
    st.download_button('Exportera värdering (CSV)',pd.DataFrame(rows).to_csv(index=False).encode('utf-8-sig'),'vardering.csv','text/csv')
with five:
    st.subheader(f'Analys om bolaget – {company}')
    st.caption('Automatisk historikbaserad analys. Inga manuellt påhittade tillväxttal, multiplar eller sannolikheter används.')
    def series_from_statement(frame, labels):
        if frame is None or frame.empty:
            return pd.Series(dtype='float64')
        for label in labels:
            if label in frame.index:
                values = pd.to_numeric(frame.loc[label], errors='coerce').dropna()
                if not values.empty:
                    return values.sort_index()
        return pd.Series(dtype='float64')

    def growth_rate(series):
        if len(series) < 2:
            return None
        first, last = float(series.iloc[0]), float(series.iloc[-1])
        years = len(series) - 1
        if first <= 0 or last <= 0:
            return None
        return (last / first) ** (1 / years) - 1

    revenues = series_from_statement(fin, ['Total Revenue', 'Operating Revenue'])
    earnings = series_from_statement(fin, ['Operating Income', 'EBIT'])
    profits = series_from_statement(fin, ['Net Income', 'Net Income Common Stockholders'])
    operating_cf = series_from_statement(cash, ['Operating Cash Flow', 'Cash Flow From Continuing Operating Activities'])
    free_cf = series_from_statement(cash, ['Free Cash Flow'])
    sales_growth = growth_rate(revenues)
    profit_growth = growth_rate(profits)
    margins = []
    for period in revenues.index.intersection(earnings.index):
        rev, earn = float(revenues.loc[period]), float(earnings.loc[period])
        if rev > 0:
            margins.append(earn / rev)
    current_margin = margins[-1] if margins else None
    margin_change = (margins[-1] - margins[0]) if len(margins) >= 2 else None
    cf_positive = int((operating_cf > 0).sum()) if not operating_cf.empty else None
    cf_total = len(operating_cf)
    shares_raw = val(info, 'sharesOutstanding') or val(quick, 'shares')
    cap = val(info, 'marketCap')
    if cap is None and price is not None and shares_raw is not None:
        cap = price * shares_raw
    ev_ebitda = (cap + netdebt) / ebitda if cap is not None and netdebt is not None and ebitda and ebitda > 0 else None
    leverage = netdebt / ebitda if netdebt is not None and ebitda and ebitda > 0 else None
    cols = st.columns(4)
    cols[0].metric('Historisk omsättningstillväxt', show(sales_growth * 100, ' %/år') if sales_growth is not None else 'Saknas')
    cols[1].metric('Historisk vinsttillväxt', show(profit_growth * 100, ' %/år') if profit_growth is not None else 'Saknas')
    cols[2].metric('Nettoskuld/EBITDA', show(leverage, '×', 2) if leverage is not None else 'Saknas')
    cols[3].metric('EV/EBITDA', show(ev_ebitda, '×', 1) if ev_ebitda is not None else 'Saknas')
    st.markdown('#### Officiella bolagsrapporter')
    official_source = OFFICIAL_REPORT_SOURCES.get(ticker)
    if official_source:
        st.markdown(f"[Öppna senaste registrerade rapport ({official_source['period']})]({official_source['report']}) · [Rapportarkiv]({official_source['archive']})")
        if st.button('Hämta rapportuppgifter från bolaget',key='official_report_button'):
            st.session_state['official_report_ticker'] = ticker
            official_report_excerpt.clear()
        if st.session_state.get('official_report_ticker') == ticker:
            official_lines, official_error = official_report_excerpt(ticker)
            if official_lines:
                st.caption('Verifierbar rapporttext från bolagets egen webbplats. Texten är inte automatiskt normaliserad till modellens beräkningar.')
                for line in official_lines:
                    st.write(line)
            elif official_error:
                st.warning('Rapporten kunde inte hämtas: '+official_error)
            else:
                st.info('Ingen läsbar finansiell rapporttext hittades på sidan. Öppna originalrapporten via länken.')
    else:
        st.caption('Ingen officiell rapportkälla är ännu registrerad för detta bolag. Yahoo Finance är tills vidare huvudkälla.')
    st.markdown('#### Fördjupad investeringsbedömning')
    st.write(_signal_title)
    if _assessment_score is not None:
        st.metric('Modellpoäng',f'{_assessment_score*100:.0f} av 100')
    st.caption(_signal_detail)
    if 'GRÅ' in _signal_title:
        st.info('Grått betyder att appens anslutna datakällor inte räcker för en investeringsvärdering – inte att bolaget saknar offentliga rapporter.')

    if _assessment_rows:
        st.dataframe(pd.DataFrame(_assessment_rows),hide_index=True,use_container_width=True)
    with st.expander('Datatäckning och källstatus'):
        st.write('Årsperioder i resultaträkning:',len(fin.columns) if not fin.empty else 0)
        st.write('Kvartalsperioder i resultaträkning:',len(extended['quarterly_income'].columns) if not extended['quarterly_income'].empty else 0)
        st.write('Historiska värderingsperioder:',len(extended['valuation_history'].columns) if not extended['valuation_history'].empty else 0)
        if extended_errors: st.json(extended_errors)
    st.warning('Historisk backtestning med publiceringsdatum för varje rapport och jämförelse mot ett totalavkastningsindex är ännu inte genomförd. Modellen ska inte tolkas som validerad prognos.')
    st.markdown('#### Historisk finansiell utveckling')
    history_rows = []
    for label, data in [('Omsättning', revenues), ('Rörelseresultat (EBIT)', earnings), ('Nettoresultat', profits), ('Operativt kassaflöde', operating_cf), ('Fritt kassaflöde', free_cf)]:
        for period, amount in data.items():
            history_rows.append({'Nyckeltal': label, 'Rapportperiod': str(period)[:10], 'Belopp (miljoner)': round(float(amount)/1e6, 1)})
    if history_rows:
        frame_history = pd.DataFrame(history_rows)
        st.dataframe(frame_history.pivot_table(index='Nyckeltal', columns='Rapportperiod', values='Belopp (miljoner)', aggfunc='first').reset_index(), hide_index=True, use_container_width=True)
    else:
        st.warning('Historiska rapportdata saknas. Ingen fundamental bedömning kan göras utan verifierbara underlag.')

    st.markdown('#### Modellbaserad bedömning')
    tests = []
    def check(name, value, condition, reason):
        if value is not None:
            tests.append({'Faktor': name, 'Bedömning': 'Positiv' if condition(value) else 'Negativ', 'Förklaring': reason})
    check('Omsättningstillväxt', sales_growth, lambda x: x > 0, 'Historisk genomsnittlig årstillväxt över eller under noll')
    check('Vinsttillväxt', profit_growth, lambda x: x > 0, 'Historisk genomsnittlig årstillväxt över eller under noll')
    check('Rörelsemarginal', current_margin, lambda x: x > 0, 'Senaste jämförbara årsperiod med positiv EBIT-marginal')
    check('Marginalutveckling', margin_change, lambda x: x > 0, 'Förändring i EBIT-marginal mellan första och senaste tillgängliga period')
    check('Skuldsättning', leverage, lambda x: x < 3, 'Illustrativ riskgräns: nettoskuld/EBITDA under 3×')
    if cf_total:
        check('Operativt kassaflöde', cf_positive / cf_total, lambda x: x >= 0.75, 'Positivt kassaflöde i minst 75 % av tillgängliga årsperioder')
    if tests:
        positives = sum(row['Bedömning'] == 'Positiv' for row in tests)
        st.metric('Positiva finansiella signaler', f'{positives} av {len(tests)}')
        st.dataframe(pd.DataFrame(tests), hide_index=True, use_container_width=True)
        st.write('**Tolkning:** ' + ('Övervägande positiva historiska signaler.' if positives > len(tests)/2 else 'Övervägande negativa historiska signaler.' if positives < len(tests)/2 else 'Blandade historiska signaler.'))
    else:
        st.info('För få historiska nyckeltal för att göra en bedömning.')
    st.markdown('#### Historisk kursutveckling och risk')
    if not hist.empty and 'Close' in hist:
        prices = pd.to_numeric(hist['Close'], errors='coerce').dropna()
        if len(prices) >= 60:
            returns = prices.pct_change().dropna()
            yearly = (prices.iloc[-1]/prices.iloc[0])**(252/max(len(prices)-1,1))-1 if prices.iloc[0]>0 else None
            volatility = returns.std() * (252**0.5)
            drawdown = (prices/prices.cummax()-1).min()
            x,y,z = st.columns(3)
            x.metric('Genomsnittlig årlig kursutveckling', show(yearly*100,' %') if yearly is not None else 'Saknas')
            y.metric('Historisk volatilitet', show(volatility*100,' %'))
            z.metric('Största historiska nedgång', show(drawdown*100,' %'))
            st.caption('Beräknat på tillgängliga handelsdagar. Kursförändringen exkluderar utdelningar när ej justerad kurs används. Historisk utveckling är ingen prognos.')
        else:
            st.info('För kort kurshistorik för riskmåtten.')
    else:
        st.info('Kurshistorik saknas.')
    st.caption('Detta är en transparent indikatoranalys, inte en statistiskt kalibrerad sannolikhet för uppgång eller nedgång. Tröskelvärdena är metodval, inte empiriskt fastställda prognoser. Värdering kräver dessutom jämförelse mot bransch och historiska multiplar. Banken, investmentbolag och fastighetsbolag behöver andra nyckeltal.')

st.divider()
st.caption('Källor: Yahoo Finance via yfinance för nyckeltal och historik; Google News RSS för nyhetssökning; användarens egna uppladdade dokument. Inga automatiska AI-slutsatser påstås ha gjorts. Appen körs på egen Streamlit-server och filerna behandlas i den sessionen.')
