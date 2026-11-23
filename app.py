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

@st.cache_data(ttl=21600,show_spinner=False)
def roic_quarterly_income(identifier, api_key):
    """Kompletterande rapporterade kvartal, med samma identifierare som årsdata."""
    if not identifier or ':' not in identifier or not api_key:
        return [],'ROIC-instrument eller API-nyckel saknas'
    try:
        url='https://api.roic.ai/v3.0.0/fundamental/income-statement/'+urllib.parse.quote(identifier,safe=':')
        response=requests.get(url,params={'period_type':'quarterly','limit':12,'order':'desc'},
                              headers={'Authorization':'Bearer '+api_key},timeout=12)
        if not response.ok:
            return [],'ROIC kvartalsdata: HTTP '+str(response.status_code)
        payload=response.json()
        records=payload.get('data',[]) if isinstance(payload,dict) else []
        records=[v for v in records if isinstance(v,dict) and v.get('symbol')==identifier and v.get('period_type')=='quarterly' and v.get('period_end_date')]
        records.sort(key=lambda v:str(v.get('period_end_date')),reverse=True)
        return records,None if records else 'ROIC API lämnade inga kvartalsrapporter för instrumentet'
    except (requests.RequestException,ValueError,TypeError):
        return [],'ROIC kvartalsdata kunde inte hämtas'

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

# Jämförelsebolag och dataladdare för Investeringsanalys.
# Definitionerna måste finnas innan fliken använder dem.
PEER_UNIVERSE={
    'BETS-B.ST':[('Evolution','EVO.ST','Spelbransch, annan affärsmodell'),('Flutter Entertainment','FLUT','Internationell speloperatör')],
    'VOLO.ST':[('Bergman & Beving','BERG-B.ST','Svensk industrigrupp'),('Lagercrantz','LAGR-B.ST','Svensk industrigrupp'),('Indutrade','INDT.ST','Svensk industrigrupp')],
    'SALIX.ST':[('Bergman & Beving','BERG-B.ST','Svensk industrigrupp'),('Addtech','ADDT-B.ST','Svensk industrigrupp'),('Lagercrantz','LAGR-B.ST','Svensk industrigrupp')]
}
@st.cache_data(ttl=21600,show_spinner=False)
def peer_snapshot(symbol):
    try:
        p=yf.Ticker(symbol).info or {}
        pe=numeric(p.get('trailingPE'))
        margin=numeric(p.get('operatingMargins'))
        growth=numeric(p.get('revenueGrowth'))
        debt=numeric(p.get('totalDebt'))
        cash=numeric(p.get('totalCash'))
        ebitda=numeric(p.get('ebitda'))
        leverage=(debt-cash)/ebitda if None not in (debt,cash,ebitda) and ebitda>0 else None
        return {'pe':pe if pe is not None and pe>0 else None,
                'margin':margin*100 if margin is not None else None,
                'growth':growth*100 if growth is not None else None,
                'leverage':leverage}
    except Exception:
        return {}

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

# Värderingskategori skiljs från branschgrupp. Okända typer får ingen P/E-köpsignal.
INVESTMENT_TICKERS={
    'INVE-A.ST','INVE-B.ST','INDU-C.ST','LATO-B.ST','LUND-B.ST',
    'KINV-B.ST','BURE.ST','EQT.ST'
}
BANK_TICKERS={'SEB-A.ST','SHB-A.ST','SWED-A.ST','NDA-SE.ST','AZA.ST','SAVE.ST','TFBANK.ST'}
INSURANCE_TICKERS={'SAMPO-SDB.ST','SFAB.ST'}
PROPERTY_TICKERS={
    'BALD-B.ST','CAST.ST','CATE.ST','FABG.ST','SAGA-B.ST','WALL-B.ST',
    'WIHL.ST','HUFV-A.ST','ATRLJ-B.ST','PNDX-B.ST','NP3.ST','CORE-B.ST',
    'DIOS.ST','PLAZ-B.ST','STEF-B.ST','SLP-B.ST','CIBUS.ST','EAST.ST','HEBA-B.ST'
}
# Verifierade substansvärden från publicerade bolagsrapporter.
# Datum avser värderingsdatum, inte datum då börskursen hämtades.
REPORTED_NAV={
    'LATO-B.ST':{'nav':223.0,'date':'2026-11-02','source':'https://www.mfn.se/cis/a/latour/delarsrapport-januari-september-2026-d993ea5e.iframe','kind':'Substansvärde'},
    'INDU-C.ST':{'nav':517.0,'date':'2026-09-30','source':'https://mfn.se/a/industrivarden/delarsrapport-1-januari-30-september-2026.iframe','kind':'Substansvärde'}
}
def borsapi_key():
    try: return str(st.secrets.get('BORSAPI_API_KEY','')).strip()
    except Exception: return ''

@st.cache_data(ttl=86400,show_spinner=False)
def borsapi_reports(symbol,api_key):
    """Sekundär rapportkälla för hela universum, inte bara utvalda bolag."""
    if not api_key: return {},'BörsAPI-nyckel saknas'
    code=symbol.removesuffix('.ST').removesuffix('.CO')
    try:
        response=requests.get(
            'https://borsapi.se/api/v1/companies/'+urllib.parse.quote(code,safe='')+'/reports',
            params={'period_type':'year','entity_type':'CONSOLIDATED','merge':'true','limit':5},
            headers={'Authorization':'Bearer '+api_key},timeout=18)
        if not response.ok: return {},'BörsAPI HTTP '+str(response.status_code)
        payload=response.json()
        data=payload.get('data',[]) if isinstance(payload,dict) else []
        return {'periods':data if isinstance(data,list) else []},''
    except Exception as exc:
        return {},type(exc).__name__

def valuation_category(symbol, info=None):
    """Konservativ kategorisering: särskilda affärsmodeller kräver andra metoder."""
    if symbol in INVESTMENT_TICKERS: return 'Investmentbolag / kapitalförvaltning'
    if symbol in BANK_TICKERS: return 'Bank / finansiella tjänster'
    if symbol in INSURANCE_TICKERS: return 'Försäkring'
    if symbol in PROPERTY_TICKERS: return 'Fastigheter'
    details=info or {}
    sector=str(details.get('sector') or '').lower()
    industry=str(details.get('industry') or '').lower()
    combined=sector+' '+industry
    if 'reit' in combined or 'real estate' in combined: return 'Fastigheter'
    if 'bank' in combined or 'credit services' in combined or 'capital markets' in combined: return 'Bank / finansiella tjänster'
    if 'insurance' in combined: return 'Försäkring'
    if 'asset management' in combined or 'investment holding' in combined: return 'Investmentbolag / kapitalförvaltning'
    if 'financial services' in sector or 'financial data' in industry: return 'Bank / finansiella tjänster'
    if not sector and not industry: return 'Ej klassificerad'
    return 'Rörelsedrivande bolag'

@st.cache_data(ttl=21600,show_spinner=False)
def screener_snapshot(symbol):
    """En bolagsobservation med samma P/E-metod som Värderingskontroll."""
    try:
        obj=yf.Ticker(symbol)
        info=obj.info or {}
        category=valuation_category(symbol,info)
        # Försök komplettera varje bolag med en andra, rapportbaserad källa.
        extra,extra_error=borsapi_reports(symbol,borsapi_key())
        report_count=len(extra.get('periods',[]))

        if category!='Rörelsedrivande bolag':
            # Använd sektorsrelevanta nyckeltal men ge inte falska substansrabatter.
            pb=numeric(info.get('priceToBook'))
            roe=numeric(info.get('returnOnEquity'))
            book=numeric(info.get('bookValue'))
            price=numeric(info.get('currentPrice') or info.get('regularMarketPrice'))
            data={'category':category,'currency':info.get('currency'),
                  'pb':pb if pb is not None and pb>0 else None,
                  'roe':roe*100 if roe is not None else None,
                  'book_value':book,'price':price}
            if category=='Investmentbolag / kapitalförvaltning':
                data['method']='NAV / substansrabatt'
                reported=REPORTED_NAV.get(symbol)
                if reported and price is not None and price>0 and info.get('currency')=='SEK':
                    data['nav']=reported['nav']
                    data['nav_date']=reported['date']
                    data['nav_source']=reported['source']
                    data['gap']=(price/reported['nav']-1)*100
                    data['status']='Rapporterad substansrabatt' if data['gap']<0 else 'Rapporterad substanspremie'
                else:
                    data['status']='Substansvärde krävs'
            elif category=='Fastigheter':
                data['status']='Fastighets-NAV krävs'
                data['method']='EPRA NRV/NTA + driftnetto'
            elif category in ('Bank / finansiella tjänster','Försäkring'):
                data['status']='Nyckeltal tillgängliga' if data['pb'] is not None and data['roe'] is not None else 'Otillräckliga sektordata'
                data['method']='P/B + ROE' if category=='Bank / finansiella tjänster' else 'P/B + ROE + solvens'
            else:
                data['status']='Ej klassificerad'
                data['method']='Granskning krävs'
            data['secondary_report_periods']=report_count
            return data
        current_pe=numeric(info.get('trailingPE'))
        price=numeric(info.get('currentPrice') or info.get('regularMarketPrice'))
        if price is None or price<=0:
            try: price=numeric(obj.fast_info.get('last_price'))
            except Exception: pass
        if current_pe is None or not 0<current_pe<=80:
            eps=numeric(info.get('trailingEps'))
            if eps is not None and eps>0 and price is not None and price>0:
                current_pe=price/eps
        if current_pe is None or not 0<current_pe<=80 or price is None or price<=0:
            missing=[]
            if price is None or price<=0: missing.append('aktuell kurs')
            if current_pe is None or not 0<current_pe<=80: missing.append('positiv vinst per aktie (TTM)')
            return {'status':'Otillräckliga börsdata','category':category,'missing':', '.join(missing)}
        history=obj.get_valuation_measures(freq='quarterly',periods=12)
        pe_values=[]
        if isinstance(history,pd.DataFrame) and not history.empty:
            for field in ('PeRatio','PE Ratio','Trailing P/E','Trailing PE','peRatio','trailingPE','TrailingPeRatio'):
                if field in history.index:
                    pe_values+=pd.to_numeric(history.loc[field],errors='coerce').dropna().tolist()
                if field in history.columns:
                    pe_values+=pd.to_numeric(history[field],errors='coerce').dropna().tolist()
        pe_values=[float(v) for v in pe_values if 0<float(v)<=80]
        if len(pe_values)<3:
            return {'status':'P/E-historik saknas','pe':current_pe,'price':price,'observations':len(pe_values),'category':category,'missing':'minst tre jämförbara historiska P/E-observationer'}
        median=float(pd.Series(pe_values).median())
        gap=(current_pe/median-1)*100
        return {'status':'Intressant att analysera vidare' if gap<=-20 else 'Högt värderad' if gap>=20 else 'Avvakta',
                'pe':current_pe,'median':median,'gap':gap,'price':price,
                'indicative_price':price*median/current_pe,'observations':len(pe_values),
                'currency':info.get('currency'),'category':category,'secondary_report_periods':report_count}
    except Exception:
        return {'status':'Kunde inte hämtas'}

st.markdown('### Hitta undervärderade aktier')
if not borsapi_key():
    st.caption('Extra rapportkälla: BörsAPI kan aktiveras med BORSAPI_API_KEY i Streamlit Secrets. Utan nyckeln används Yahoo. BörsAPI:s standardiserade rapporter är inte samma sak som rapporterat substansvärde.')
st.caption('Bolagen klassificeras efter värderingsmetod. P/E-signaler visas bara för rörelsedrivande bolag. Investmentbolag, banker, försäkringsbolag och fastighetsbolag markeras separat tills rätt värderingsmodell finns. En rabatt är inte ett köpbeslut.')
universe=list(dict.fromkeys(ALL_COMPANIES.values()))
screener_names={ticker:name for name,ticker in ALL_COMPANIES.items()}
# Tidigare screeningresultat beräknades utan kategorifilter och måste förnyas.
if st.session_state.get('screener_model_version')!='category-v5':
    st.session_state.screener_results={}
    st.session_state.screener_running=False
    st.session_state.screener_model_version='category-v5'
if 'screener_results' not in st.session_state:
    st.session_state.screener_results={}
if 'screener_running' not in st.session_state:
    st.session_state.screener_running=False
scan_a,scan_b,scan_c=st.columns([2,2,2])
if scan_a.button('Sök igenom alla bolag',key='scan_all_stocks',use_container_width=True):
    st.session_state.screener_running=True
if scan_b.button('Pausa sökningen',key='pause_scan_stocks',use_container_width=True):
    st.session_state.screener_running=False
if scan_c.button('Börja om',key='reset_scan_stocks',use_container_width=True):
    st.session_state.screener_running=False
    st.session_state.screener_results={}
if st.session_state.screener_running:
    pending=[symbol for symbol in universe if symbol not in st.session_state.screener_results]
    for symbol in pending[:8]:
        st.session_state.screener_results[symbol]=screener_snapshot(symbol)
    if len(st.session_state.screener_results)>=len(universe):
        st.session_state.screener_running=False
    else:
        st.caption('Sökningen fortsätter i omgångar för att minska belastningen på datakällan.')
completed=len(st.session_state.screener_results)
st.progress(completed/max(1,len(universe)),text=f'Genomsökta bolag: {completed} av {len(universe)}')
if st.session_state.screener_results:
    scan_rows=[]
    for symbol,result in st.session_state.screener_results.items():
        scan_rows.append({
            'Bolag':screener_names.get(symbol,symbol),'Ticker':symbol,
            'Värderingskategori':result.get('category',valuation_category(symbol)),
            'Rapportperioder (BörsAPI)':result.get('secondary_report_periods'),
            'Metod':result.get('method','P/E-historik' if result.get('category')=='Rörelsedrivande bolag' else 'Ej fastställd'),
            'P/B':round(result['pb'],2) if result.get('pb') is not None else None,
            'Rapporterat substansvärde/aktie':result.get('nav'),
            'Substansvärdets datum':result.get('nav_date',''),
            'Källa till substansvärde':result.get('nav_source',''),
            'ROE (%)':round(result['roe'],1) if result.get('roe') is not None else None,
            'Signal':result.get('status','Saknas'),
            'Behöver kompletteras':result.get('missing','' if result.get('nav') is not None else 'Rapporterat substansvärde per aktie' if result.get('category')=='Investmentbolag / kapitalförvaltning' else 'EPRA-substansvärde per aktie' if result.get('category')=='Fastigheter' else 'Solvens- och kapitalkrav' if result.get('category')=='Försäkring' else ''),
            'Rabatt mot historisk P/E (%)':round(-result['gap'],1) if result.get('gap') is not None else None,
            'Nuvarande P/E':round(result['pe'],1) if result.get('pe') is not None else None,
            'Historisk median P/E':round(result['median'],1) if result.get('median') is not None else None,
            'Aktiekurs':round(result['price'],2) if result.get('price') is not None else None,
            'Indikativ kurs vid historisk P/E':round(result['indicative_price'],2) if result.get('indicative_price') is not None else None,
            'Valuta':result.get('currency',''),
            'P/E-observationer':result.get('observations',0)
        })
    scan_df=pd.DataFrame(scan_rows)
    filter_choice=st.radio('Visa', ['Intressanta','Alla resultat','Andra värderingsmetoder','Saknade uppgifter'],horizontal=True,key='screener_filter')
    if filter_choice=='Intressanta':
        scan_df=scan_df[scan_df['Signal']=='Intressant att analysera vidare']
    elif filter_choice=='Andra värderingsmetoder':
        scan_df=scan_df[scan_df['Värderingskategori']!='Rörelsedrivande bolag']
    elif filter_choice=='Saknade uppgifter':
        scan_df=scan_df[scan_df['Signal'].isin(['Otillräckliga börsdata','P/E-historik saknas','Kunde inte hämtas','Ej klassificerad','Otillräckliga sektordata','Substansvärde krävs','Fastighets-NAV krävs'])]
    scan_df=scan_df.sort_values('Rabatt mot historisk P/E (%)',ascending=False,na_position='last')
    st.dataframe(scan_df,hide_index=True,use_container_width=True)
    st.caption('För investmentbolag krävs rapporterat substansvärde per aktie för substansrabatt. För fastigheter krävs rapporterat EPRA-substansvärde. P/B är inte samma sak som substansrabatt. Banker och försäkringsbolag visas med P/B och ROE, men får ingen automatisk köp-/säljsignal utan kapital- och riskkontroll.')
    st.caption('Beräkningen bygger på senaste hämtade Yahoo-data. Positiv rabatt = lägre nuvarande P/E än historisk median. P/E-historik och aktuell P/E kan ha olika definitioner eller perioder; resultaten måste granskas. Saknade uppgifter ger ingen investeringssignal.')
    st.download_button('Exportera sökresultat (CSV)',data=scan_df.to_csv(index=False).encode('utf-8-sig'),file_name='equitylab_aktiescreener.csv',mime='text/csv')
else:
    st.info('Klicka på ”Sök igenom alla bolag” för att börja. Resultaten visas allt eftersom bolagen kontrolleras.')
if st.session_state.screener_running:
    st.rerun()
st.divider()

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
# Den äldre historiska poängmodellen visas enbart som jämförelse under Analys om bolaget.

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
# EV/EBITA: ROIC som huvudkälla, Yahoo årsbalans som verifierad reserv.
# EBITA, nettoskuld och valutakurs måste vara jämförbara; saknat värde är inte noll.
ev_ebita = None
ev_ebita_source = 'Jämförbar EBITA, nettoskuld eller valuta saknas'
roic_balance = _roic_data.get('balance',[])
def _record_number(record, aliases):
    if not isinstance(record,dict): return None
    for key in aliases:
        value=numeric(record.get(key))
        if value is not None: return value
    return None
def _annual_balance_value(frame, labels, period):
    if frame is None or frame.empty or not period: return None
    for column in frame.columns:
        try:
            if pd.Timestamp(column).date()!=pd.Timestamp(period).date(): continue
        except (ValueError,TypeError,OverflowError): continue
        for label in labels:
            if label in frame.index:
                value=numeric(frame.loc[label,column])
                if value is not None: return value
    return None
if _roic_income and market_cap is not None and market_cap>0 and quote_ccy:
    inc=_roic_income[0]
    report_date=str(inc.get('period_end_date') or '')[:10]
    report_currency=inc.get('currency')
    ebita_annual=_record_number(inc,('ebita',))
    if report_date and report_currency and ebita_annual is not None and ebita_annual>0:
        balance_match=next((x for x in roic_balance if str(x.get('period_end_date') or '')[:10]==report_date and x.get('currency')==report_currency),None)
        debt_annual=_record_number(balance_match,(
            'bs_total_debt','total_debt','totalDebt','bs_short_long_term_debt_total',
            'bs_long_term_debt','bs_total_borrowings'))
        cash_annual=_record_number(balance_match,(
            'bs_cash_and_cash_equivalents','cash_and_cash_equivalents',
            'cash_and_short_term_investments','bs_cash_cash_equivalents_and_short_term_investments',
            'bs_cash_and_short_term_investments','cash_and_cash_equivalents_and_short_term_investments'))
        debt_source='ROIC årsbalans'
        # Yahoo-balansräkning används bara när rapportvalutan uttryckligen är densamma.
        yahoo_balance_currency=report_ccy
        if yahoo_balance_currency==report_currency:
            if debt_annual is None:
                debt_annual=_annual_balance_value(balance,(
                    'Total Debt','Long Term Debt And Capital Lease Obligation'),report_date)
                if debt_annual is not None: debt_source='ROIC EBITA + Yahoo årsbalans'
            if cash_annual is None:
                cash_annual=_annual_balance_value(balance,(
                    'Cash And Cash Equivalents',
                    'Cash Cash Equivalents And Short Term Investments'),report_date)
                if cash_annual is not None: debt_source='ROIC EBITA + Yahoo årsbalans'
        rate=fx_rate(quote_ccy,report_currency)
        if debt_annual is not None and cash_annual is not None and rate is not None and rate>0:
            ev_ebita=(market_cap*rate+debt_annual-cash_annual)/ebita_annual
            ev_ebita_source=debt_source+' '+report_date+'; '+quote_ccy+' omräknat till '+report_currency
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
# EV/EBITA visar endast verifierad beräkning; källstatus vid saknad data.
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

@st.cache_data(ttl=86400,show_spinner=False)
def swedish_ten_year_yield():
    """Senast publicerad svensk tioårig statsobligationsränta, i procent."""
    try:
        response=requests.get('https://api.riksbank.se/swea/v1/Observations/Latest/SEGVB10YC',timeout=12)
        response.raise_for_status()
        payload=response.json()
        records=payload if isinstance(payload,list) else [payload]
        for record in records:
            if not isinstance(record,dict): continue
            raw=record.get('value',record.get('Value'))
            value=numeric(raw)
            date=str(record.get('date',record.get('Date','')))
            if value is not None and -5<value<20 and date:
                try:
                    if (pd.Timestamp.now()-pd.Timestamp(date).tz_localize(None)).days>14:
                        return None,None
                except (ValueError,TypeError):
                    return None,None
                return float(value),date
    except (requests.RequestException,ValueError,TypeError,KeyError):
        pass
    return None,None

@st.cache_data(ttl=86400,show_spinner=False)
def calculated_stock_beta(ticker):
    """Veckobeta mot OMX Stockholm 30, minst 52 gemensamma veckor."""
    try:
        stock=yf.Ticker(ticker).history(period='3y',interval='1wk',auto_adjust=True)
        index=yf.Ticker('^OMX').history(period='3y',interval='1wk',auto_adjust=True)
        if stock.empty or index.empty: return None
        a=stock['Close'].copy()
        b=index['Close'].copy()
        a.index=pd.to_datetime(a.index,utc=True).tz_localize(None).to_period('W').to_timestamp()
        b.index=pd.to_datetime(b.index,utc=True).tz_localize(None).to_period('W').to_timestamp()
        weekly=pd.concat([a.groupby(level=0).last().rename('stock'),b.groupby(level=0).last().rename('market')],axis=1).dropna()
        returns=weekly.pct_change().dropna()
        if len(returns)<52: return None
        variance=returns['market'].var()
        if variance<=0: return None
        result=float(returns['stock'].cov(returns['market'])/variance)
        return result if 0<result<5 else None
    except (ValueError,TypeError,KeyError,OverflowError):
        return None
    except Exception:
        return None

def directional_indicator(good, label):
    """Visuell bedömningssymbol, inte en historisk kursgraf."""
    color='#17834b' if good else '#c33843'
    points='2,24 22,19 42,21 62,11 82,5' if good else '2,5 22,11 42,9 62,19 82,24'
    return ('<div style="display:flex;align-items:center;gap:12px;margin:7px 0 14px;">'
            '<svg viewBox="0 0 86 30" width="95" height="34" role="img" aria-label="'+('Positiv' if good else 'Negativ')+' bedömningsindikator">'
            '<polyline points="'+points+'" fill="none" stroke="'+color+'" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/>'
            '</svg><span style="font-size:0.88rem;color:'+color+';font-weight:600;">'+label+'</span></div>')

# Gemensam värderingsindikator ovanför alla flikar.
# Baseras på verifierbar års-EPS och historisk P/E-median, aldrig enbart dagens P/E.
global_pe_samples=[]
global_pe_frame=extended.get('valuation_history',pd.DataFrame())
if isinstance(global_pe_frame,pd.DataFrame) and not global_pe_frame.empty:
    for pe_label in ('PeRatio','PE Ratio','Trailing P/E','Trailing PE','peRatio','trailingPE','TrailingPeRatio'):
        if pe_label in global_pe_frame.index:
            global_pe_samples+=pd.to_numeric(global_pe_frame.loc[pe_label],errors='coerce').dropna().tolist()
        if pe_label in global_pe_frame.columns:
            global_pe_samples+=pd.to_numeric(global_pe_frame[pe_label],errors='coerce').dropna().tolist()
global_pe_samples=[float(x) for x in global_pe_samples if 0<float(x)<=80]
global_pe_median=float(pd.Series(global_pe_samples).median()) if len(global_pe_samples)>=3 else None
global_eps=None
global_eps_period=None
if _roic_income and quote_ccy:
    latest_annual=_roic_income[0]
    annual_ccy=latest_annual.get('currency')
    for eps_key in ('eps_diluted','eps_basic','eps','is_eps_diluted','is_eps_basic','is_earnings_per_share_diluted','is_earnings_per_share_basic'):
        candidate=numeric(latest_annual.get(eps_key))
        if candidate is not None and candidate>0:
            global_eps=candidate
            global_eps_period=str(latest_annual.get('period_end_date') or '')
            break
    if global_eps is not None:
        eps_fx=fx_rate(annual_ccy,quote_ccy)
        global_eps=global_eps*eps_fx if eps_fx is not None and eps_fx>0 else None
# Värderingskontroll: jämför nuvarande P/E med observerad historisk P/E.
# Ingen DCF kombineras innan kassaflödesdefinitionen och terminalvärdet har validerats.
global_fair=None
global_dcf_fair=None
global_combined=None
global_display_value=None
global_display_method='Historisk P/E-jämförelse'
st.markdown('### Värderingskontroll')
valid_price=price is not None and price>0
valid_current_pe=pe is not None and 0<pe<=80
valid_history=global_pe_median is not None and len(global_pe_samples)>=3
if valid_price and valid_current_pe and valid_history:
    # Jämför multiplar på samma skala. Detta kräver inte osäkra omräkningar
    # av historisk års-EPS till aktuell aktiekurs.
    relative_gap=(pe/global_pe_median-1)*100
    indicative_value=price*global_pe_median/pe
    global_fair=indicative_value
    global_display_value=indicative_value
    status='Dyrare än historiskt' if relative_gap>0 else 'Billigare än historiskt' if relative_gap<0 else 'I nivå med historiken'
    v1,v2,v3,v4=st.columns(4)
    v1.metric('Aktuell aktiekurs',show(price,' '+str(quote_ccy),2))
    v2.metric('Nuvarande P/E',show(pe,'×',1))
    v3.metric('Historisk median P/E',show(global_pe_median,'×',1))
    v4.metric(status,show(abs(relative_gap),' %',1))
    if relative_gap>0:
        st.caption('**Tolkning:** Aktien är '+show(abs(relative_gap),' %',1)+' dyrare sett till P/E än sin historiska nivå. Det kan tyda på relativ övervärdering, men bevisar inte att aktien är övervärderad.')
    elif relative_gap<0:
        st.caption('**Tolkning:** Aktien är '+show(abs(relative_gap),' %',1)+' billigare sett till P/E än sin historiska nivå. Det kan tyda på relativ undervärdering, men bevisar inte att aktien är undervärderad.')
    else:
        st.caption('**Tolkning:** Aktien värderas till samma P/E som sin historiska median.')
    # Beslutsstöd med uttryckliga, försiktiga trösklar.
    # Lägre P/E än historiken är inte tillräckligt för ett köpbeslut.
    st.markdown('#### Preliminär investeringssignal')
    if relative_gap<=-20:
        st.success('**INTRESSANT ATT ANALYSERA VIDARE** – Aktien handlas minst 20 % under sin historiska P/E-median. Kontrollera vinstkvalitet, risker och framtidsutsikter innan köp.')
    elif relative_gap>=20:
        st.error('**HÖGT VÄRDERAD – AVVAKTA NYA KÖP** – Aktien handlas minst 20 % över sin historiska P/E-median. En hög premie kan vara motiverad av bättre tillväxt eller lönsamhet.')
    else:
        st.info('**AVVAKTA – INGET TYDLIGT KÖPLÄGE** – Aktien handlas nära sin historiska P/E-median (inom ±20 %). Värderingen ger ingen tydlig säkerhetsmarginal.')
    st.caption('Signalen bygger endast på P/E i förhållande till historiken. Den är inte ett personligt köp-/säljråd och tar ännu inte hänsyn till prognoser, bolagsrisk eller portföljinnehav.')
    with st.expander('Visa metod, indikativt pris och säkerhetsmarginal'):
        st.write('**Indikativt pris vid historisk P/E:** '+show(indicative_value,' '+str(quote_ccy),2))
        st.write('**Pris med 20 % säkerhetsmarginal:** '+show(indicative_value*0.8,' '+str(quote_ccy),2))
        st.write('**Antal historiska P/E-observationer:** '+str(len(global_pe_samples)))
        st.caption('Indikativt pris = dagens kurs × historisk median-P/E / aktuellt P/E. Beräkningen antar oförändrad vinstnivå och jämförbara P/E-definitioner. Ingen prognos för aktiekursen.')
        st.caption('P/E kan påverkas av engångsposter och skillnader i perioddefinition. Kontrollera bolagets rapporter innan investeringsbeslut.')
else:
    v1,v2,v3=st.columns(3)
    v1.metric('Aktuell aktiekurs',show(price,' '+str(quote_ccy),2) if valid_price else 'Saknas')
    v2.metric('Nuvarande P/E',show(pe,'×',1) if valid_current_pe else 'Saknas')
    v3.metric('Historisk P/E-median',show(global_pe_median,'×',1) if valid_history else 'Saknas')
    missing=[]
    if not valid_price: missing.append('aktuell aktiekurs')
    if not valid_current_pe: missing.append('positivt aktuellt P/E')
    if not valid_history: missing.append('minst tre historiska P/E-observationer')
    st.info('**INGEN INVESTERINGSSIGNAL** – Otillräckligt underlag. Saknas: '+', '.join(missing)+'.')
st.caption('Detta är en historisk multipeljämförelse, inte en fullständig bolagsvärdering. DCF och sammanvägd köp-/säljsignal är pausade tills de har validerats.')
st.divider()

one,two,three,four=st.tabs(['Kurs och historik','Finansiella rapporter','Investeringsanalys','Bolagsvärdering'])
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
        # Visa inte 2021 i kassaflöde/balans när året helt saknar data
        # för de rapportposter som faktiskt presenteras i tabellen.
        if section_name in ('Resultaträkning','Kassaflöde','Balansräkning'):
            visible_labels={key for _,candidates,_ in metrics for key in candidates if key in frame.index}
            cols=[col for col in cols if not (
                str(pd.Timestamp(col).year if not pd.isna(pd.to_datetime(col,errors='coerce')) else col) == '2021'
                and not any(numeric(frame.at[key,col]) is not None for key in visible_labels)
            )]
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
    st.markdown('#### Kvartalsutveckling – rapporterade siffror')
    st.caption('Senaste tillgängliga kvartal från Yahoo Finance. Jämför samma kvartal föregående år när sådana uppgifter finns. Belopp i miljoner av rapportvalutan.')
    quarterly_income=extended.get('quarterly_income',pd.DataFrame())
    quarterly_cash=extended.get('quarterly_cash',pd.DataFrame())
    quarterly_balance=extended.get('quarterly_balance',pd.DataFrame())
    quarterly_definitions=[
        ('Omsättning',quarterly_income,('Total Revenue','Operating Revenue'),'M'),
        ('EBIT',quarterly_income,('Operating Income','EBIT'),'M'),
        ('Nettoresultat',quarterly_income,('Net Income','Net Income Common Stockholders'),'M'),
        ('Resultat per aktie',quarterly_income,('Diluted EPS','Basic EPS'),'EPS'),
        ('Operativt kassaflöde',quarterly_cash,('Operating Cash Flow','Cash Flow From Continuing Operating Activities'),'M'),
        ('Fritt kassaflöde',quarterly_cash,('Free Cash Flow',),'M'),
        ('Räntebärande skulder',quarterly_balance,('Total Debt',),'M'),
        ('Likvida medel',quarterly_balance,('Cash And Cash Equivalents','Cash Cash Equivalents And Short Term Investments'),'M'),
    ]
    # Komplettera enbart saknade årsjämförelser med ROIC:s rapporterade kvartal.
    roic_quarters=[]
    roic_quarter_error=None
    if roic_key() and _roic_data.get('symbol'):
        roic_quarters,roic_quarter_error=roic_quarterly_income(_roic_data['symbol'],roic_key())
    roic_quarter_fields={
        'Omsättning':('is_sales_revenue_turnover','revenue','total_revenue'),
        'EBIT':('is_operating_income','operating_income','ebit'),
        'Nettoresultat':('is_net_income','net_income','net_income_common_stockholders'),
        'Resultat per aktie':('eps_diluted','eps_basic','eps','is_eps_diluted','is_eps_basic')
    }
    quarterly_rows=[]
    for label,frame,aliases,unit in quarterly_definitions:
        if frame is None or frame.empty: continue
        matched=next((k for k in aliases if k in frame.index),None)
        if matched is None: continue
        series=pd.to_numeric(frame.loc[matched],errors='coerce').dropna()
        if series.empty: continue
        try:
            series.index=pd.to_datetime(series.index)
            series=series.sort_index(ascending=False)
        except (ValueError,TypeError):
            continue
        latest_date=series.index[0]
        latest_value=float(series.iloc[0])
        # Samma räkenskapskvartal föregående år; välj närmaste periodslut,
        # inte godtyckligt första datum i ett brett intervall.
        target_date=latest_date-pd.DateOffset(years=1)
        candidates=series.iloc[1:]
        candidates=candidates[abs((candidates.index-target_date).days)<=45]
        prior_date=None
        prior_value=None
        if not candidates.empty:
            prior_date=min(candidates.index,key=lambda d:abs((d-target_date).days))
            prior_value=float(candidates.loc[prior_date])
        yoy=None
        reason=''
        yoy_source='Yahoo Finance'
        # ROIC får endast fylla en lucka om både period och valuta kan verifieras.
        if prior_value is None and label in roic_quarter_fields and roic_quarters and report_ccy:
            def roic_quarter_value(record):
                if record.get('currency')!=report_ccy: return None
                for field in roic_quarter_fields[label]:
                    value=numeric(record.get(field))
                    if value is not None: return value
                return None
            current_match=next((record for record in roic_quarters
                if abs((pd.Timestamp(record['period_end_date'])-latest_date).days)<=45
                and roic_quarter_value(record) is not None),None)
            past_match=next((record for record in roic_quarters
                if abs((pd.Timestamp(record['period_end_date'])-target_date).days)<=45
                and roic_quarter_value(record) is not None),None)
            if current_match is not None and past_match is not None:
                latest_value=roic_quarter_value(current_match)
                prior_value=roic_quarter_value(past_match)
                prior_date=pd.Timestamp(past_match['period_end_date'])
                latest_date=pd.Timestamp(current_match['period_end_date'])
                yoy_source='ROIC AI · rapporterade kvartal'
        if prior_value is None:
            reason='Kvartalsdata för samma period föregående år saknas – se separat helårsjämförelse nedan'
        elif label in ('Räntebärande skulder','Likvida medel'):
            reason='Balanspost: procentuell förändring ska inte tolkas som automatisk förbättring eller försämring'
        elif prior_value<=0:
            reason='Procentuell tillväxt kan bli missvisande när jämförelsevärdet är noll eller negativt'
        else:
            yoy=(latest_value/prior_value-1)*100
        # Kvartal mot närmast föregående kvartal med redan hämtad data.
        previous_quarter_date=None
        previous_quarter_value=None
        previous_quarter_change=None
        q_target=latest_date-pd.DateOffset(months=3)
        q_candidates=series.iloc[1:]
        q_candidates=q_candidates[abs((q_candidates.index-q_target).days)<=40]
        if not q_candidates.empty:
            previous_quarter_date=min(q_candidates.index,key=lambda d:abs((d-q_target).days))
            previous_quarter_value=float(q_candidates.loc[previous_quarter_date])
            if previous_quarter_value>0:
                previous_quarter_change=(latest_value/previous_quarter_value-1)*100
        # För balansposter visar vi jämförelse med föregående balansdag,
        # men inte en godtycklig tolkning av bättre eller sämre.
        if label in ('Räntebärande skulder','Likvida medel') and prior_value is not None and prior_value>0:
            yoy=(latest_value/prior_value-1)*100
            reason='Balanspost: förändringen är inte automatiskt positiv eller negativ'
        quarterly_rows.append({
            'Nyckeltal':label,'Senaste kvartal':str(latest_date.date()),
            'Rapporterat värde':show(latest_value/1e6,' M',1) if unit=='M' else show(latest_value,'',2),
            'Föregående kvartal':str(previous_quarter_date.date()) if previous_quarter_date is not None else 'Saknas',
            'Förändring mot föregående kvartal':show(previous_quarter_change,' %',1) if previous_quarter_change is not None else 'Saknas',
            'Jämförelsekvartal':str(prior_date.date()) if prior_date is not None else 'Saknas',
            'Värde föregående år':(show(prior_value/1e6,' M',1) if unit=='M' else show(prior_value,'',2)) if prior_value is not None else 'Saknas',
            'Förändring mot samma kvartal föregående år':show(yoy,' %',1) if yoy is not None else 'Saknas'
        })
    if quarterly_rows:
        quarter_table=pd.DataFrame(quarterly_rows)
        if quarter_table['Förändring mot samma kvartal föregående år'].eq('Saknas').all():
            st.info('Samma kvartal föregående år finns inte i tillgängliga rapportdata. Tabellen visar kvartalsutvecklingen där den går att beräkna; längre ned finns en separat jämförelse mellan två rapporterade helår.')
            quarter_table=quarter_table.drop(columns=['Jämförelsekvartal','Värde föregående år','Förändring mot samma kvartal föregående år'])
        st.dataframe(quarter_table,hide_index=True,use_container_width=True)
        st.caption('QoQ jämför med föregående rapporterade kvartal. YoY jämför med motsvarande kvartal ett år tidigare, endast när båda värdena finns. Helårsförändringen redovisas separat nedan. Ingen saknad kvartalsdata uppskattas.')
    else:
        st.info('Kvartalsrapporter saknas hos den anslutna datakällan för detta bolag.')

    if roic_quarter_error:
        st.caption('Kompletterande ROIC-kvartal: '+roic_quarter_error+'. Årsjämförelser kan fortfarande saknas.')
    elif roic_quarters:
        st.caption('ROIC AI levererade '+str(len(roic_quarters))+' rapporterade kvartal. Dessa används endast när motsvarande Yahoo-jämförelse saknas och valuta och period överensstämmer.')

    st.markdown('#### Årsförändring – beräknad från befintliga årsrapporter')
    st.caption('Förändring mellan två rapporterade helår, beräknad från de faktiska årsredovisningsposterna. Detta är inte en jämförelse mellan enskilda kvartal.')
    annual_comparisons=[
        ('Omsättning',fin,('Total Revenue','Operating Revenue'),'M'),
        ('EBIT',fin,('Operating Income','EBIT'),'M'),
        ('Nettoresultat',fin,('Net Income','Net Income Common Stockholders'),'M'),
        ('EPS',fin,('Diluted EPS','Basic EPS'),'EPS'),
        ('Operativt kassaflöde',cash,('Operating Cash Flow','Cash Flow From Continuing Operating Activities'),'M'),
        ('Fritt kassaflöde',cash,('Free Cash Flow',),'M'),
        ('Räntebärande skulder',balance,('Total Debt',),'M'),
        ('Likvida medel',balance,('Cash And Cash Equivalents','Cash Cash Equivalents And Short Term Investments'),'M')
    ]
    annual_rows=[]
    for label,annual_frame,aliases,unit in annual_comparisons:
        if annual_frame is None or annual_frame.empty: continue
        key=next((k for k in aliases if k in annual_frame.index),None)
        if key is None: continue
        vals=pd.to_numeric(annual_frame.loc[key],errors='coerce').dropna()
        try:
            vals.index=pd.to_datetime(vals.index)
            vals=vals.sort_index(ascending=False)
        except (ValueError,TypeError):
            continue
        if len(vals)<2: continue
        current_date=vals.index[0]
        earlier_candidates=vals.iloc[1:]
        earlier_candidates=earlier_candidates[abs((earlier_candidates.index-(current_date-pd.DateOffset(years=1))).days)<=45]
        if earlier_candidates.empty: continue
        older_date=min(earlier_candidates.index,key=lambda d:abs((d-(current_date-pd.DateOffset(years=1))).days))
        current=float(vals.iloc[0])
        older=float(earlier_candidates.loc[older_date])
        growth=(current/older-1)*100 if older>0 else None
        annual_rows.append({
            'Nyckeltal':label,
            'Senaste helår':str(current_date.date()),
            'Värde senaste helår':show(current/1e6,' M',1) if unit=='M' else show(current,'',2),
            'Föregående helår':str(older_date.date()),
            'Värde föregående helår':show(older/1e6,' M',1) if unit=='M' else show(older,'',2),
            'Förändring helår':show(growth,' %',1) if growth is not None else 'Ej meningsfull'
        })
    if annual_rows:
        st.dataframe(pd.DataFrame(annual_rows),hide_index=True,use_container_width=True)
        st.caption('Årsförändringen beräknas som (senaste helår / föregående helår − 1) × 100. En negativ eller nollstor jämförelsebas ger ingen vanlig procentuell tillväxt.')
    else:
        st.caption('Även två jämförbara årsperioder saknas för de tillgängliga posterna.')

    st.markdown('#### Fullständiga kvartalsrapporter')
    st.caption('Samtliga tillgängliga originalposter från dataleverantören visas nedan, inte bara de utvalda nyckeltalen i sammanfattningen. Rapportperioder och saknade värden bevaras.')
    for report_title,report_frame in [
        ('Resultaträkning – kvartal',quarterly_income),
        ('Kassaflödesanalys – kvartal',quarterly_cash),
        ('Balansräkning – kvartal',quarterly_balance)
    ]:
        with st.expander(report_title):
            if not isinstance(report_frame,pd.DataFrame) or report_frame.empty:
                st.info('Inga kvartalsuppgifter levererades för denna rapport. Det betyder inte att bolaget saknar offentlig kvartalsrapport.')
            else:
                raw_quarters=report_frame.copy()
                raw_quarters.columns=[str(pd.Timestamp(c).date()) if not pd.isna(pd.to_datetime(c,errors='coerce')) else str(c) for c in raw_quarters.columns]
                raw_quarters=raw_quarters.reset_index().rename(columns={'index':'Rapportpost'})
                st.dataframe(raw_quarters,use_container_width=True,hide_index=True)
                st.caption(str(len(report_frame.index))+' rapportposter · '+str(len(report_frame.columns))+' rapportperioder · Yahoo Finance. Belopp visas i originalenheter (inte miljoner).')
                st.download_button('Ladda ner '+report_title+' (CSV)',data=raw_quarters.to_csv(index=False).encode('utf-8-sig'),file_name=ticker.replace('.','_')+'_'+('resultat' if 'Resultat' in report_title else 'kassaflode' if 'Kassaflöde' in report_title else 'balans')+'_kvartal.csv',mime='text/csv',key='quarter_export_'+report_title)
    st.caption('Kvartalsrapporter kan innehålla kumulativa kassaflöden beroende på datakällans definition. Summera inte kvartalsvärden utan att kontrollera att de är fristående kvartal.')
    st.markdown('#### Datatäckning och rapporternas kvalitet')
    coverage_frames=[('Årsresultat',fin),('Årskassaflöde',cash),('Årsbalans',balance),
                     ('Kvartalsresultat',quarterly_income),('Kvartalskassaflöde',quarterly_cash),('Kvartalsbalans',quarterly_balance)]
    coverage=[]
    for label,frame in coverage_frames:
        count=len(frame.columns) if isinstance(frame,pd.DataFrame) and not frame.empty else 0
        coverage.append({'Rapport':label,'Tillgängliga perioder':count})
    st.dataframe(pd.DataFrame(coverage),hide_index=True,use_container_width=True)
    st.caption('Datatäckning visar vilka rapportserier som har hämtats, inte att varje post är reviderad eller fullständigt kvalitetssäkrad. Kontrollera viktiga värden mot bolagets publicerade rapporter.')

with three:
    st.subheader(f'Investeringsanalys – {company}')
    st.markdown('### Investeringsprocess i fem steg')
    st.caption('Följ bolaget från affärsmodell och finansiell kvalitet till bransch, värdering och uppföljning. Kontrollera underlaget innan en investeringsbedömning görs.')
    with st.container(border=True):
        st.markdown('**01 · Investeringsram**')
        st.write('Tidshorisont: fem år i prognosmodellen. En rimlig säkerhetsmarginal och risktolerans behöver alltid bedömas utifrån investerarens situation.')
        st.caption('Investeringshorisonten är modellens, inte ett personligt råd.')
    with st.container(border=True):
        st.markdown('**02 · Bolagets kvalitet**')
        q1,q2,q3=st.columns(3)
        q1.metric('Omsättningstillväxt',show(sales_growth_pct,' %/år',1) if sales_growth_pct is not None else 'Saknas')
        q2.metric('Rörelsemarginal',show(margin*100,' %',1) if margin is not None else 'Saknas')
        q3.metric('Nettoskuld / EBITDA',show(netdebt_ebitda,'×',2) if netdebt_ebitda is not None else 'Saknas')
        st.caption('Tillväxt, lönsamhet och skuldsättning är separata observationer. Ett enskilt nyckeltal avgör inte om bolaget är en bra investering.')
        if opcf is not None:
            st.write('Operativt kassaflöde: **'+show(opcf/1e6,' M',1)+'** (kontrollera rapportperiod under Finansiella rapporter).')
        else:
            st.write('Operativt kassaflöde saknas i hämtat underlag.')
    with st.container(border=True):
        st.markdown('**03 · Affärsmodell och konkurrens**')
        st.write('Bransch: **'+str(info.get('industry') or 'Ej klassificerad')+'** · Sektor: **'+str(info.get('sector') or 'Ej klassificerad')+'**')
        description=str(info.get('longBusinessSummary') or '').strip()
        if description:
            st.write(description[:650]+('…' if len(description)>650 else ''))
        else:
            st.caption('Bolagsbeskrivning saknas från datakällan.')
        st.markdown('**Jämförelse med börsnoterade konkurrenter**')
        peers=PEER_UNIVERSE.get(ticker,[])
        if peers:
            peer_rows=[{'Bolag':company,'P/E':pe,'Rörelsemarginal (%)':margin*100 if margin is not None else None,
                        'Omsättningstillväxt (%)':sales_growth_pct,'Nettoskuld/EBITDA':netdebt_ebitda,
                        'Jämförbarhet':'Bolaget som analyseras'}]
            for peer_name,peer_ticker,peer_note in peers:
                peer_data=peer_snapshot(peer_ticker)
                peer_rows.append({'Bolag':peer_name,'P/E':peer_data.get('pe'),
                    'Rörelsemarginal (%)':peer_data.get('margin'),
                    'Omsättningstillväxt (%)':peer_data.get('growth'),
                    'Nettoskuld/EBITDA':peer_data.get('leverage'),
                    'Jämförbarhet':peer_note})
            peer_table=pd.DataFrame(peer_rows)
            st.dataframe(peer_table,hide_index=True,use_container_width=True,
                column_config={'P/E':st.column_config.NumberColumn('P/E',format='%.1f×'),
                'Rörelsemarginal (%)':st.column_config.NumberColumn('Rörelsemarginal (%)',format='%.1f'),
                'Omsättningstillväxt (%)':st.column_config.NumberColumn('Omsättningstillväxt (%)',format='%.1f'),
                'Nettoskuld/EBITDA':st.column_config.NumberColumn('Nettoskuld/EBITDA',format='%.2f×')})
            available_pe=[row['P/E'] for row in peer_rows[1:] if row['P/E'] is not None and row['P/E']>0]
            if len(available_pe)>=2 and pe is not None and pe>0:
                peer_median=float(pd.Series(available_pe).median())
                st.write('Bolagets P/E: **'+show(pe,'×',1)+'** · Jämförelsebolagens median-P/E: **'+show(peer_median,'×',1)+'**.')
            else:
                st.caption('För få jämförbara P/E-observationer för en meningsfull median.')
            st.caption('Urvalet är manuellt granskat för branschlikhet men inte identiska affärsmodeller. Yahoo-uppgifterna kan avse olika rapportperioder, redovisningsprinciper och valutor. Jämförelsen är vägledande, inte en automatisk riktkurs eller investeringssignal.')
        else:
            st.caption('Inga kvalitetssäkrade börsnoterade jämförelsebolag har valts för detta bolag ännu. Vi visar inte godtyckliga konkurrenter.')
        st.caption('Marknadsandelar och konkurrensfördelar kräver fortfarande separata branschkällor.')
    with st.container(border=True):
        st.markdown('**04 · Branschutsikter och risker**')
        st.write('Granska reglering, konjunktur, konkurrens, geografisk exponering och hur dessa faktorer påverkar tillväxt och kassaflöde.')
        st.caption('Dessa risker kräver verifierade bolagsrapporter och branschkällor. EquityLab saknar ännu ett kvalitetssäkrat, automatiskt branschunderlag och ger därför ingen poäng.')
    with st.container(border=True):
        st.markdown('**05 · Värdering och uppföljning**')
        st.write('Aktuell kurs: **'+(show(price,' '+str(quote_ccy),2) if price is not None else 'Saknas')+'**.')
        st.write('Kontrollerat motiverat värde: **Ej fastställt**. Tidigare automatiska DCF- och rabattbedömningar är inte godkända som investeringssignal.')
        st.caption('Följ kvartalsrapporter, prognosförändringar, kassaflöde och värderingsantaganden. Använd fliken Bolagsvärdering och bolagens publicerade rapporter för underlaget.')
    st.markdown('#### Samlad bedömning')
    st.info('Analysen är ännu inte tillräckligt verifierad för en köp-/säljbedömning. Finansiell historik finns delvis, men branschbedömning och kvalitetssäkrad värdering återstår.')
    st.divider()

    st.markdown('### Vad kan aktien vara värd om fem år?')
    st.caption('Utifrån bolagets rapporterade resultat beräknar modellen ett möjligt framtida aktiepris och vad det skulle innebära för din avkastning.')
    st.markdown('**Så använder du analysen:** Se de tre möjliga utfallen och kontrollera antagandena bakom beräkningen. Du behöver inte fylla i några siffror.')
    def scenario_field(record, keys):
        for key in keys:
            result = numeric(record.get(key))
            if result is not None:
                return result
        return None
    annual = _roic_income
    forecast_currency = annual[0].get('currency') if annual else None
    latest_report = annual[0] if annual else {}
    forecast_revenue = scenario_field(latest_report, ('is_sales_revenue_turnover','revenue','total_revenue'))
    forecast_ebita = scenario_field(latest_report, ('ebita',))
    forecast_period = str(latest_report.get('period_end_date') or 'Okänd')[:10]
    # Utgå endast från rapporterad EBITA och omsättning för samma period och valuta.
    if not (forecast_revenue and forecast_revenue>0 and forecast_ebita is not None and forecast_currency):
        st.info('Framtidsvärdering kräver omsättning, EBITA och rapportvaluta för samma årsperiod från ROIC AI. Inga värden uppskattas när rapportdata saknas.')
    else:
        historical_growth = None
        if len(annual)>1:
            older = annual[1]
            older_revenue = scenario_field(older, ('is_sales_revenue_turnover','revenue','total_revenue'))
            if older_revenue and older_revenue>0 and older.get('currency')==forecast_currency:
                try:
                    years = (pd.Timestamp(forecast_period)-pd.Timestamp(older.get('period_end_date'))).days/365.25
                    if 0.8<=years<=3.2:
                        historical_growth = (forecast_revenue/older_revenue)**(1/years)-1
                except (ValueError,TypeError,OverflowError):
                    pass
        latest_margin = forecast_ebita/forecast_revenue
        suggested_growth = max(-10.0,min(20.0,round(100*historical_growth,1))) if historical_growth is not None else 5.0
        suggested_margin = max(0.0,min(35.0,round(latest_margin*100,1)))
        st.markdown('#### 1. Utgångsläge')
        baseline_a,baseline_b,baseline_c=st.columns(3)
        baseline_a.metric('Omsättning senaste helår',show(forecast_revenue/1e6,' M '+str(forecast_currency),0))
        baseline_b.metric('EBITA-marginal',show(latest_margin*100,' %',1))
        baseline_c.metric('Historisk omsättningstillväxt',show(historical_growth*100,' %/år',1) if historical_growth is not None else 'Saknas')
        st.caption('Rapportperiod: '+forecast_period+' · Historiken avser endast tillgängliga jämförbara årsperioder.')
        # Inga manuella inmatningar: historikbaserade utgångspunkter och öppet redovisade antaganden.
        base_growth=float(suggested_growth)
        base_margin=float(suggested_margin)
        hurdle=10.0
        base_multiple=float(ev_ebita) if ev_ebita is not None and 2.0<=ev_ebita<=40.0 else 12.0
        multiple_source='Nuvarande beräknad EV/EBITA' if ev_ebita is not None and 2.0<=ev_ebita<=40.0 else 'Illustrativ standardmultipel – ingen verifierad historisk värdering'
        st.markdown('#### 2. Så räknar modellen')
        st.write('Modellen använder rapporterad omsättning och EBITA, uppskattar utvecklingen under fem år och värderar därefter bolaget med en EV/EBITA-multipel.')
        with st.expander('Visa beräkningsantaganden'):
            st.write('Årlig omsättningstillväxt: **'+show(base_growth,' %',1)+'**'+(' (historiskt härledd)' if historical_growth is not None else ' (illustrativt antagande – historik saknas)'))
            st.write('EBITA-marginal år fem: **'+show(base_margin,' %',1)+'** (utgår från senast rapporterad marginal)')
            st.write('EV/EBITA år fem: **'+show(base_multiple,'×',1)+'** ('+multiple_source+')')
            st.write('Avkastningskrav: **10 % per år** (modellantagande)')
            st.caption('Negativt och positivt scenario varierar tillväxt ±5 procentenheter, marginal ±2 procentenheter och EV/EBITA ±3×. Antagandena är inte bolagsprognoser.')
        # Nettoskuld måste komma från samma årsbalans som EBITA.
        balance_period = next((x for x in _roic_data.get('balance',[]) if str(x.get('period_end_date') or '')[:10]==forecast_period and x.get('currency')==forecast_currency),None)
        debt_annual = scenario_field(balance_period or {},('bs_total_debt','total_debt','totalDebt'))
        cash_annual = scenario_field(balance_period or {},('bs_cash_and_cash_equivalents','cash_and_cash_equivalents','cash_and_short_term_investments'))
        annual_netdebt = debt_annual-cash_annual if debt_annual is not None and cash_annual is not None else None
        market_fx = fx_rate(quote_ccy,forecast_currency) if quote_ccy and forecast_currency else None
        can_value = all(x is not None for x in (annual_netdebt,market_fx,market_cap,price)) and market_cap>0 and price>0 and market_fx>0
        if not can_value:
            st.warning('EBITA-prognosen kan beräknas, men aktievärde och avkastning kräver verifierad nettoskuld från ROIC:s årsbalans, börsvärde, aktiekurs och valutakurs.')
        else:
            st.caption('Värderingen antar oförändrad nominell nettoskuld under fem år, oförändrat antal aktier och ingen utdelning. Den inkluderar inte förvärvskostnader eller framtida kapitalanskaffningar.')
        scenarios = [
            ('Negativt',max(-25.0,base_growth-5),max(0.0,base_margin-2),max(2.0,base_multiple-3)),
            ('Bas',base_growth,base_margin,base_multiple),
            ('Positivt',min(35.0,base_growth+5),min(50.0,base_margin+2),base_multiple+3),
        ]
        scenario_rows = []
        for scenario_name,growth_pct,margin_pct,multiple in scenarios:
            projected_sales = forecast_revenue*(1+growth_pct/100)**5
            projected_ebita = projected_sales*margin_pct/100
            projected_ev = projected_ebita*multiple
            future_equity = max(0.0,projected_ev-annual_netdebt) if can_value else None
            current_equity_report = market_cap*market_fx if can_value else None
            future_price = price*future_equity/current_equity_report if can_value else None
            annual_return = ((future_price/price)**(1/5)-1)*100 if future_price is not None and future_price>0 else (-100.0 if future_price==0 else None)
            fair_price = future_price/(1+hurdle/100)**5 if future_price is not None else None
            safety_margin = (1-price/fair_price)*100 if fair_price is not None and fair_price>0 else None
            scenario_rows.append({
                'Scenario':scenario_name,
                'Tillväxt/år':f'{growth_pct:.1f} %',
                'EBITA-marginal år 5':f'{margin_pct:.1f} %',
                'EV/EBITA år 5':f'{multiple:.1f}×',
                'EBITA år 5 (M '+str(forecast_currency)+')':round(projected_ebita/1e6,1),
                'Aktiekurs år 5':show(future_price,' '+str(quote_ccy),2) if future_price is not None else 'Saknas',
                'Avkastning/år':show(annual_return,' %',1) if annual_return is not None else 'Saknas',
                'Motiverat pris idag':show(fair_price,' '+str(quote_ccy),2) if fair_price is not None else 'Saknas',
                'Säkerhetsmarginal':show(safety_margin,' %',1) if safety_margin is not None else 'Saknas'
            })
        st.markdown('#### 3. Resultat – tre möjliga utvecklingar')
        if can_value:
            st.caption('Beräknade priser och kursavkastning, exklusive utdelning. Resultaten bygger på de redovisade antagandena ovan.')
            for item in scenario_rows:
                with st.container(border=True):
                    st.markdown('**'+item['Scenario']+' scenario**')
                    sc1,sc2,sc3=st.columns(3)
                    sc1.metric('Aktiekurs om fem år',item['Aktiekurs år 5'])
                    sc2.metric('Årlig kursavkastning',item['Avkastning/år'])
                    sc3.metric('Motiverat pris idag',item['Motiverat pris idag'])
        else:
            st.info('Scenariernas resultat och EBITA kan beräknas, men ett aktiepris kräver verifierad nettoskuld och valuta.')
        with st.expander('Visa detaljerad scenariotabell'):
            st.dataframe(pd.DataFrame(scenario_rows),hide_index=True,use_container_width=True)
        if can_value:
            base_case = scenario_rows[1]
            base_future_ev = forecast_revenue*(1+base_growth/100)**5*(base_margin/100)*base_multiple
            base_equity_report = max(0.0,base_future_ev-annual_netdebt)
            base_fair_price = price*base_equity_report/(market_cap*market_fx)/(1+hurdle/100)**5
            base_return = (base_equity_report/(market_cap*market_fx))**(1/5)-1 if base_equity_report>0 else -1.0
            st.caption('Motiverat pris idag diskonterar värdet om fem år med avkastningskravet. Säkerhetsmarginalen är '+show((1-price/base_fair_price)*100,' %',1)+' i basscenariot.' if base_fair_price>0 else 'Basscenariot ger inget positivt motiverat aktiepris.')
            # Omvänd värdering: vilken omsättningstillväxt behövs för att nå avkastningskravet?
            target_equity = market_cap*market_fx*(1+hurdle/100)**5
            required_ev = target_equity+annual_netdebt
            required_revenue = required_ev/(base_multiple*(base_margin/100)) if base_multiple>0 and base_margin>0 and required_ev>0 else None
            required_growth = ((required_revenue/forecast_revenue)**(1/5)-1)*100 if required_revenue is not None and required_revenue>0 else None
            st.markdown('#### Vad måste bolaget prestera för att nå ditt avkastningskrav?')
            if required_growth is None:
                st.info('Kan inte beräknas med dessa antaganden. Kontrollera positiv marginal och framtida företagsvärde.')
            else:
                st.metric('Nödvändig omsättningstillväxt per år',show(required_growth,' %',1))
                if historical_growth is not None:
                    difference=required_growth-historical_growth*100
                    st.write('Krävd tillväxt ligger '+show(abs(difference),' procentenheter ',1)+('över' if difference>0 else 'under eller i nivå med')+' den senast uppmätta historiska tillväxttakten.')
                st.caption('Omvänd beräkning håller EBITA-marginal, exitmultipel och nettoskuld oförändrade enligt basscenariot. Den är inte en prognos.')
            st.markdown('#### Möjligheter och risker')
            if historical_growth is not None and historical_growth>0:
                st.write(f'Förutsättning: omsättningen har ökat med {historical_growth*100:.1f} % per år mellan de två senaste jämförbara årsrapporterna. Det är historik, inte en garanti för framtida tillväxt.')
            if latest_margin>0:
                st.write(f'Förutsättning: senaste rapporterade EBITA-marginal är {latest_margin*100:.1f} %. Scenarierna prövar hur marginalförändringar påverkar aktievärdet.')
            st.write('Risk: en lägre tillväxt, pressade marginaler eller lägre framtida värderingsmultipel kan ge betydligt sämre avkastning. Modellen antar dessutom oförändrad nettoskuld.')
            st.caption('Scenarierna är känslighetsberäkningar, inte prognoser. Oförändrad nettoskuld och antal aktier antas; utdelning och framtida finansieringsbehov ingår inte.')
    st.divider()
    st.markdown('#### Historisk poängmodell – äldre jämförelse')
    st.caption('Den äldre färgmodellen är inte kalibrerad för framtida tillväxt och används inte för scenariovärderingen ovan.')
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
    st.markdown('#### Äldre investeringsbedömning – historiska gränsvärden')
    st.write(_signal_title)
    if _assessment_score is not None:
        st.metric('Modellpoäng',f'{_assessment_score*100:.0f} av 100')
    st.caption(_signal_detail)
    # Sammanfattad investeringsanalys med begripliga enheter och tydlig motivering.
    if _assessment_rows:
        st.markdown('##### Vad talar för och emot bolaget?')
        factor_explanations = {
            'Omsättning CAGR': ('Omsättningstillväxt','Omsättningen har vuxit över den analyserade perioden.'),
            'Nettovinst CAGR': ('Vinsttillväxt','Nettovinsten har förändrats över den analyserade perioden.'),
            'Senaste EBIT-marginal': ('Rörelsemarginal','Visar hur stor del av omsättningen som blir rörelseresultat.'),
            'Marginalförändring': ('Förändrad rörelsemarginal','Visar hur lönsamheten har utvecklats mellan rapportåren.'),
            'Andel positiva kassaflödesår': ('Operativt kassaflöde','Andel analyserade år med positivt operativt kassaflöde.'),
            'Andel positiva FCF-år': ('Fritt kassaflöde','Andel analyserade år med positivt fritt kassaflöde.'),
            'ROE senaste år': ('Avkastning på eget kapital','Visar avkastningen på aktieägarnas redovisade kapital.'),
            'FCF-yield senaste år': ('Fritt kassaflöde i förhållande till börsvärde','Senaste årets fria kassaflöde jämfört med börsvärdet.'),
            'P/E': ('P/E-tal','Aktiekursens värdering i förhållande till vinst per aktie.'),
            'Nettoskuld/EBITDA': ('Skuldsättning','Nettoskuld i förhållande till EBITDA.')
        }
        def factor_display(item):
            raw_name = str(item.get('Faktor','Nyckeltal'))
            label, description = factor_explanations.get(raw_name,(raw_name,str(item.get('Metod',''))))
            raw_value = numeric(item.get('Värde'))
            if raw_value is None:
                formatted = 'Saknas'
            elif raw_name in ('Omsättning CAGR','Nettovinst CAGR','Senaste EBIT-marginal','Marginalförändring','Andel positiva kassaflödesår','Andel positiva FCF-år','ROE senaste år','FCF-yield senaste år'):
                formatted = show(raw_value*100,' %',1)
                if raw_name=='Marginalförändring':
                    formatted = show(raw_value*100,' procentenheter',1)
            else:
                formatted = show(raw_value,'×',1)
            return {'Faktor':label,'Värde':formatted,'Förklaring':description}
        positives = [factor_display(x) for x in _assessment_rows if x.get('Utfall')=='Positivt']
        negatives = [factor_display(x) for x in _assessment_rows if x.get('Utfall')=='Negativt']
        # Kompakta analyskort utan upprepade rubriker eller stora styckesavstånd.
        from html import escape as html_escape
        def render_factor_cards(factors, empty_text, accent):
            if not factors:
                st.caption(empty_text)
                return
            cards = []
            for factor in factors:
                title = html_escape(str(factor['Faktor']))
                value = html_escape(str(factor['Värde']))
                description = html_escape(str(factor['Förklaring']))
                cards.append(
                    '<div style="border:1px solid #e4e8ef;border-left:3px solid '+accent+';'
                    'border-radius:9px;padding:11px 13px;margin:0 0 8px 0;background:transparent;">'
                    '<div style="display:flex;justify-content:space-between;align-items:baseline;'
                    'gap:12px;flex-wrap:wrap;">'
                    '<span style="font-size:14px;font-weight:600;">'+title+'</span>'
                    '<span style="font-size:16px;font-weight:700;white-space:nowrap;">'+value+'</span>'
                    '</div><div style="font-size:12px;opacity:.72;line-height:1.45;margin-top:3px;">'
                    +description+'</div></div>'
                )
            st.markdown(''.join(cards),unsafe_allow_html=True)
        col_strengths,col_risks=st.columns(2,gap='large')
        with col_strengths:
            st.markdown('**Styrkor**')
            render_factor_cards(positives,'Inga positiva faktorer har kunnat fastställas.','#2c8b69')
        with col_risks:
            st.markdown('**Utmaningar och risker**')
            render_factor_cards(negatives,'Inga negativa faktorer identifierades med tillgängliga data.','#c58a45')
        st.markdown('**Sammanfattande bedömning**')
        if _assessment_score is None:
            st.write('Underlaget räcker inte för ett sammanvägt betyg. De enskilda styrkorna och riskerna visas ändå.')
        else:
            score_pct = _assessment_score*100
            st.progress(min(1.0,max(0.0,_assessment_score)),text='Historisk modellpoäng: '+show(score_pct,' av 100',1))
            debt_warning = netdebt is not None and ebitda is not None and ebitda>0 and netdebt/ebitda>=4
            if 'GRÖN' in _signal_title:
                conclusion = 'Bolaget når modellens gräns för grönt, även om de redovisade riskfaktorerna fortfarande behöver följas.'
            elif debt_warning:
                conclusion = 'Hög skuldsättning utlöser modellens riskspärr. Lägre skuldsättning krävs för att spärren ska upphöra.'
            elif score_pct<75:
                conclusion = 'Modellpoängen ligger '+show(75-score_pct,' poäng',1)+' under gränsen för grönt. Förbättringar i de negativa faktorerna kan höja poängen.'
            else:
                conclusion = 'Modellpoängen är tillräcklig för grönt, men ytterligare riskvillkor behöver kontrolleras.'
            st.write(conclusion)
        st.caption('Detta är den äldre, historiska poängmodellen. Den bygger främst på Yahoo Finance och fasta metodgränser; den nya femårsprognosen visas separat ovan.')
    if 'GRÅ' in _signal_title:
        st.info('Grått betyder att appens anslutna datakällor inte räcker för en investeringsvärdering – inte att bolaget saknar offentliga rapporter.')

    if _assessment_rows:
        with st.expander('Visa detaljerade modellfaktorer'):
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


with four:
    st.subheader(f'Bolagsvärdering – {company}')
    st.caption('Automatiska historikbaserade antaganden och öppet redovisade modellpolicyer. Inga manuella reglage. Resultaten är inte riktkurser eller köprekommendationer.')
    if price is None or price<=0 or not quote_ccy:
        st.warning('Aktuell aktiekurs och handelsvaluta krävs för att jämföra värderingarna.')
    else:
        st.metric('Aktuell aktiekurs',show(price,' '+quote_ccy,2))
        st.markdown('#### 1. P/E-värdering')
        st.caption('Motiverat pris = vinst per aktie × antaget P/E-tal. Års-EPS och dagens kurs måste vara i samma valuta.')
        valuation_eps=None
        valuation_eps_note='Ingen verifierad års-EPS'
        if _roic_income:
            first_report=_roic_income[0]
            eps_report_ccy=first_report.get('currency')
            for field in ('eps_diluted','eps_basic','eps','is_eps_diluted','is_eps_basic','is_earnings_per_share_diluted','is_earnings_per_share_basic'):
                candidate=numeric(first_report.get(field))
                if candidate is not None and candidate>0:
                    valuation_eps=candidate
                    valuation_eps_note='ROIC års-EPS '+str(first_report.get('period_end_date') or '')
                    break
            if valuation_eps is None and fin is not None and not fin.empty and eps_report_ccy and report_ccy==eps_report_ccy:
                report_year=str(first_report.get('period_end_date') or '')[:4]
                for column in fin.columns:
                    try: same_year=str(pd.Timestamp(column).year)==report_year
                    except (ValueError,TypeError): same_year=False
                    if not same_year: continue
                    for label in ('Diluted EPS','Basic EPS'):
                        if label in fin.index:
                            candidate=numeric(fin.loc[label,column])
                            if candidate is not None and candidate>0:
                                valuation_eps=candidate
                                valuation_eps_note='Yahoo års-EPS '+report_year
                                break
                    if valuation_eps is not None: break
            if valuation_eps is not None:
                eps_rate=fx_rate(eps_report_ccy,quote_ccy)
                valuation_eps=valuation_eps*eps_rate if eps_rate is not None and eps_rate>0 else None
                valuation_eps_note+='; omräknad '+str(eps_report_ccy)+' till '+quote_ccy
        elif report_ccy==quote_ccy and eps is not None and eps>0:
            valuation_eps=eps
            valuation_eps_note='Yahoo EPS i handelsvaluta; kontrollera rapportperiod'
        default_pe=max(5.0,min(35.0,round(float(pe),1))) if pe is not None and pe>0 else 15.0
        valuation_frame=extended.get('valuation_history',pd.DataFrame())
        pe_samples=[]
        if isinstance(valuation_frame,pd.DataFrame) and not valuation_frame.empty:
            for label in ('PeRatio','PE Ratio','Trailing P/E','Trailing PE','peRatio','trailingPE','TrailingPeRatio'):
                if label in valuation_frame.index:
                    pe_samples+=pd.to_numeric(valuation_frame.loc[label],errors='coerce').dropna().tolist()
                if label in valuation_frame.columns:
                    pe_samples+=pd.to_numeric(valuation_frame[label],errors='coerce').dropna().tolist()
        pe_samples=[float(v) for v in pe_samples if 0<float(v)<=80]
        pe_multiple=float(pd.Series(pe_samples).median()) if len(pe_samples)>=3 else None
        st.caption('P/E-antagande: median av minst tre observerade historiska P/E-tal. Saknas historik beräknas inget motiverat P/E-värde.')
        if valuation_eps is None or pe_multiple is None:
            st.info('P/E-värdering kräver verifierbar EPS och minst tre historiska P/E-observationer.')
            pe_fair=None
        else:
            pe_fair=valuation_eps*pe_multiple
            c1,c2,c3=st.columns(3)
            c1.metric('Vinst per aktie',show(valuation_eps,' '+quote_ccy,2))
            c2.metric('Motiverat pris enligt P/E',show(pe_fair,' '+quote_ccy,2))
            c3.metric('Skillnad mot kurs',show((pe_fair/price-1)*100,' %',1))
            st.caption('EPS-källa: '+valuation_eps_note+'. Historisk median P/E: '+show(pe_multiple,'×',1)+'. Historisk värdering är ingen garanti för framtida rimligt P/E.')
        st.markdown('#### 2. DCF – diskonterat fritt kassaflöde')
        st.warning('DCF-beräkningen nedan är preliminär och inte kvalitetssäkrad för investeringsbeslut. Den används inte i den övergripande värderingsindikatorn.')
        st.caption('Förenklad aktieägar-DCF: prognostiserat fritt kassaflöde diskonteras till idag, inklusive ett terminalvärde. Modellen kräver kassaflöde och antal aktier i verifierbar valuta.')
        fcf_annual=statement_value(cash,['Free Cash Flow'])
        if fcf_annual is None:
            cash_operating=statement_value(cash,['Operating Cash Flow','Cash Flow From Continuing Operating Activities'])
            cash_capex=statement_value(cash,['Capital Expenditure'])
            if cash_operating is not None and cash_capex is not None:
                fcf_annual=cash_operating-abs(cash_capex)
        fcf_currency=report_ccy
        fcf_fx=fx_rate(fcf_currency,quote_ccy) if fcf_currency and quote_ccy else None
        valid_shares=shares_count if shares_count is not None and shares_count>0 else (market_cap/price if market_cap is not None and market_cap>0 else None)
        fcf_series=pd.Series(dtype='float64')
        if cash is not None and not cash.empty:
            if 'Free Cash Flow' in cash.index:
                fcf_series=pd.to_numeric(cash.loc['Free Cash Flow'],errors='coerce').dropna().sort_index()
            elif 'Operating Cash Flow' in cash.index and 'Capital Expenditure' in cash.index:
                ocf=pd.to_numeric(cash.loc['Operating Cash Flow'],errors='coerce')
                capex=pd.to_numeric(cash.loc['Capital Expenditure'],errors='coerce')
                fcf_series=(ocf-capex.abs()).dropna().sort_index()
        dcf_growth=None
        if len(fcf_series)>=2 and fcf_series.iloc[0]>0 and fcf_series.iloc[-1]>0:
            dcf_growth=max(-10.0,min(12.0,((float(fcf_series.iloc[-1])/float(fcf_series.iloc[0]))**(1/(len(fcf_series)-1))-1)*100))
        # CAPM för SEK: Riksbankens tioårsränta + observerad beta × antagen aktieriskpremie.
        # Använd inte ränta i fel valuta och ersätt inte saknad beta med 1,0.
        riskfree,rate_date=swedish_ten_year_yield() if quote_ccy=='SEK' else (None,None)
        beta=numeric(info.get('beta'))
        beta_source='Yahoo Finance'
        if beta is None or not 0<beta<5:
            beta=calculated_stock_beta(ticker)
            beta_source='Beräknad veckobeta mot OMX Stockholm 30 (3 år)'
        equity_risk_premium=5.5  # Öppen modellpolicy, inte observerad aktuell marknadsriskpremie.
        discount_rate=(riskfree+beta*equity_risk_premium) if riskfree is not None and beta is not None and 0<beta<5 else None
        terminal_growth=2.0
        st.markdown('**Bolagets avkastningskrav (CAPM)**')
        if discount_rate is None:
            st.caption('Riskbaserat avkastningskrav är inte tillgängligt för detta bolag. Se saknade uppgifter nedan.')
        else:
            risk_a,risk_b,risk_c=st.columns(3)
            risk_a.metric('Riskfri ränta',show(riskfree,' %',2))
            risk_b.metric('Beta – börskänslighet',show(beta,'',2))
            risk_c.metric('Avkastningskrav',show(discount_rate,' %',2))
            st.markdown(directional_indicator(beta<=1.0,'Lägre eller normal marknadskänslighet' if beta<=1.0 else 'Högre marknadskänslighet'),unsafe_allow_html=True)
            st.caption('Här visar linjen endast om beta ligger över eller under 1,0 – inte om aktien är bra eller dålig. Högre avkastningskrav är inte i sig ett dåligt resultat.')
            st.caption('Beta mäter hur känslig aktien historiskt varit för börsens rörelser. Beta 1,0 ≈ samma känslighet som marknaden; 1,5 ≈ 50 % större; 0,5 ≈ hälften. Det är ett statistiskt mått, inte en prognos.')
            with st.expander('Vad betyder beta och varför påverkar det värderingen?'):
                st.write('**Beta 1,0:** aktien har historiskt rört sig ungefär i takt med marknaden. **Beta 1,5:** aktien har i genomsnitt reagerat cirka 50 % mer på marknadsrörelser. **Beta 0,5:** aktien har reagerat cirka hälften så mycket. Sambanden gäller inte vid varje enskild börsrörelse.')
                st.write('I CAPM innebär högre beta ett högre avkastningskrav. Ett högre avkastningskrav minskar normalt det beräknade DCF-värdet, eftersom framtida kassaflöden diskonteras hårdare.')
                st.write('**Bolagets beta:** '+show(beta,'',2)+' · **Källa:** '+beta_source+'.')
                st.write('**Beräkning:** '+show(riskfree,' %',2)+' riskfri ränta + '+show(beta,'',2)+' × 5,5 % aktieriskpremie = '+show(discount_rate,' %',2)+' avkastningskrav.')
                st.caption('Beta mäter marknadskänslighet, inte alla affärs- eller finansieringsrisker. 5,5 % aktieriskpremie är ett modellantagande.')
            st.caption('Riskfri ränta: Riksbanken '+str(rate_date)+'. Beta: '+beta_source+'. CAPM-avkastningskrav: '+show(discount_rate,' %',2)+'.')
        st.caption('DCF använder historisk FCF-tillväxt från minst två jämförbara årsperioder (begränsad till −10–12 %) och terminaltillväxt 2 % (modellantagande).')
        dcf_missing=[]
        if fcf_annual is None or fcf_annual<=0: dcf_missing.append('Positivt fritt kassaflöde')
        if dcf_growth is None: dcf_missing.append('Minst två jämförbara år med positivt FCF')
        if fcf_fx is None: dcf_missing.append('Verifierad rapportvaluta och valutakurs')
        if valid_shares is None: dcf_missing.append('Antal aktier')
        if quote_ccy!='SEK': dcf_missing.append('CAPM-ränta i rätt handelsvaluta (endast SEK stöds)')
        elif riskfree is None: dcf_missing.append('Svensk tioårig statsobligationsränta')
        if beta is None or not 0<beta<5: dcf_missing.append('Aktiebeta')
        if discount_rate is not None and discount_rate<=terminal_growth: dcf_missing.append('Avkastningskrav högre än terminaltillväxt')
        if dcf_missing:
            dcf_fair=None
            st.warning('DCF kan inte beräknas ännu. Saknas: '+', '.join(dcf_missing)+'.')
        else:
            fcf_quote=fcf_annual*fcf_fx
            projected=[fcf_quote*(1+dcf_growth/100)**year for year in range(1,6)]
            r=discount_rate/100
            g=terminal_growth/100
            terminal=projected[-1]*(1+g)/(r-g)
            equity_value=sum(amount/(1+r)**year for year,amount in enumerate(projected,1))+terminal/(1+r)**5
            dcf_fair=equity_value/valid_shares
            d1,d2,d3=st.columns(3)
            d1.metric('Fritt kassaflöde, basår',show(fcf_quote/1e6,' M '+quote_ccy,1))
            d2.metric('DCF-värde per aktie',show(dcf_fair,' '+quote_ccy,2))
            d3.metric('Skillnad mot kurs',show((dcf_fair/price-1)*100,' %',1))
            st.caption('DCF utgår från senast tillgängligt Yahoo-års-FCF ('+str(fcf_currency)+'). Modellen behandlar det som kassaflöde till aktieägarna och gör därför inget extra avdrag för nettoskuld. Klassificering av räntor i kassaflödet kan påverka jämförbarheten. Terminalvärdet är särskilt känsligt för avkastningskravet.')
        st.markdown('#### 3. Framtida totalavkastning – fem år')
        st.caption('Kursutveckling från framtida EPS och P/E plus uppskattade kontantutdelningar. Årlig totalavkastning beräknas som CAGR för kurs plus utdelningar, utan återinvestering.')
        earnings_growth=None
        if len(_roic_income)>=2:
            def get_report_eps(record):
                for k in ('eps_diluted','eps_basic','eps','is_eps_diluted','is_eps_basic'):
                    value=numeric(record.get(k))
                    if value is not None: return value
                return None
            eps_new=get_report_eps(_roic_income[0])
            eps_old=get_report_eps(_roic_income[1])
            try:
                year_gap=(pd.Timestamp(_roic_income[0]['period_end_date'])-pd.Timestamp(_roic_income[1]['period_end_date'])).days/365.25
                if eps_new is not None and eps_old is not None and eps_new>0 and eps_old>0 and year_gap>=0.8 and _roic_income[0].get('currency')==_roic_income[1].get('currency'):
                    earnings_growth=max(-15.0,min(20.0,((eps_new/eps_old)**(1/year_gap)-1)*100))
            except (ValueError,TypeError,KeyError,OverflowError):
                pass
        exit_pe=pe_multiple
        dividend_growth=0.0
        st.caption('Historisk EPS-tillväxt (begränsad till −15–20 %), historisk median P/E och oförändrad utdelning används automatiskt. Dessa är modellantaganden, inte bolagets prognoser.')
        dividend_per_share=(price*dividend_yield/100) if dividend_yield is not None and dividend_yield>=0 else None
        if valuation_eps is None or dividend_per_share is None or earnings_growth is None or exit_pe is None:
            total_cagr=None
            st.info('Totalavkastning kräver verifierbar EPS och utdelningsdata. Om någon del saknas visas ingen fullständig totalavkastning.')
        else:
            future_eps=valuation_eps*(1+earnings_growth/100)**5
            exit_price=future_eps*exit_pe
            total_dividends=sum(dividend_per_share*(1+dividend_growth/100)**year for year in range(1,6))
            total_end_value=exit_price+total_dividends
            total_cagr=((total_end_value/price)**(1/5)-1)*100 if total_end_value>0 else -100.0
            t1,t2,t3=st.columns(3)
            t1.metric('Beräknad kurs år fem',show(exit_price,' '+quote_ccy,2))
            t2.metric('Utdelningar under fem år',show(total_dividends,' '+quote_ccy,2))
            t3.metric('Årlig totalavkastning',show(total_cagr,' %',1))
            st.caption('Utdelningen utgår från nuvarande/historisk direktavkastning och växer enligt valt antagande. Inga utdelningar återinvesteras och ingen skatt ingår.')
        st.markdown('#### 4. Säkerhetsmarginal')
        st.caption('Visar ett möjligt inköpspris med rabatt mot respektive metods motiverade värde.')
        discount=20
        st.caption('Fast säkerhetsmarginal: 20 %.')
        rows=[]
        for method,fair in [('P/E',pe_fair),('DCF',dcf_fair)]:
            if fair is not None and fair>0:
                rows.append({'Metod':method,'Beräknat värde':show(fair,' '+quote_ccy,2),'Köppris med rabatt':show(fair*(1-discount/100),' '+quote_ccy,2),'Kurs jämfört med värde':show((price/fair-1)*100,' %',1)})
        if rows:
            st.dataframe(pd.DataFrame(rows),hide_index=True,use_container_width=True)
        else:
            st.info('Ingen verifierbar värdering finns för att beräkna säkerhetsmarginal.')
        st.caption('Metoderna kan ge olika resultat. Ingen sammanvägd riktkurs skapas automatiskt eftersom kassaflödes- och vinstvärdering bygger på olika antaganden.')

st.divider()
st.caption('Källor: Yahoo Finance via yfinance för nyckeltal och historik; Google News RSS för nyhetssökning; användarens egna uppladdade dokument. Inga automatiska AI-slutsatser påstås ha gjorts. Appen körs på egen Streamlit-server och filerna behandlas i den sessionen.')
