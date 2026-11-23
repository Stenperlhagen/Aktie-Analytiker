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
    st.caption('Sök svenska börsbolag i samma lista, oavsett börssegment.')
    q=st.text_input('Sök bolagsnamn eller ticker',placeholder='T.ex. Volati, Investor, VBG')
    options=sorted(
        [n for n,t in ALL_COMPANIES.items() if q.casefold() in n.casefold() or q.casefold() in t.casefold()],
        key=str.casefold
    )
    if not options:
        st.warning('Inga träffar i listan. Ange en egen Yahoo Finance-ticker nedan.')
    name=st.selectbox('Bolag',options) if options else None
    custom=st.text_input('Egen ticker (om bolaget saknas)',placeholder='Exempel: VOLO.ST')
    ticker=custom.strip().upper() or (ALL_COMPANIES[name] if name else '')
    st.caption(f'Ticker: {ticker or "–"}')
    if st.button('Uppdatera data',use_container_width=True):
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
if pe is None and eps is not None and eps>0 and price is not None: pe=price/eps


st.markdown('### Finansiell översikt')
cols=st.columns(4)
for col,(label,value) in zip(cols,[('Aktiekurs',show(price, ' ' + (info.get('currency') or quick.get('currency') or ''), 2)),('Omsättning (LTM/helår)',show(revenue/1e6,' M',0) if revenue is not None else 'Saknas'),('Rörelsemarginal',show(margin*100,' %') if margin is not None else 'Saknas'),('P/E',show(pe,'×') if pe is not None else 'Saknas')]):col.metric(label,value)
cols=st.columns(4)
for col,(label,value) in zip(cols,[('EBITDA (LTM/helår)',show(ebitda/1e6,' M',0) if ebitda is not None else 'Saknas'),('Nettoskuld',show(netdebt/1e6,' M',0) if netdebt is not None else 'Saknas'),('Operativt kassaflöde',show(opcf/1e6,' M',0) if opcf is not None else 'Saknas'),('EPS senaste 12 mån',show(eps,'',2))]):col.metric(label,value)
st.caption('Omsättning kan avse senaste rapporterade helår när LTM saknas; kontrollera perioderna i rapporttabellerna. M = miljoner i angiven rapporteringsvaluta. Rörelsemarginal är inte samma sak som EBITA-marginal. Nettoskuld = rapporterade skulder minus likvida medel (förenklad beräkning). LTM-perioder kan skilja sig mellan fält.')

one,two,three,four,five=st.tabs(['Kurs och historik','Rapporterade data','Nyheter – 6 månader','Dokument och scenarier','Analys om bolaget'])
with one:
    if not hist.empty:
        h=hist.reset_index();h['Datum']=pd.to_datetime(h['Date']).dt.tz_localize(None)
        st.plotly_chart(px.line(h,x='Datum',y='Close',title='Aktiekurs – senaste fem åren'),use_container_width=True)
    else:st.warning('Kursserie saknas.')
    st.markdown('**Bolagsprofil**');st.write(info.get('longBusinessSummary','Beskrivning saknas.'))
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
    st.caption('Publicerade artiklar de senaste sex månaderna via Google News och Bing News RSS. Artiklar kan handla om andra bolag med liknande namn.')
    if st.button('Uppdatera nyheter', key='refresh_news'):
        news.clear()
        st.rerun()
    with st.spinner('Söker nyheter…'):
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
        if not news_diagnostics:
            st.text('Ingen diagnostik tillgänglig.')
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
    st.caption('DCF, EV/EBITA och historisk multipel enligt din tidigare bolagsvärderingsmodell. Belopp i MSEK och aktiekurs i SEK. Saknade underlag måste anges manuellt – inga påhittade värden.')
    shares_raw = val(info, 'sharesOutstanding')
    if shares_raw is None:
        shares_raw = val(quick, 'shares')
    # FCFF är inte samma sak som operativt kassaflöde eller fritt kassaflöde efter ränta.
    # EBITA kan inte generellt härledas från EBITDA utan tillförlitliga avskrivningsuppgifter.
    st.markdown('#### 1. Bolagsuppgifter')
    x1,x2,x3 = st.columns(3)
    with x1:
        p = st.number_input('Aktiekurs (kr)', min_value=0.0, value=float(price or 0), step=1.0, key='bv_price')
        shares = st.number_input('Antal aktier (miljoner)', min_value=0.0, value=float(shares_raw/1e6) if shares_raw else 0.0, step=1.0, key='bv_shares')
    with x2:
        nd = st.number_input('Nettoskuld (MSEK)', value=float(netdebt/1e6) if netdebt is not None else 0.0, step=10.0, key='bv_debt')
        ebita = st.number_input('Senaste EBITA (MSEK)', value=0.0, step=10.0, key='bv_ebita')
    with x3:
        fcff = st.number_input('FCFF – kassaflöde till företaget (MSEK)', value=0.0, step=10.0, key='bv_fcff')
        st.caption('EBITA och FCFF fylls inte automatiskt med EBITDA eller operativt kassaflöde. Ange verifierade belopp från bolagets rapport.')
    st.markdown('#### 2. Prognos och värderingsantaganden')
    c1,c2,c3 = st.columns(3)
    with c1:
        growth = st.number_input('Årlig FCFF-tillväxt år 1–5 (%)', value=6.0, key='bv_growth')
        wacc = st.number_input('WACC / diskonteringsränta (%)', value=10.0, min_value=0.01, key='bv_wacc')
    with c2:
        terminal = st.number_input('Evig tillväxt efter år 5 (%)', value=2.0, key='bv_terminal')
        multiple = st.number_input('Motiverad EV/EBITA-multipel', value=12.0, min_value=0.0, key='bv_multiple')
    with c3:
        historical = st.number_input('Historisk genomsnittlig EV/EBITA', value=11.0, min_value=0.0, key='bv_historical')
        margin_safe = st.number_input('Säkerhetsmarginal (%)', value=20.0, min_value=0.0, max_value=100.0, key='bv_margin')
    st.markdown('#### 3. Modellvikter')
    v1,v2,v3 = st.columns(3)
    wd = v1.number_input('DCF-vikt', value=50.0, min_value=0.0, key='bv_wd')
    wm = v2.number_input('EV/EBITA-vikt', value=30.0, min_value=0.0, key='bv_wm')
    wh = v3.number_input('Historisk multipel-vikt', value=20.0, min_value=0.0, key='bv_wh')
    if p <= 0 or shares <= 0:
        st.warning('Aktiekurs och antal aktier måste vara större än noll för att beräkna värde per aktie.')
    elif wacc <= terminal or growth <= -100 or terminal <= -100 or wd+wm+wh <= 0:
        st.error('Kontrollera antagandena: WACC måste överstiga evig tillväxt, tillväxten vara över −100 % och minst en vikt positiv.')
    elif (wd > 0 and fcff == 0) or ((wm > 0 or wh > 0) and ebita == 0):
        st.info('Fyll i EBITA och FCFF från bolagets rapport för att få en meningsfull värdering. Modellen använder inte exempelbelopp som verklig data.')
    else:
        g,w,t = growth/100,wacc/100,terminal/100
        forecast = [{'År':i,'FCFF (MSEK)':fcff*(1+g)**i,'Nuvärde (MSEK)':fcff*(1+g)**i/(1+w)**i} for i in range(1,6)]
        terminal_pv = (fcff*(1+g)**5*(1+t)/(w-t))/(1+w)**5
        dcf_ev = sum(row['Nuvärde (MSEK)'] for row in forecast)+terminal_pv
        methods = [('DCF',(dcf_ev-nd)/shares,wd),('EV/EBITA',(ebita*multiple-nd)/shares,wm),('Historisk EV/EBITA',(ebita*historical-nd)/shares,wh)]
        total_weight = wd+wm+wh
        fair = sum(value*weight/total_weight for _,value,weight in methods)
        upside = (fair/p-1)*100
        buy = fair*(1-margin_safe/100)
        status = 'Potentiellt undervärderad' if upside > 20 else ('Potentiellt övervärderad' if upside < -20 else 'Nära motiverat värde')
        k1,k2,k3,k4 = st.columns(4)
        k1.metric('Aktiekurs',show(p,' kr',2))
        k2.metric('Motiverat aktievärde',show(fair,' kr',2))
        k3.metric('Potential',show(upside,' %',1))
        k4.metric('Bedömning',status)
        left,right = st.columns(2)
        with left:
            st.markdown('#### Värdering per metod')
            st.dataframe(pd.DataFrame([{'Metod':n,'Värde/aktie (kr)':round(v,2),'Vikt (%)':round(weight/total_weight*100,1)} for n,v,weight in methods]),hide_index=True,use_container_width=True)
        with right:
            st.markdown('#### Säkerhetsmarginal')
            st.metric('Högsta köpkurs med säkerhetsmarginal',show(buy,' kr',2))
            st.caption('Aktiekursen ligger under köpkursen.' if p <= buy else 'Aktiekursen ligger över köpkursen.')
        st.markdown('#### DCF-prognos – fem år')
        st.dataframe(pd.DataFrame(forecast).round(2),hide_index=True,use_container_width=True)
        st.caption(f'Nuvärde av terminalvärdet: {show(terminal_pv/1, " MSEK", 1)}. DCF-företagsvärde: {show(dcf_ev, " MSEK", 1)}.')
        export = pd.DataFrame([{'Bolag':company,'Aktiekurs':p,'Motiverat värde':fair,'Potential (%)':upside,'Köpkurs med säkerhetsmarginal':buy}])
        st.download_button('Exportera analys (CSV)',export.to_csv(index=False).encode('utf-8-sig'),file_name='bolagsvardering.csv',mime='text/csv')
    st.caption('Modellen är en indikativ värdering och ingen köp- eller säljrekommendation. Kontrollera normaliserat kassaflöde, EBITA, nettoskuld, aktieantal, utspädning och multiplar mot rapporterna. Historisk EV/EBITA fylls manuellt eftersom den inte kan fastställas säkert från nuvarande datakälla.')

st.divider()
st.caption('Källor: Yahoo Finance via yfinance för nyckeltal och historik; Google News RSS för nyhetssökning; användarens egna uppladdade dokument. Inga automatiska AI-slutsatser påstås ha gjorts. Appen körs på egen Streamlit-server och filerna behandlas i den sessionen.')
