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
'Övriga bolag': {'Volati B':'VOLAT-B.ST','Salix Group':'SALIX.ST','Röko':'ROKO.ST','Asker Healthcare':'ASKER.ST','Verisure':'VSURE.ST','Paradox Interactive':'PDX.ST'}
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
    obj=yf.Ticker(ticker)
    info=obj.info or {}
    history=obj.history(period='5y',auto_adjust=False)
    fin=obj.income_stmt
    cash=obj.cashflow
    balance=obj.balance_sheet
    return info,history,fin,cash,balance

@st.cache_data(ttl=3600,show_spinner=False)
def news(company):
    since=dt.datetime.now(dt.timezone.utc)-dt.timedelta(days=183)
    q=urllib.parse.quote(f'"{company}" (kvartalsrapport OR förvärv OR pressmeddelande OR resultat) when:6m')
    url=f'https://news.google.com/rss/search?q={q}&hl=sv&gl=SE&ceid=SE:sv'
    feed=feedparser.parse(requests.get(url,timeout=12).content)
    out=[]
    for entry in feed.entries:
        published=entry.get('published_parsed')
        if not published:continue
        date=dt.datetime(*published[:6],tzinfo=dt.timezone.utc)
        if date<since:continue
        out.append({'Datum':date.date().isoformat(),'Rubrik':entry.get('title',''),'Länk':entry.get('link','')})
    return out[:35]

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
st.caption('Automatisk bolagsöversikt · Svenska storbolag · Egna rapporter · Scenariovärdering')
st.info('Data hämtas vid val av bolag via Yahoo Finance (inofficiell källa, kan vara fördröjd eller ofullständig). Saknade fält visas som ”Saknas” – inga värden hittas på. Källorna bör kontrolleras mot bolagets rapporter.')
with st.sidebar:
    st.header('Välj bolag')
    q=st.text_input('Sök bolagsnamn',placeholder='T.ex. Investor, Volati, Atlas Copco')
    sector=st.selectbox('Bransch', ['Alla']+list(GROUPS))
    options=[n for n,(t,g) in CAT.items() if (sector=='Alla' or g==sector) and q.casefold() in n.casefold()]
    if not options:
        st.warning('Inga träffar. Du kan ange en egen Yahoo Finance-ticker nedan.')
    name=st.selectbox('Bolag',options) if options else None
    custom=st.text_input('Egen ticker (valfritt)',placeholder='Exempel: VOLAT-B.ST')
    ticker=custom.strip().upper() or (CAT[name][0] if name else '')
    st.caption(f'Ticker: {ticker or "–"}')
    if st.button('Uppdatera data',use_container_width=True): st.cache_data.clear();st.rerun()
if not ticker:st.stop()
with st.spinner(f'Hämtar data för {name or ticker}…'):
    try: info,hist,fin,cash,balance=fetch(ticker)
    except Exception as ex:
        st.error(f'Datakällan kunde inte hämtas: {ex}. Försök igen senare eller välj annan ticker.');st.stop()
if not info and hist.empty:
    st.error('Ingen data hittades för den tickern. Kontrollera bolaget eller prova igen.');st.stop()
company=info.get('longName') or name or ticker
st.subheader(company)
st.caption(f'{ticker} · Valuta: {info.get("currency","okänd")} · Börsdata hämtad {dt.datetime.now().strftime("%Y-%m-%d %H:%M")} · Yahoo Finance')

price=val(info,'currentPrice','regularMarketPrice','previousClose')
if price is None and not hist.empty:price=numeric(hist['Close'].iloc[-1])
revenue=val(info,'totalRevenue');ebitda=val(info,'ebitda');margin=val(info,'operatingMargins');debt=val(info,'totalDebt');cash_val=val(info,'totalCash');opcf=val(info,'operatingCashflow');eps=val(info,'trailingEps');pe=val(info,'trailingPE')
netdebt=debt-cash_val if debt is not None and cash_val is not None else None

st.markdown('### Finansiell översikt')
cols=st.columns(4)
for col,(label,value) in zip(cols,[('Aktiekurs',show(price,' '+info.get('currency',''),2)),('Omsättning LTM',show(revenue/1e6,' M',0) if revenue is not None else 'Saknas'),('Rörelsemarginal',show(margin*100,' %') if margin is not None else 'Saknas'),('P/E',show(pe,'×') if pe is not None else 'Saknas')]):col.metric(label,value)
cols=st.columns(4)
for col,(label,value) in zip(cols,[('EBITDA LTM',show(ebitda/1e6,' M',0) if ebitda is not None else 'Saknas'),('Nettoskuld',show(netdebt/1e6,' M',0) if netdebt is not None else 'Saknas'),('Operativt kassaflöde',show(opcf/1e6,' M',0) if opcf is not None else 'Saknas'),('EPS senaste 12 mån',show(eps,'',2))]):col.metric(label,value)
st.caption('M = miljoner i angiven rapporteringsvaluta. Rörelsemarginal är inte samma sak som EBITA-marginal. Nettoskuld = rapporterade skulder minus likvida medel (förenklad beräkning). LTM-perioder kan skilja sig mellan fält.')

one,two,three,four=st.tabs(['Kurs och historik','Rapporterade data','Nyheter – 6 månader','Ladda upp dokument och värdera'])
with one:
    if not hist.empty:
        h=hist.reset_index();h['Datum']=pd.to_datetime(h['Date']).dt.tz_localize(None)
        st.plotly_chart(px.line(h,x='Datum',y='Close',title='Aktiekurs – senaste fem åren'),use_container_width=True)
    else:st.warning('Kursserie saknas.')
    st.markdown('**Bolagsprofil**');st.write(info.get('longBusinessSummary','Beskrivning saknas.'))
with two:
    st.caption('Rapporttabeller från dataleverantören. Kontrollera redovisningsvaluta och rapportdatum.')
    for label,frame in [('Resultaträkning',fin),('Kassaflöde',cash),('Balansräkning',balance)]:
        with st.expander(label,expanded=label=='Resultaträkning'):
            if frame is None or frame.empty:st.warning('Data saknas.')
            else:
                display=frame.copy();display.columns=[str(c.date()) if hasattr(c,'date') else str(c) for c in display.columns]
                st.dataframe(display,use_container_width=True)
with three:
    st.caption('Google News RSS, filtrerat till publiceringsdatum inom sex månader. Rubriker är inte verifierade bolagsmeddelanden och kan innehålla andra bolag.')
    if st.button('Hämta senaste nyheter'):
        try:
            items=news(company)
            if not items:st.warning('Inga träffar kunde hämtas.')
            for item in items:st.markdown(f'**{item["Datum"]}** · [{item["Rubrik"]}]({item["Länk"]})')
        except Exception as ex:st.error(f'Nyheter kunde inte hämtas: {ex}')
    st.markdown(f'[Sök bolagets pressmeddelanden](https://www.google.com/search?q={urllib.parse.quote(company+" investor relations press releases")})')
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
st.divider()
st.caption('Källor: Yahoo Finance via yfinance för nyckeltal och historik; Google News RSS för nyhetssökning; användarens egna uppladdade dokument. Inga automatiska AI-slutsatser påstås ha gjorts. Appen körs på egen Streamlit-server och filerna behandlas i den sessionen.')
