import ast
import io
import pathlib
import csv
import math
import os
import requests
import urllib.parse
import datetime as dt
import pandas as pd
from pypdf import PdfReader
import re
import yfinance as yf
import sys
from functools import lru_cache

# Load exact production definitions without executing Streamlit UI.
tree=ast.parse(pathlib.Path('app.py').read_text(encoding='utf-8'))
names={'emission_adjustment','investment_score','show','roic_company_data','GROUPS','CAT','ADDITIONAL_COMPANIES','ALL_COMPANIES','INVESTMENT_TICKERS','BANK_TICKERS','INSURANCE_TICKERS','PROPERTY_TICKERS','VOLATI_Q3_2026','REPORTED_NAV','VALUATION_PROFILES','PROFILE_TICKERS','PROFILE_BY_SYMBOL','RESEARCH_TICKERS','HYBRID_PHARMA_TICKERS','ESTABLISHED_PHARMA_TICKERS','KNOWN_GROUP_BY_TICKER','ADDITIONAL_GROUP_FALLBACK','numeric','val','fx_rate','ecb_exchange_rates','borsapi_instruments','borsapi_instrument','borsapi_instrument_events','borsapi_reports','borsapi_history','borsapi_annual_frames','borsapi_r12_records','borsapi_r12_response','request_borsdata','borsapi_market_cap_map','bd_field','yahoo_annual_statements','yahoo_r12_bridge','matched_yahoo_r12','comparable_yahoo_periods','fill_matching_statement_periods','abb_r12_ebita','reported_operating_metrics','trailing_dividend_yield','report_ebita_extract','avanza_market_snapshot','verified_capitalization','fetch','operating_valuation_profile','comparable_ev_ebita','valuation_category','valuation_engine'}
nodes=[]
for node in tree.body:
    if isinstance(node,ast.FunctionDef) and node.name in names:
        node.decorator_list=[]
        nodes.append(node)
    elif isinstance(node,ast.Assign) and any(isinstance(t,ast.Name) and t.id in names for t in node.targets):
        nodes.append(node)
    elif isinstance(node,ast.Expr) and ast.unparse(node).startswith(('KNOWN_GROUP_BY_TICKER.update(','VALUATION_PROFILES.update(','PROFILE_BY_SYMBOL.update(')):
        nodes.append(node)
exec(compile(ast.Module(body=nodes,type_ignores=[]),'app.py','exec'))
fx_rate=lru_cache(maxsize=128)(fx_rate)
roic_company_data=lru_cache(maxsize=256)(roic_company_data)
def roic_key():
    return os.environ.get('ROIC_API_KEY','').strip()

def borsapi_key():
    # CI does not have Streamlit Secrets. Do not require credentials for mathematical tests.
    return os.environ.get('BORSAPI_API_KEY','').strip()

production_adapters={name:globals()[name] for name in ('borsapi_instruments','borsapi_instrument','borsapi_reports','borsapi_history','borsapi_annual_frames','borsapi_r12_records','borsapi_r12_response','borsapi_market_cap_map','avanza_market_snapshot','verified_capitalization')}

def borsapi_instrument(symbol, api_key):
    return {},'Offline audit'

def borsapi_reports(symbol, api_key):
    # Deterministic offline fallback; integration is tested separately.
    return {}, 'Börsdata API unavailable in offline audit'

def borsapi_history(symbol, api_key):
    return pd.DataFrame(), 'Börsdata API unavailable in offline audit'

def borsapi_annual_frames(symbol, api_key):
    return None, 'Börsdata API unavailable in offline audit'

def borsapi_r12_response(symbol, api_key):
    return [],'Offline audit'

def borsapi_r12_records(symbol, api_key):
    # Offline audit: R12 endpoint is unavailable without credentials.
    return []

def borsapi_instruments(api_key):
    return [],'Offline audit'
def borsapi_market_cap_map(api_key):
    return {},'Offline audit'
def avanza_market_snapshot(symbol):
    return {},'Offline audit'
def verified_capitalization(symbol,info,quick,fin,history):
    return {'market_cap':None,'shares':None,'source':'Offline audit','warnings':[],'verified':False}
def tests():
    # Registry is a checked-in set of independently observed Avanza listing identities,
    # never a substitute for fresh detail validation in production.
    registry_path=pathlib.Path(__file__).with_name('avanza-registry.json')
    if registry_path.exists():
        import json
        registry=json.loads(registry_path.read_text(encoding='utf-8'))['listings']
        assert set(registry)==set(ALL_COMPANIES.values()), 'Avanza registry does not cover active universe'
        assert len({r['orderbook_id'] for r in registry.values()})==len(registry)
        for symbol,r in registry.items():
            assert str(r['orderbook_id']).isdigit()
            assert re.fullmatch(r'[A-Z]{2}[A-Z0-9]{9}[0-9]',r['isin'])
            assert re.sub(r'[^A-Z0-9]','',r['ticker'].upper())==re.sub(r'[^A-Z0-9]','',symbol.rsplit('.',1)[0].upper())
            assert r['currency']==('EUR' if symbol=='VSURE.ST' else 'DKK' if symbol.endswith('.CO') else 'SEK')
            assert r['market_code']==('FNSE' if symbol in ('BOKUS.ST','SUS.ST','SES.ST') else 'XCSE' if symbol.endswith('.CO') else 'XSTO')
        print('Avanza registry coverage:',len(registry),'of',len(set(ALL_COMPANIES.values())))
    assert abs(comparable_ev_ebita(245e6,1331e6,79406571,12,1)-20.263)<0.01
    assert comparable_ev_ebita(100,None,10,12,1) is None
    assert comparable_ev_ebita(100,1300,10,12,1) is None
    p=pd.Timestamp('2025-12-31')
    fin=pd.DataFrame({p:[1000e6,100e6,90e6]},index=['Total Revenue','EBITA','Operating Income'])
    balance=pd.DataFrame({p:[200e6]},index=['Net Debt'])
    cash=pd.DataFrame({p:[20e6,-25e6,-5e6]},index=['Depreciation And Amortization','Capital Expenditure','Change In Working Capital'])
    info={'currency':'SEK','financialCurrency':'SEK','currentPrice':100,'sharesOutstanding':10e6,'marketCap':1e9,'_cap_verified':True,'_preferred_cap':0.0,'_common_cap':1e9}
    v=valuation_engine('VOLO-TEST.ST',{**info,'sector':'Industrials'},fin,cash,balance,{},pd.DataFrame())
    assert abs(v['models']['EV/EBITA']-90)<1e-8
    # Marknadsdata utan verifierad kapitalstruktur får aldrig ge automatisk riktkurs.
    unverified=valuation_engine('VOLO-TEST.ST',{**info,'_cap_verified':False},fin,cash,balance,{},pd.DataFrame(),use_roic=False)
    assert unverified.get('target') is None
    assert not unverified.get('models')
    assert any('Börsvärde ej verifierat' in msg for msg in unverified.get('audit',[]))
    assert abs(v['inputs']['FCFF']-61.46e6)<1
    assert v['operating_inputs']==v['inputs']
    assert v['operating_period']==str(p)[:10]
    assert abs(v['target']-sum(v['models'].values())/len(v['models']))<1e-8
    assert valuation_category('AAK.ST')=='Rörelsedrivande bolag'
    assert operating_valuation_profile('AAK.ST')=='Livsmedelsproduktion'
    assert comparable_ev_ebita(4777e6,3134e6,259708000,11,1)>0
    # A matched partial R12 is not replaced by the old AAK snapshot.
    # In particular missing EBITA must not erase FCFF or override share count.
    r12={'currency':'SEK','report_End_Date':'2026-09-30','revenues':1000,
         'operating_Income':90,'net_Debt':200,'depreciation_And_Amortization':20,
         'capital_Expenditure':25,'change_In_Working_Capital':-5}
    original_r12=borsapi_r12_records
    original_bridge=matched_yahoo_r12
    try:
        globals()['borsapi_r12_records']=lambda *args:[r12]
        globals()['matched_yahoo_r12']=lambda *args:{}
        partial_fin=fin.drop(index='EBITA')
        partial=valuation_engine('AAK.ST',info,partial_fin,cash,balance,{},pd.DataFrame())
        assert partial['operating_period']=='2026-09-30'
        assert partial['operating_inputs']['EBIT']==90e6
        assert partial['operating_inputs']['Antal aktier']==10e6
        assert abs(partial['operating_inputs']['FCFF']-61.46e6)<1
        assert partial['operating_inputs']['Nettoskuld']==200e6
    finally:
        globals()['borsapi_r12_records']=original_r12
        globals()['matched_yahoo_r12']=original_bridge
    assert abs(v['buy_limit']-v['target']*.8)<1e-8
    assert abs(v['upside']-(v['target']/v['price']-1)*100)<1e-8
    score,items,coverage=investment_score(v,info)
    assert score is None or 0<=score<=100
    assert 0<=coverage<=100
    for eps,expected in [(5,20),(-5,-20),(0,None)]:
        result=valuation_engine('VOLO-TEST.ST',{**info,'sector':'Industrials','trailingEps':eps},fin,cash,balance,{},pd.DataFrame())
        assert result.get('pe_ttm')==expected, (eps,result.get('pe_ttm'))
        assert result.get('pe_zero')==(eps==0)
    q=pd.Timestamp('2024-12-31')
    bank_fin=pd.DataFrame({p:[150e6],q:[120e6]},index=['Net Income Common Stockholders'])
    bank_balance=pd.DataFrame({p:[1100e6],q:[900e6]},index=['Common Stock Equity'])
    bank=valuation_engine('SEB-A.ST',info,bank_fin,pd.DataFrame(),bank_balance,{},pd.DataFrame())
    assert abs(bank['roe']-.15)<1e-9
    assert abs(bank['inputs']['Motiverat P/B']-1.625)<1e-9
    # Missing profile metadata must not be mistaken for a successful valuation.
    no_currency=valuation_engine('VOLO-TEST.ST',{**info,'financialCurrency':None,'sector':'Industrials'},fin,cash,balance,{},pd.DataFrame(),use_roic=False)
    assert no_currency.get('target') is None
    recovered=valuation_engine('VOLO-TEST.ST',{**info,'sector':'Industrials'},fin,cash,balance,{},pd.DataFrame(),use_roic=False)
    assert recovered.get('target') is not None
    assert 'VOLV-A.ST' not in ALL_COMPANIES.values() and 'VOLV-B.ST' in ALL_COMPANIES.values()
    reported={'currency':'SEK','operating_Income':100,'net_Debt':50,'free_Cash_Flow':20}
    assert reported_operating_metrics(reported,'SEK','SEK',950e6)=={
        'ebit':100e6,'net_debt':50e6,'free_cash_flow':20e6,'ev_ebit':10.0,
    }
    assert reported_operating_metrics(reported,'USD','SEK',950e6)=={}
    assert reported_operating_metrics({**reported,'operating_Income':-100},'SEK','SEK',950e6)['ev_ebit'] is None
    abb=abb_r12_ebita('ABB.ST',{'currency':'USD','operating_Income':6716},'2026-09-30')
    assert abb['value']==6911e6
    assert abs(abb['fcff']-(6716*(1-1601/6702)+880-1077+208)*1e6)<1
    assert abb_r12_ebita('ABB.ST',{'currency':'USD','operating_Income':6000},'2026-09-30') is None
    bokus={'currency':'SEK','revenues':2210.3,'operating_Income':106.7}
    official=abb_r12_ebita('BOKUS.ST',bokus,'2026-09-30')
    assert official['value']==152.6e6 and 'fcff' not in official
    bokus_full={**bokus,'cash_Flow_From_Operating_Activities':268.6,'net_Debt':453.0}
    bokus_bridge=abb_r12_ebita('BOKUS.ST',bokus_full,'2026-09-30')
    assert abs(bokus_bridge['fcff']-(106.7*(1-18.5/105.3)+201.9-26.1-13.2-4.7)*1e6)<1
    assert abs(bokus_bridge['fcff_tax_rate']-18.5/105.3)<1e-10
    assert 'fcff' not in abb_r12_ebita('BOKUS.ST',{**bokus_full,'net_Debt':177.3},'2026-09-30')
    assert 'fcff' not in abb_r12_ebita('BOKUS.ST',{**bokus_full,'cash_Flow_From_Operating_Activities':194.0},'2026-09-30')
    assert abb_r12_ebita('BOKUS.ST',{**bokus,'currency':'EUR'},'2026-09-30') is None
    assert abb_r12_ebita('BOKUS.ST',{**bokus,'operating_Income':100},'2026-09-30') is None
    assert abb_r12_ebita('BOKUS.ST',{**bokus,'revenues':2000},'2026-09-30') is None
    assert abb_r12_ebita('BOKUS.ST',bokus,'2026-06-30') is None
    assert abb_r12_ebita('BOKUS-OTHER.ST',bokus,'2026-09-30') is None
    addtech={'currency':'SEK','revenues':23493,'operating_Income':3288,'cash_Flow_From_Operating_Activities':2982}
    add=abb_r12_ebita('ADDT-B.ST',addtech,'2026-09-30')
    assert add['value']==3888e6
    assert abs(add['fcff']-(3288*(2357/3064)+1086-192-310)*1e6)<1
    assert abs(add['fcff_tax_rate']-707/3064)<1e-10
    assert 'fcff' not in abb_r12_ebita('ADDT-B.ST',{**addtech,'cash_Flow_From_Operating_Activities':3500},'2026-09-30')
    assert 'fcff' not in abb_r12_ebita('ADDT-B.ST',{k:v for k,v in addtech.items() if k!='cash_Flow_From_Operating_Activities'},'2026-09-30')
    assert abb_r12_ebita('ADDT-B.ST',addtech,'2026-06-30') is None
    brav={'currency':'SEK','revenues':29614,'operating_Income':1996,
          'cash_Flow_From_Operating_Activities':2411,'net_Debt':2564}
    brav_bridge=abb_r12_ebita('BRAV.ST',brav,'2026-09-30')
    assert brav_bridge['value']==1996e6
    assert abs(brav_bridge['fcff']-(2724-(1996-158)*0.206)*1e6)<1
    assert brav_bridge['fcff_tax_rate']==0.206
    assert 'fcff' not in abb_r12_ebita('BRAV.ST',{**brav,'net_Debt':1115},'2026-09-30')
    assert 'fcff' not in abb_r12_ebita('BRAV.ST',{**brav,'cash_Flow_From_Operating_Activities':1453},'2026-09-30')
    assert abb_r12_ebita('BRAV.ST',{**brav,'currency':'EUR'},'2026-09-30') is None
    assert abb_r12_ebita('BRAV.ST',brav,'2026-06-30') is None
    instal={'currency':'SEK','revenues':14705,'net_Sales':14604,'operating_Income':858,
            'cash_Flow_From_Operating_Activities':723,'net_Debt':4724}
    instal_bridge=abb_r12_ebita('INSTAL.ST',instal,'2026-09-30')
    assert instal_bridge['value']==1001e6
    assert abs(instal_bridge['fcff']-(859*0.794+572-151-226)*1e6)<1
    assert abb_r12_ebita('INSTAL.ST',{**instal,'revenues':14604},'2026-09-30')['value']==1001e6
    assert abb_r12_ebita('INSTAL.ST',{**instal,'revenues':14900},'2026-09-30') is None
    assert abb_r12_ebita('INSTAL.ST',{**instal,'net_Sales':14000},'2026-09-30') is None
    assert abb_r12_ebita('INSTAL.ST',{**instal,'operating_Income':850},'2026-09-30') is None
    assert abb_r12_ebita('INSTAL.ST',instal,'2026-06-30') is None
    assert abb_r12_ebita('INSTAL.ST',{**instal,'currency':'EUR'},'2026-09-30') is None
    assert 'fcff' not in abb_r12_ebita('INSTAL.ST',{**instal,'net_Debt':3846},'2026-09-30')
    assert 'fcff' not in abb_r12_ebita('INSTAL.ST',{**instal,'cash_Flow_From_Operating_Activities':800},'2026-09-30')
    atlas={'currency':'SEK','revenues':176216,'operating_Income':35930,
           'cash_Flow_From_Operating_Activities':32147,'net_Debt':14244}
    atlas_bridge=abb_r12_ebita('ATCO-B.ST',atlas,'2026-09-30')
    assert atlas_bridge['value']==38603e6
    assert abs(atlas_bridge['fcff_tax_rate']-7700/35385)<1e-10
    assert abs(atlas_bridge['fcff']-(35930*(1-7700/35385)+10289-7407-2335)*1e6)<1
    assert abb_r12_ebita('ATCO-A.ST',atlas,'2026-09-30')==atlas_bridge
    assert abb_r12_ebita('ATCO-B.ST',{**atlas,'revenues':170000},'2026-09-30') is None
    assert abb_r12_ebita('ATCO-B.ST',{**atlas,'operating_Income':35000},'2026-09-30') is None
    assert abb_r12_ebita('ATCO-B.ST',{**atlas,'currency':'USD'},'2026-09-30') is None
    assert abb_r12_ebita('ATCO-B.ST',atlas,'2026-06-30') is None
    assert 'fcff' not in abb_r12_ebita('ATCO-B.ST',{**atlas,'net_Debt':13000},'2026-09-30')
    assert 'fcff' not in abb_r12_ebita('ATCO-B.ST',{**atlas,'cash_Flow_From_Operating_Activities':30000},'2026-09-30')
    addlife={'currency':'SEK','revenues':10698,'net_Sales':10698,'operating_Income':1079,
             'cash_Flow_From_Operating_Activities':1393,'net_Debt':4746}
    alif=abb_r12_ebita('ALIF-B.ST',addlife,'2026-09-30')
    assert alif['value']==1501e6
    assert abs(alif['fcff_tax_rate']-248/894)<1e-10
    assert abs(alif['fcff']-(1079*(1-248/894)+740-9-279)*1e6)<1
    assert abb_r12_ebita('ALIF-B.ST',{**addlife,'net_Sales':10000},'2026-09-30') is None
    assert abb_r12_ebita('ALIF-B.ST',{**addlife,'operating_Income':1000},'2026-09-30') is None
    assert abb_r12_ebita('ALIF-B.ST',{**addlife,'currency':'EUR'},'2026-09-30') is None
    assert abb_r12_ebita('ALIF-B.ST',addlife,'2026-06-30') is None
    assert abb_r12_ebita('ALIF-A.ST',addlife,'2026-09-30') is None
    assert 'fcff' not in abb_r12_ebita('ALIF-B.ST',{**addlife,'net_Debt':3866},'2026-09-30')
    assert 'fcff' not in abb_r12_ebita('ALIF-B.ST',{**addlife,'cash_Flow_From_Operating_Activities':1000},'2026-09-30')
    assa={'currency':'SEK','net_Sales':154088,'revenues':154088,'operating_Income':25702,
          'cash_Flow_From_Operating_Activities':25041,'net_Debt':64212}
    matched=abb_r12_ebita('ASSA-B.ST',assa,'2026-09-30')
    assert matched['value']==27442e6
    assert abs(matched['fcff']-(25702*(1-5747/22717)+6287-3129+814)*1e6)<1
    assert abb_r12_ebita('ASSA-B.ST',{**assa,'currency':'USD'},'2026-09-30') is None
    assert abb_r12_ebita('ASSA-B.ST',assa,'2026-06-30') is None
    assert abb_r12_ebita('ASSA-B.ST',{**assa,'operating_Income':25690},'2026-09-30') is None
    assert abb_r12_ebita('ASSA-A.ST',assa,'2026-09-30') is None
    assert 'fcff' not in abb_r12_ebita('ASSA-B.ST',{**assa,'net_Debt':57039},'2026-09-30')
    assert 'fcff' not in abb_r12_ebita('ASSA-B.ST',{**assa,'cash_Flow_From_Operating_Activities':24662},'2026-09-30')
    sand={'currency':'SEK','net_Sales':137625,'revenues':137625,'operating_Income':26376,
          'cash_Flow_From_Operating_Activities':21904,'net_Debt':36949}
    sandvik=abb_r12_ebita('SAND.ST',sand,'2026-09-30')
    assert sandvik['value']==28195e6
    assert abs(sandvik['fcff']-(26376*(1-6176/25261)+7917-4808-7680)*1e6)<1
    assert abb_r12_ebita('SAND.ST',{**sand,'currency':'EUR'},'2026-09-30') is None
    assert abb_r12_ebita('SAND.ST',sand,'2026-06-30') is None
    assert abb_r12_ebita('SAND.ST',{**sand,'operating_Income':27514},'2026-09-30') is None
    assert abb_r12_ebita('SAND.ST',{**sand,'net_Sales':120680},'2026-09-30') is None
    assert 'fcff' not in abb_r12_ebita('SAND.ST',{**sand,'net_Debt':30387},'2026-09-30')
    assert 'fcff' not in abb_r12_ebita('SAND.ST',{**sand,'cash_Flow_From_Operating_Activities':14499},'2026-09-30')
    lifco={'currency':'SEK','net_Sales':29822,'revenues':29822,'operating_Income':5597,
           'cash_Flow_From_Operating_Activities':5500,'net_Debt':12321}
    matched=abb_r12_ebita('LIFCO-B.ST',lifco,'2026-09-30')
    assert matched['value']==6764e6
    assert abs(matched['fcff']-(5597*(1-0.206)+1987-80-420)*1e6)<1
    assert abb_r12_ebita('LIFCO-B.ST',{**lifco,'currency':'EUR'},'2026-09-30') is None
    assert abb_r12_ebita('LIFCO-B.ST',lifco,'2026-06-30') is None
    assert abb_r12_ebita('LIFCO-B.ST',{**lifco,'operating_Income':6790},'2026-09-30') is None
    assert abb_r12_ebita('LIFCO-B.ST',{**lifco,'net_Sales':28251},'2026-09-30') is None
    assert 'fcff' not in abb_r12_ebita('LIFCO-B.ST',{**lifco,'net_Debt':9603},'2026-09-30')
    assert 'fcff' not in abb_r12_ebita('LIFCO-B.ST',{**lifco,'cash_Flow_From_Operating_Activities':3514},'2026-09-30')
    vol={'currency':'SEK','net_Sales':4594,'revenues':4594,'operating_Income':155,
         'cash_Flow_From_Operating_Activities':566,'net_Debt':1669}
    matched=abb_r12_ebita('VOLO.ST',vol,'2026-09-30')
    assert matched['value']==245e6
    assert abs(matched['fcff']-(155*(1-0.206)+306-12+85-97)*1e6)<1
    assert abb_r12_ebita('VOLO.ST',{**vol,'net_Sales':6500},'2026-09-30') is None
    assert abb_r12_ebita('VOLO.ST',{**vol,'currency':'EUR'},'2026-09-30') is None
    assert abb_r12_ebita('VOLO.ST',{**vol,'operating_Income':235},'2026-09-30') is None
    assert abb_r12_ebita('VOLO.ST',vol,'2026-06-30') is None
    assert 'fcff' not in abb_r12_ebita('VOLO.ST',{**vol,'net_Debt':1331},'2026-09-30')
    assert 'fcff' not in abb_r12_ebita('VOLO.ST',{**vol,'cash_Flow_From_Operating_Activities':321},'2026-09-30')
    lag={'currency':'SEK','revenues':11576,'net_Sales':11576,'operating_Income':1838,
         'cash_Flow_From_Operating_Activities':1498,'net_Debt':5190}
    matched=abb_r12_ebita('LAGR-B.ST',lag,'2026-09-30')
    assert matched['value']==2080e6
    assert abs(matched['fcff_tax_rate']-335/1667)<1e-10
    assert abs(matched['fcff']-(1838*(1-335/1667)+614-226-159-299+152)*1e6)<1
    for wrong in ({'currency':'EUR'},{'operating_Income':1900},{'net_Sales':11000}):
        assert abb_r12_ebita('LAGR-B.ST',{**lag,**wrong},'2026-09-30') is None
    assert abb_r12_ebita('LAGR-B.ST',lag,'2026-06-30') is None
    assert 'fcff' not in abb_r12_ebita('LAGR-B.ST',{**lag,'net_Debt':4659},'2026-09-30')
    assert 'fcff' not in abb_r12_ebita('LAGR-B.ST',{**lag,'cash_Flow_From_Operating_Activities':1300},'2026-09-30')
    # No secondary provider may replace the annual frames passed by fetch().
    original_roic=roic_company_data
    original_reports=borsapi_reports
    def forbidden(*args):
        raise AssertionError('valuation_engine must not re-query annual statements')
    try:
        globals()['roic_company_data']=forbidden
        globals()['borsapi_reports']=forbidden
        preserved=valuation_engine('VOLO-TEST.ST',{**info,'sector':'Industrials'},fin,cash,balance,{},pd.DataFrame())
        assert preserved['operating_inputs']==v['operating_inputs']
        assert preserved['target']==v['target']
    finally:
        globals()['roic_company_data']=original_roic
        globals()['borsapi_reports']=original_reports
    supplemental=pd.DataFrame({p:[200,20],q:[175,18]},index=['Operating Income','Capital Expenditure'])
    merged=fill_matching_statement_periods(fin,supplemental,fin.columns)
    assert merged.at['Operating Income',p]==90e6
    assert merged.at['Capital Expenditure',p]==20
    assert q not in merged.columns
    assert comparable_yahoo_periods(fin,supplemental).empty
    supplemental.at['Operating Income',p]=90e6
    assert comparable_yahoo_periods(fin,supplemental).empty
    supplemental.loc['Total Revenue',p]=1000e6
    assert list(comparable_yahoo_periods(fin,supplemental))==[p]
    print('Mathematical regression tests passed')

if __name__=='__main__':
    tests()
    if '--live' not in sys.argv:
        print('Live coverage was not checked. Pass --live to check production adapters.')
        sys.exit(0)
    if not borsapi_key():
        print('Live coverage requires BORSAPI_API_KEY.',file=sys.stderr)
        sys.exit(2)
    globals().update({name:lru_cache(maxsize=512)(adapter) if name not in ('verified_capitalization','borsapi_annual_frames') else adapter for name,adapter in production_adapters.items()})
    globals()['ecb_exchange_rates']=lru_cache(maxsize=1)(ecb_exchange_rates)
    ALL_COMPANIES={name:sym for name,sym in ALL_COMPANIES.items() if sym not in {'FNOX.ST','COALA.ST','ADDV-B.ST','ADDV-A.ST'}}
    if '--probe' in sys.argv:
        for symbol in ('ABB.ST','VOLO.ST','AAK.ST','BOL.ST','ATCO-B.ST','TRUE-B.ST','BOKUS.ST','BRAV.ST','ADDT-B.ST','INSTAL.ST','ALIF-B.ST','LAGR-B.ST','LIFCO-B.ST','SAND.ST','ASSA-B.ST'):
            key=borsapi_key()
            info,hist,fin,cash,balance,quick,errors=fetch(symbol)
            annual,_=borsapi_reports(symbol,key)
            trailing=borsapi_r12_records(symbol,key)
            model=valuation_engine(symbol,info,fin,cash,balance,quick,hist)
            annual_records=annual.get('periods',[])
            print(symbol,'annual',len(annual_records),'R12',len(trailing),'prices',len(hist),
                  'report currency',info.get('financialCurrency'),'listing currency',info.get('currency'),
                  'EPS present',numeric(info.get('trailingEps')) is not None,'P/E present',numeric(model.get('pe_ttm')) is not None,
                  'dividend yield present',trailing_dividend_yield(symbol,info.get('currentPrice')) is not None,
                  'report source',info.get('_report_source'),'market cap present',info.get('_cap_verified'),'profile',model.get('profile'),
                  'model names',list(model.get('models',{})),'issue count',len(model.get('missing',[])),'provider error labels',list(errors),
                  flush=True)
            # Diagnose reserve coverage without printing restricted financial values.
            if trailing:
                record=trailing[0]
                period=pd.Timestamp(record.get('report_End_Date') or record.get('reportEndDate')).normalize()
                listing=yf.Ticker(symbol)
                try:
                    profile=listing.get_info()
                    identity_ok=(str(profile.get('symbol') or '').upper()==symbol.upper() and
                                 str(profile.get('currency') or '').upper()==info.get('currency') and
                                 str(profile.get('financialCurrency') or '').upper()==info.get('financialCurrency'))
                    print(symbol,'Yahoo identity match',identity_ok,flush=True)
                    for attrs,is_ttm in ((('quarterly_income_stmt','quarterly_cashflow'),False),
                                         (('ttm_income_stmt','ttm_cashflow'),True)):
                        frames=[]
                        for attr in attrs:
                            try:
                                frame=getattr(listing,attr).copy()
                                frame.columns=pd.to_datetime(frame.columns,errors='coerce',utc=True).tz_localize(None).normalize()
                                frame=frame.loc[:,frame.columns.notna() & ~frame.columns.duplicated(keep=False)]
                            except Exception:
                                frame=pd.DataFrame()
                            frames.append(frame)
                        needed=pd.DatetimeIndex([period] if is_ttm else [period-pd.offsets.MonthEnd(3*i) for i in range(4)])
                        bridge=yahoo_r12_bridge(*frames,record,period,trailing=is_ttm) if identity_ok else {}
                        print(symbol,'Yahoo',attrs[0],'income periods',int(needed.isin(frames[0].columns).sum()),
                              'cash periods',int(needed.isin(frames[1].columns).sum()),
                              'income labels',sorted(str(x) for x in frames[0].index if any(k in str(x).lower() for k in ('revenue','operating income','amort','ebita'))),
                              'cash labels',sorted(str(x) for x in frames[1].index if any(k in str(x).lower() for k in ('amort','cash flow','capital expenditure','working capital'))),
                              'matched components',sorted(bridge),flush=True)
                except Exception as exc:
                    print(symbol,'Yahoo diagnostic failure',type(exc).__name__,flush=True)
            if symbol in ('ABB.ST','LIFCO-B.ST','VOLO.ST','BOKUS.ST','BRAV.ST','ADDT-B.ST','INSTAL.ST','ATCO-B.ST','ALIF-B.ST','LAGR-B.ST','SAND.ST','ASSA-B.ST'):
                if trailing:
                    record=trailing[0]
                    period=str(record.get('report_End_Date') or record.get('reportEndDate'))[:10]
                    matched=abb_r12_ebita(symbol,record,period)
                    print(symbol,'official filing matched',bool(matched),'R12 period',period,
                          'R12 reporting currency',record.get('currency'),'matched R12 EBITA evidence',bool(model.get('r12_ebita_evidence')),flush=True)
                else: print(symbol,'no R12 record to reconcile',flush=True)
            if symbol=='ABB.ST':
                for event_kind in ('calendar','insider'):
                    event_rows,event_error=borsapi_instrument_events(symbol,key,event_kind)
                    print('ABB.ST',event_kind,'rows',len(event_rows),'status',event_error or 'OK',flush=True)
                    if event_error: raise AssertionError(event_error)
                for period,records in (('annual',annual_records),('R12',trailing)):
                    print(period,'response fields', sorted(records[0]) if records else [],flush=True)
        sys.exit(0)
    rows=[]
    failures=[]
    for symbol in dict.fromkeys(ALL_COMPANIES.values()):
        try:
            info,hist,fin,cash,balance,quick,errors=fetch(symbol)
            v=valuation_engine(symbol,info,fin,cash,balance,quick,hist)
            if v.get('target') is not None:
                assert math.isfinite(v['target']) and v['target']>0, 'Invalid target'
                assert abs(v['target']-sum(v['models'].values())/len(v['models']))<1e-7, 'Inconsistent model target'
                if v.get('price') and v['price']>0:
                    assert abs(v['buy_limit']-v['target']*.8)<1e-7, 'Inconsistent buy limit'
                    assert abs(v['upside']-(v['target']/v['price']-1)*100)<1e-7, 'Inconsistent upside'
            avanza=info.get('_avanza_snapshot') or {}
            rows.append({'avanza_identity_verified':bool(avanza.get('verified')),'avanza_fresh_quote':bool(avanza.get('verified') and numeric(avanza.get('price')) is not None and avanza['price']>0),'ticker':symbol,'category':v.get('category'),'target':v.get('target'),'price':v.get('price'),'method':v.get('method'),'period':v.get('period'),'inputs':str(v.get('inputs')),'models':str(v.get('models')),'missing':str(v.get('missing')),'source':v.get('source','Source not recorded'),'fetch_errors':str(errors),'annual_reports':not fin.empty,'price_history':not hist.empty,'report_verified':False,'status':v.get('status'),'audit':str(v.get('audit',[])),'finite_positive_target':bool(v.get('target') is not None and math.isfinite(v['target']) and v['target']>0)})
        except Exception as e:
            failures.append(symbol+': '+str(e))
            rows.append({'ticker':symbol,'missing':type(e).__name__+': '+str(e),'report_verified':False})
        if rows[-1].get('category'):
            inputs=v.get('operating_inputs',v.get('inputs',{}))
            earnings=inputs.get('EBITA')
            if earnings is None: earnings=inputs.get('EBITA (beräknad)')
            debt,fx,shares,price=(inputs.get('Nettoskuld'),inputs.get('Valutakurs'),inputs.get('Antal aktier'),v.get('price'))
            applicable=v.get('category') in ('Rörelsedrivande bolag','Etablerade läkemedelsbolag')
            multiple=(price*shares/fx+debt)/earnings if applicable and all(x is not None for x in (earnings,debt,fx,shares,price)) and min(earnings,fx,shares,price)>0 else None
            r12_evidence=v.get('r12_ebita_evidence') or {}
            rows[-1].update({'r12_ebita_matched':bool(r12_evidence),'r12_ebita_source':r12_evidence.get('source'),
                             'model_period':v.get('operating_period',v.get('period')),
                             'ebita':earnings,'ev_ebita':multiple,'ebita_status':'Ej tillämpligt' if not applicable else 'Saknas EBITA' if earnings is None else 'Saknas jämförbar EV/EBITA' if multiple is None else 'Underlag finns',
                             'fcff_status':'Ej tillämpligt' if not applicable else 'Saknas FCFF' if numeric(inputs.get('FCFF')) is None else 'FCFF <= 0' if inputs['FCFF']<=0 else 'Underlag finns',
                             'dual_model_status':'Ej tillämpligt' if not applicable else 'Underlag finns' if {'EV/EBITA','DCF'}<=set(v.get('models',{})) and v.get('target') is not None else 'Ofullständig huvudmodell'})
        print(symbol,'target present',bool(rows[-1].get('finite_positive_target')),rows[-1].get('ebita_status'),rows[-1].get('fcff_status'),rows[-1].get('dual_model_status'),'matched R12 EBITA',bool(rows[-1].get('r12_ebita_matched')),flush=True)
        pd.DataFrame(rows).to_csv('valuation-audit.csv',index=False)
    print('Total:',len(rows),'With target:',sum(r.get('target') is not None for r in rows))

    print('EBITA coverage:',pd.Series([r.get('ebita_status','Fel') for r in rows]).value_counts().to_dict(),flush=True)
    print('FCFF coverage:',pd.Series([r.get('fcff_status','Fel') for r in rows]).value_counts().to_dict(),flush=True)
    print('Dual model coverage:',pd.Series([r.get('dual_model_status','Fel') for r in rows]).value_counts().to_dict(),flush=True)
    print('Matched R12 EBITA evidence (separate from model):',sum(bool(r.get('r12_ebita_matched')) for r in rows),flush=True)
    print('Live Avanza identity coverage:',sum(bool(r.get('avanza_identity_verified')) for r in rows),'of',len(rows),flush=True)
    print('Live Avanza fresh quote coverage:',sum(bool(r.get('avanza_fresh_quote')) for r in rows),'of',len(rows),flush=True)
    print('Execution failures:',len(failures),flush=True)
    if not any(row.get('annual_reports') for row in rows) or not any(row.get('price_history') for row in rows):
        raise AssertionError('No live annual reports or price history were loaded')
    if not any(row.get('finite_positive_target') for row in rows):
        raise AssertionError('No live valuation produced a target')
    if failures: raise AssertionError('; '.join(failures))
    # Successful execution is not equivalent to complete valuation coverage.
    expected=set(ALL_COMPANIES.values())
    actual={r['ticker'] for r in rows}
    assert actual==expected and len(rows)==len(expected), 'Incomplete active universe audit'
    missing_targets=[r['ticker'] for r in rows if not r.get('finite_positive_target')]
    missing_dual=[r['ticker'] for r in rows if r.get('dual_model_status') not in ('Ej tillämpligt','Underlag finns')]
    missing_identity=[r['ticker'] for r in rows if not r.get('avanza_identity_verified')]
    missing_quotes=[r['ticker'] for r in rows if not r.get('avanza_fresh_quote') or
                    numeric(r.get('price')) is None or r['price']<=0]
    incomplete=bool(missing_targets or missing_dual or missing_identity or missing_quotes)
    print('Publication coverage gate:', 'NOT READY' if incomplete else 'PASS',flush=True)
    print('Missing positive targets:',len(missing_targets),'Missing applicable dual models:',len(missing_dual),
          'Missing verified identities:',len(missing_identity),'Missing fresh quotes:',len(missing_quotes),flush=True)
    if '--require-complete' in sys.argv and incomplete:
        raise AssertionError('Publication coverage gate failed: incomplete targets, models, identities or prices')
