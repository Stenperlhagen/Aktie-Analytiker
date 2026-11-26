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
import sys
from functools import lru_cache

# Load exact production definitions without executing Streamlit UI.
tree=ast.parse(pathlib.Path('app.py').read_text(encoding='utf-8'))
names={'emission_adjustment','investment_score','show','roic_company_data','GROUPS','CAT','ADDITIONAL_COMPANIES','ALL_COMPANIES','INVESTMENT_TICKERS','BANK_TICKERS','INSURANCE_TICKERS','PROPERTY_TICKERS','VOLATI_Q3_2026','REPORTED_NAV','VALUATION_PROFILES','PROFILE_TICKERS','PROFILE_BY_SYMBOL','RESEARCH_TICKERS','HYBRID_PHARMA_TICKERS','ESTABLISHED_PHARMA_TICKERS','KNOWN_GROUP_BY_TICKER','ADDITIONAL_GROUP_FALLBACK','numeric','val','fx_rate','ecb_exchange_rates','borsapi_instruments','borsapi_instrument','borsapi_instrument_events','borsapi_reports','borsapi_history','borsapi_annual_frames','borsapi_r12_records','borsapi_r12_response','request_borsdata','borsapi_market_cap_map','bd_field','abb_r12_ebita','reported_operating_metrics','trailing_dividend_yield','report_ebita_extract','avanza_market_snapshot','verified_capitalization','fetch','operating_valuation_profile','comparable_ev_ebita','valuation_category','valuation_engine'}
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
OPERATING_CATEGORIES={'Rörelsedrivande bolag','Etablerade läkemedelsbolag'}

def has_complete_operating_models(model, info, annual_reports, price_history):
    """Coverage gate only; not independent source verification."""
    if model.get('category') not in OPERATING_CATEGORIES:
        return False
    models=model.get('models') or {}
    values=[numeric(models.get(key)) for key in ('EV/EBITA','DCF')]
    inputs=model.get('inputs') or {}
    required=[numeric(inputs.get(key)) for key in ('FCFF','Antal aktier','Valutakurs')]
    earnings=numeric(inputs.get('EBITA'))
    if earnings is None:
        earnings=numeric(inputs.get('EBITA (beräknad)'))
    return bool(annual_reports and price_history and info.get('_cap_verified')
                and all(v is not None and v>0 for v in values+required+[earnings,numeric(model.get('target'))])
                and numeric(inputs.get('Nettoskuld')) is not None)

def tests():
    sample={'category':'Rörelsedrivande bolag','models':{'EV/EBITA':90,'DCF':110},'target':100,'inputs':{'EBITA':100,'FCFF':50,'Antal aktier':10,'Valutakurs':1,'Nettoskuld':0}}
    assert has_complete_operating_models(sample,{'_cap_verified':True},True,True)
    assert not has_complete_operating_models({**sample,'models':{'EV/EBITA':90}}, {'_cap_verified':True},True,True)
    assert not has_complete_operating_models({**sample,'models':{'EV/EBIT':90,'DCF':110}}, {'_cap_verified':True},True,True)
    assert not has_complete_operating_models(sample,{'_cap_verified':False},True,True)
    assert not has_complete_operating_models(sample,{'_cap_verified':True},False,True)
    assert not has_complete_operating_models({**sample,'inputs':{**sample['inputs'],'FCFF':None}}, {'_cap_verified':True},True,True)
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
    assert abs(v['target']-sum(v['models'].values())/len(v['models']))<1e-8
    assert valuation_category('AAK.ST')=='Rörelsedrivande bolag'
    assert operating_valuation_profile('AAK.ST')=='Livsmedelsproduktion'
    assert comparable_ev_ebita(4777e6,3134e6,259708000,11,1)>0
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
    assert abb_r12_ebita('ABB.ST',{'currency':'USD','operating_Income':6716},'2026-09-30')['value']==6911e6
    assert abb_r12_ebita('ABB.ST',{'currency':'USD','operating_Income':6000},'2026-09-30') is None
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
        for symbol in ('ABB.ST','VOLO.ST','AAK.ST','BOL.ST','ATCO-A.ST','TRUE-B.ST'):
            key=borsapi_key()
            info,hist,fin,cash,balance,quick,errors=fetch(symbol)
            annual,_=borsapi_reports(symbol,key)
            trailing=borsapi_r12_records(symbol,key)
            model=valuation_engine(symbol,info,fin,cash,balance,quick,hist)
            annual_records=annual.get('periods',[])
            print(symbol,'annual',len(annual_records),'R12',len(trailing),'prices',len(hist),
                  'report currency',info.get('financialCurrency'),'listing currency',info.get('currency'),
                  'EPS',info.get('trailingEps'),'P/E',model.get('pe_ttm'),
                  'dividend yield',trailing_dividend_yield(symbol,info.get('currentPrice')),
                  'reported operating metrics',reported_operating_metrics(trailing[0] if trailing else {},info.get('financialCurrency'),info.get('currency'),info.get('marketCap')),
                  'market cap present',info.get('_cap_verified'),'profile',model.get('profile'),
                  'inputs',model.get('inputs'),'issues',model.get('missing'),'provider status',errors,
                  flush=True)
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
        v={}
        try:
            info,hist,fin,cash,balance,quick,errors=fetch(symbol)
            v=valuation_engine(symbol,info,fin,cash,balance,quick,hist)
            if v.get('target') is not None:
                assert math.isfinite(v['target']) and v['target']>0, 'Invalid target'
                assert abs(v['target']-sum(v['models'].values())/len(v['models']))<1e-7, 'Inconsistent model target'
                if v.get('price') and v['price']>0:
                    assert abs(v['buy_limit']-v['target']*.8)<1e-7, 'Inconsistent buy limit'
                    assert abs(v['upside']-(v['target']/v['price']-1)*100)<1e-7, 'Inconsistent upside'
            rows.append({'ticker':symbol,'category':v.get('category'),'target':v.get('target'),'price':v.get('price'),'method':v.get('method'),'period':v.get('period'),'inputs':str(v.get('inputs')),'models':str(v.get('models')),'missing':str(v.get('missing')),'source':v.get('source','Source not recorded'),'fetch_errors':str(errors),'annual_reports':not fin.empty,'price_history':not hist.empty,'report_verified':False,'both_models_present':has_complete_operating_models(v,info,not fin.empty,not hist.empty),'status':v.get('status'),'audit':str(v.get('audit',[])),'finite_positive_target':bool(v.get('target') is not None and math.isfinite(v['target']) and v['target']>0)})
        except Exception as e:
            failures.append(symbol+': '+str(e))
            rows.append({'ticker':symbol,'missing':type(e).__name__+': '+str(e),'report_verified':False})
        if v and rows[-1].get('category'):
            inputs=v.get('inputs',{})
            earnings=inputs.get('EBITA')
            if earnings is None: earnings=inputs.get('EBITA (beräknad)')
            debt,fx,shares,price=(inputs.get('Nettoskuld'),inputs.get('Valutakurs'),inputs.get('Antal aktier'),v.get('price'))
            applicable=v.get('category') in OPERATING_CATEGORIES
            multiple=(price*shares/fx+debt)/earnings if applicable and all(x is not None for x in (earnings,debt,fx,shares,price)) and min(earnings,fx,shares,price)>0 else None
            rows[-1].update({'ebita':earnings,'ev_ebita':multiple,'ebita_status':'Ej tillämpligt' if not applicable else 'Saknas EBITA' if earnings is None else 'Saknas jämförbar EV/EBITA' if multiple is None else 'Underlag finns'})
        print(symbol,rows[-1].get('target'),rows[-1].get('ebita_status'),flush=True)
        pd.DataFrame(rows).to_csv('valuation-audit.csv',index=False)
    print('Total:',len(rows),'With target:',sum(r.get('target') is not None for r in rows))

    print('EBITA coverage:',pd.Series([r.get('ebita_status','Fel') for r in rows]).value_counts().to_dict(),flush=True)
    print('Execution failures:',len(failures),flush=True)
    if not any(row.get('annual_reports') for row in rows) or not any(row.get('price_history') for row in rows):
        raise AssertionError('No live annual reports or price history were loaded')
    if not any(row.get('finite_positive_target') for row in rows):
        raise AssertionError('No live valuation produced a target')
    if failures: raise AssertionError('; '.join(failures))
    applicable_rows=[r for r in rows if r.get('category') in OPERATING_CATEGORIES]
    missing_both=[r['ticker'] for r in applicable_rows if not r.get('both_models_present')]
    print('Both operating models present:',len(applicable_rows)-len(missing_both),'/',len(applicable_rows),flush=True)
    if not applicable_rows or missing_both:
        raise AssertionError('Incomplete live EV/EBITA + DCF coverage: '+', '.join(missing_both))
    print('Coverage passed; independent filing and share-class verification is still required.',flush=True)
