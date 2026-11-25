import ast
import pathlib
import csv
import math
import os
import requests
import urllib.parse
import datetime as dt
import pandas as pd
import yfinance as yf
from functools import lru_cache

# Load exact production definitions without executing Streamlit UI.
tree=ast.parse(pathlib.Path('app.py').read_text(encoding='utf-8'))
names={'emission_adjustment','investment_score','show','roic_company_data','GROUPS','CAT','ADDITIONAL_COMPANIES','ALL_COMPANIES','INVESTMENT_TICKERS','BANK_TICKERS','INSURANCE_TICKERS','PROPERTY_TICKERS','VOLATI_Q3_2026','REPORTED_NAV','VALUATION_PROFILES','PROFILE_TICKERS','PROFILE_BY_SYMBOL','RESEARCH_TICKERS','HYBRID_PHARMA_TICKERS','ESTABLISHED_PHARMA_TICKERS','KNOWN_GROUP_BY_TICKER','ADDITIONAL_GROUP_FALLBACK','numeric','val','fx_rate','fetch','operating_valuation_profile','comparable_ev_ebita','valuation_category','valuation_engine'}
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

def borsapi_reports(symbol, api_key):
    # Deterministic offline fallback; integration is tested separately.
    return {}, 'Börsdata API unavailable in offline audit'

def borsapi_history(symbol, api_key):
    return pd.DataFrame(), 'Börsdata API unavailable in offline audit'

def borsapi_annual_frames(symbol, api_key):
    return None, 'Börsdata API unavailable in offline audit'

def borsapi_r12_records(symbol, api_key):
    # Offline audit: R12 endpoint is unavailable without credentials.
    return []

def tests():
    assert abs(comparable_ev_ebita(245e6,1331e6,79406571,12,1)-20.263)<0.01
    assert comparable_ev_ebita(100,None,10,12,1) is None
    assert comparable_ev_ebita(100,1300,10,12,1) is None
    p=pd.Timestamp('2025-12-31')
    fin=pd.DataFrame({p:[1000e6,100e6,90e6]},index=['Total Revenue','EBITA','Operating Income'])
    balance=pd.DataFrame({p:[200e6]},index=['Net Debt'])
    cash=pd.DataFrame({p:[20e6,-25e6,-5e6]},index=['Depreciation And Amortization','Capital Expenditure','Change In Working Capital'])
    info={'currency':'SEK','financialCurrency':'SEK','currentPrice':100,'sharesOutstanding':10e6,'marketCap':1e9}
    v=valuation_engine('VOLO-TEST.ST',{**info,'sector':'Industrials'},fin,cash,balance,{},pd.DataFrame())
    assert abs(v['models']['EV/EBITA']-90)<1e-8
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
    print('Mathematical regression tests passed')

if __name__=='__main__':
    tests()
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
            rows.append({'ticker':symbol,'category':v.get('category'),'target':v.get('target'),'price':v.get('price'),'method':v.get('method'),'period':v.get('period'),'inputs':str(v.get('inputs')),'models':str(v.get('models')),'missing':str(v.get('missing')),'source':v.get('source','Yahoo Finance – not independently verified'),'fetch_errors':str(list(errors)),'report_verified':False,'status':v.get('status'),'audit':str(v.get('audit',[])),'finite_positive_target':bool(v.get('target') is not None and math.isfinite(v['target']) and v['target']>0)})
        except Exception as e:
            failures.append(symbol+': '+str(e))
            rows.append({'ticker':symbol,'missing':type(e).__name__+': '+str(e),'report_verified':False})
        if 'v' in locals() and rows[-1].get('category'):
            inputs=v.get('inputs',{})
            earnings=inputs.get('EBITA')
            if earnings is None: earnings=inputs.get('EBITA (beräknad)')
            debt,fx,shares,price=(inputs.get('Nettoskuld'),inputs.get('Valutakurs'),inputs.get('Antal aktier'),v.get('price'))
            applicable=v.get('category')=='Rörelsedrivande bolag'
            multiple=(price*shares/fx+debt)/earnings if applicable and all(x is not None for x in (earnings,debt,fx,shares,price)) and min(earnings,fx,shares,price)>0 else None
            rows[-1].update({'ebita':earnings,'ev_ebita':multiple,'ebita_status':'Ej tillämpligt' if not applicable else 'Saknas EBITA' if earnings is None else 'Saknas jämförbar EV/EBITA' if multiple is None else 'Underlag finns'})
        print(symbol,rows[-1].get('target'),rows[-1].get('ebita_status'),flush=True)
        pd.DataFrame(rows).to_csv('valuation-audit.csv',index=False)
    print('Total:',len(rows),'With target:',sum(r.get('target') is not None for r in rows))

    print('EBITA coverage:',pd.Series([r.get('ebita_status','Fel') for r in rows]).value_counts().to_dict(),flush=True)
    print('Execution failures:',len(failures),flush=True)
    if failures: raise AssertionError('; '.join(failures))
