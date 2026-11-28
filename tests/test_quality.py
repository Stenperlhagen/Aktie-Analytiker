import ast
import math
import re
import unittest
from pathlib import Path
from unittest.mock import Mock
import pandas as pd

SOURCE = Path("app.py").read_text(encoding="utf-8")
TREE = ast.parse(SOURCE)

def load_function(name, namespace):
    node = next(n for n in TREE.body if isinstance(n, ast.FunctionDef) and n.name == name)
    node.decorator_list = []
    exec(compile(ast.Module(body=[node], type_ignores=[]), "app.py", "exec"), namespace)
    return namespace[name]

class QualityTests(unittest.TestCase):
    def test_reported_nav_is_current_currency_matched_and_sourced(self):
        ns={"math":math,"pd":pd}
        load_function("numeric",ns)
        node=next(n for n in TREE.body if isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id=="REPORTED_NAV" for t in n.targets))
        ns["REPORTED_NAV"]=ast.literal_eval(node.value)
        calc=load_function("current_reported_nav",ns)
        for symbol,amount in (("BALD-B.ST",97.03),("CATE.ST",468.87),("WALL-B.ST",65.0)):
            result=calc(symbol,"SEK","2026-11-27")
            self.assertEqual(result["nav"],amount)
            self.assertEqual(result["date"],"2026-09-30")
            self.assertTrue(result["source"].startswith("https://"))
            self.assertIsNone(calc(symbol,"EUR","2026-11-27"))
            self.assertIsNone(calc(symbol,"SEK","2026-09-29"))
            self.assertIsNone(calc(symbol,"SEK","2027-04-01"))
        self.assertIsNone(calc("BALD-A.ST","SEK","2026-11-27"))
        ns["REPORTED_NAV"]["BAD.ST"]={"nav":10,"date":"2026-09-30","kind":"NAV"}
        self.assertIsNone(calc("BAD.ST","SEK","2026-11-27"))
        self.assertIn("nav=current_reported_nav(symbol,currency)",SOURCE)

    def test_screener_does_not_bypass_primary_r12_on_missing_target(self):
        calls=[]
        def engine(*args,**kwargs):
            calls.append(kwargs)
            if kwargs.get("use_roic") is False:
                return {"target":100,"category":"Rörelsedrivande bolag","missing":[]}
            return {"target":None,"category":"Rörelsedrivande bolag","missing":["FCFF <= 0"]}
        ns={"fetch":lambda symbol: ({},None,None,None,None,{},{}),
            "valuation_engine":engine,"roic_key":lambda:"configured",
            "growth_stability":lambda fin:None,"val":lambda *args:None,
            "valuation_category":lambda symbol:"Rörelsedrivande bolag"}
        calc=load_function("screener_snapshot",ns)
        result=calc("ABC.ST")
        self.assertIsNone(result["target"])
        self.assertEqual(calls,[{}])
        self.assertIn("FCFF <= 0",result["missing"])
        def failing(*args,**kwargs):
            calls.append(kwargs)
            raise RuntimeError("primary unavailable")
        ns["valuation_engine"]=failing
        self.assertIsNone(calc("ABC.ST")["target"])
        self.assertEqual(calls,[{},{}])

    def test_roko_r12_filing_includes_all_debt_and_unadjusted_ebita(self):
        ns={"math":math,"pd":pd,"re":re}
        load_function("numeric",ns)
        load_function("bd_field",ns)
        calc=load_function("abb_r12_ebita",ns)
        record={"currency":"SEK","net_Sales":7754,"operating_Income":1326,
                "cash_Flow_From_Operating_Activities":1329,"net_Debt":2811}
        period="2026-09-30"
        result=calc("ROKO-B.ST",record,period)
        self.assertEqual(result["value"],1602e6)
        self.assertEqual(result["net_debt"],5609e6)
        self.assertAlmostEqual(result["fcff"],(1326*(1-333/1196)+482+15-86)*1e6)
        for debt in (2044,4842,5609):
            self.assertEqual(calc("ROKO-B.ST",{**record,"net_Debt":debt},period)["net_debt"],5609e6)
        self.assertNotIn("net_debt",calc("ROKO-B.ST",{**record,"net_Debt":123},period))
        self.assertNotIn("fcff",calc("ROKO-B.ST",{**record,"net_Debt":123},period))
        self.assertNotIn("fcff",calc("ROKO-B.ST",{**record,"cash_Flow_From_Operating_Activities":1000},period))
        self.assertIsNone(calc("ROKO-B.ST",{**record,"currency":"EUR"},period))
        self.assertIsNone(calc("ROKO-B.ST",record,"2026-06-30"))
        self.assertIsNone(calc("ROKO-B.ST",{**record,"operating_Income":1000},period))

    def test_numeric(self):
        numeric = load_function("numeric", {"math": math, "pd": pd})
        for x in (None, float("nan"), float("inf"), -float("inf"), "bad"):
            self.assertIsNone(numeric(x))
        self.assertEqual(numeric("0.97"), 0.97)

    def test_dividend_yield(self):
        records=[{'currency':'SEK','dividend':0.11}]
        calc=load_function("trailing_dividend_yield",{"borsapi_key":lambda:'fixture',"borsapi_r12_records":lambda *args:records,"bd_field":lambda r,k:r.get(k),"fx_rate":lambda source,dest:1 if source==dest else 11 if source=="EUR" and dest=="SEK" else None})
        self.assertAlmostEqual(calc("STOR-B.ST",11.10),0.11/11.10*100)
        records[0]['currency']='EUR'
        self.assertAlmostEqual(calc("STOR-B.ST",11.10),0.11*11/11.10*100)
        records[0]['currency']='XXX'
        self.assertIsNone(calc("STOR-B.ST",11.10))
        self.assertIsNone(calc("STOR-B.ST", 0))
        self.assertIsNone(calc("STOR-B.ST", -1))

    def test_no_ambiguous_yield_field(self):
        self.assertNotIn("dividend_yield = val(info,'dividendYield')", SOURCE)

    def test_ev_ebita(self):
        ns = {"math": math, "pd": pd}
        load_function("numeric", ns)
        calc = load_function("comparable_ev_ebita", ns)
        self.assertAlmostEqual(calc(100, 200, 10, 12, 1), 100)
        self.assertAlmostEqual(calc(100, 200, 10, 12, 2), 200)
        for args in [(0,200,10,12,1),(100,None,10,12,1),(100,200,0,12,1),(100,200,10,12,0)]:
            self.assertIsNone(calc(*args))

    def test_instrument_resolver_preserves_exchange_and_share_class(self):
        instruments=[
            {"insId":1,"ticker":"ABC A","yahoo":"ABC-A.ST"},
            {"insId":2,"ticker":"ABC B","yahoo":"ABC-B.ST"},
            {"insId":3,"ticker":"ABC B","yahoo":"ABC-B.CO"},
            {"insId":4,"ticker":"ABCB","yahoo":"ABCB.ST"},
        ]
        resolver=load_function("borsapi_instrument", {
            "re":re,"borsapi_instruments":lambda key:(instruments,""),
        })
        self.assertEqual(resolver("ABC-B.ST","test")[0]["instrument_id"],"2")
        self.assertEqual(resolver("ABC-B.CO","test")[0]["instrument_id"],"3")
        self.assertFalse(resolver("ABC-B.OL","test")[0])
        instruments.append({"insId":5,"ticker":"ABC B","yahoo":"ABC-B.ST"})
        self.assertFalse(resolver("ABC-B.ST","test")[0])

    def test_fx_cross_rate_and_missing_currency(self):
        calc=load_function("fx_rate", {"ecb_exchange_rates":lambda:{"EUR":1,"SEK":11,"DKK":7.5}})
        self.assertAlmostEqual(calc("SEK","DKK"),7.5/11)
        self.assertEqual(calc("SEK","SEK"),1)
        self.assertIsNone(calc("SEK","USD"))

    def test_historical_pe_uses_report_publication_date(self):
        ns={"math":math,"pd":pd}
        load_function("numeric",ns)
        calc=load_function("reconstructed_pe_history",ns)
        period=pd.Timestamp("2025-06-30")
        prices=pd.DataFrame({"Close":[100,200]},index=pd.to_datetime(["2025-07-01","2025-08-21"]))
        income=pd.DataFrame({period:[10]},index=["Basic EPS"])
        self.assertEqual(calc(income,prices,"SEK","SEK",{period:pd.Timestamp("2025-08-20")}),[20])
        self.assertEqual(calc(income,prices,"SEK","SEK",{}),[])

    def test_overview_money_displays_stockholm_amounts_in_sek(self):
        ns={"math":math}
        load_function("numeric",ns)
        load_function("show",ns)
        metric=load_function("overview_money_metric",ns)
        self.assertEqual(metric("Omsättning",36_962e6,"USD","SEK",10),("Omsättning (MSEK)","369 620"))
        self.assertEqual(metric("EBITA",6_911e6,"USD","SEK",None),("EBITA (MUSD)","6 911"))
        self.assertEqual(metric("Kassa",0,"SEK","SEK",1),("Kassa (MSEK)","0"))

    def test_yahoo_r12_bridge_reconciles_all_four_quarters(self):
        ns={"math":math,"pd":pd,"re":re}
        load_function("numeric",ns)
        load_function("bd_field",ns)
        calc=load_function("yahoo_r12_bridge",ns)
        periods=pd.to_datetime(["2026-09-30","2026-06-30","2026-03-31","2025-12-31"])
        income=pd.DataFrame([[250e6]*4,[25e6]*4,[2e6]*4],index=["Total Revenue","Operating Income","Amortization Of Intangibles"],columns=periods)
        cash=pd.DataFrame([[30e6]*4,[5e6]*4,[-6e6]*4,[-1e6]*4],index=["Operating Cash Flow","Depreciation And Amortization","Capital Expenditure","Change In Working Capital"],columns=periods)
        record={"revenues":1000,"operating_Income":100,"cash_Flow_From_Operating_Activities":120}
        result=calc(income,cash,record,periods[0])
        self.assertAlmostEqual(result["ebita"],108e6)
        self.assertAlmostEqual(result["fcff"],71.4e6)
        self.assertEqual(calc(income.iloc[:,:3],cash,record,periods[0]),{})
        self.assertEqual(calc(income,cash,{**record,"revenues":1500},periods[0]),{})
        self.assertNotIn("fcff",calc(income,cash,{**record,"cash_Flow_From_Operating_Activities":200},periods[0]))
        self.assertNotIn("ebita",calc(income.drop(index="Amortization Of Intangibles"),cash,record,periods[0]))
        self.assertNotIn("fcff",calc(income,cash.drop(index="Change In Working Capital"),record,periods[0]))

    def test_yahoo_trailing_bridge_requires_dated_and_reconciled_flows(self):
        ns={"math":math,"pd":pd,"re":re}
        load_function("numeric",ns)
        load_function("bd_field",ns)
        calc=load_function("yahoo_r12_bridge",ns)
        end=pd.Timestamp("2026-09-30")
        income=pd.DataFrame({end:[1010e6,1000e6,100e6,8e6]},index=["Total Revenue","Operating Revenue","Operating Income","Amortization Of Intangibles Income Statement"])
        cash=pd.DataFrame({end:[120e6,20e6,-24e6,-4e6]},index=["Operating Cash Flow","Depreciation And Amortization","Capital Expenditure","Change In Working Capital"])
        record={"revenues":1100,"net_Sales":1000,"operating_Income":100,"cash_Flow_From_Operating_Activities":120}
        self.assertEqual(calc(income,cash,record,end),{})
        result=calc(income,cash,record,end,trailing=True)
        self.assertAlmostEqual(result["ebita"],108e6)
        self.assertAlmostEqual(result["fcff"],71.4e6)
        self.assertEqual(calc(income,cash,record,end-pd.offsets.MonthEnd(3),trailing=True),{})
        self.assertEqual(calc(income,cash,{**record,"net_Sales":900},end,trailing=True),{})
        self.assertNotIn("fcff",calc(income,cash,{**record,"cash_Flow_From_Operating_Activities":200},end,trailing=True))
        self.assertNotIn("ebita",calc(income.rename(index={"Amortization Of Intangibles Income Statement":"Amortization Cash Flow"}),cash,record,end,trailing=True))

    def test_yahoo_reconciles_net_sales_without_overwriting_revenue(self):
        ns={"math":math,"pd":pd,"re":re}
        load_function("numeric",ns)
        load_function("bd_field",ns)
        annual=load_function("comparable_yahoo_periods",ns)
        fill=load_function("fill_matching_statement_periods",ns)
        trailing=load_function("yahoo_r12_bridge",ns)
        end=pd.Timestamp("2026-09-30")
        primary=pd.DataFrame({end:[1100e6,1000e6,100e6]},index=["Total Revenue","Net Sales","Operating Income"])
        income=pd.DataFrame({end:[1000e6,100e6,8e6]},index=["Total Revenue","Operating Income","Amortization Of Intangibles"])
        cash=pd.DataFrame({end:[120e6,20e6,-24e6,-4e6]},index=["Operating Cash Flow","Depreciation And Amortization","Capital Expenditure","Change In Working Capital"])
        record={"revenues":1100,"net_Sales":1000,"operating_Income":100,"cash_Flow_From_Operating_Activities":120}
        self.assertEqual(list(annual(primary,income)),[end])
        merged=fill(primary,income,annual(primary,income))
        self.assertEqual(merged.at["Total Revenue",end],1100e6)
        self.assertEqual(merged.at["Amortization Of Intangibles",end],8e6)
        self.assertAlmostEqual(trailing(income,cash,record,end,trailing=True)["ebita"],108e6)
        wrong=income.copy()
        wrong.at["Operating Income",end]=90e6
        self.assertTrue(annual(primary,wrong).empty)
        self.assertEqual(trailing(wrong,cash,record,end,trailing=True),{})
        self.assertEqual(trailing(income,cash,{**record,"net_Sales":900},end,trailing=True),{})

    def test_yahoo_matches_reported_ebit_not_adjusted_ebit(self):
        ns={"math":math,"pd":pd,"re":re}
        load_function("numeric",ns)
        load_function("bd_field",ns)
        annual=load_function("comparable_yahoo_periods",ns)
        fill=load_function("fill_matching_statement_periods",ns)
        trailing=load_function("yahoo_r12_bridge",ns)
        end=pd.Timestamp("2026-09-30")
        primary=pd.DataFrame({end:[1000e6,100e6]},index=["Total Revenue","Operating Income"])
        income=pd.DataFrame({end:[1000e6,90e6,100e6,8e6]},index=["Total Revenue","Operating Income","Total Operating Income As Reported","Amortization Of Intangibles"])
        cash=pd.DataFrame({end:[120e6,20e6,-24e6,-4e6]},index=["Operating Cash Flow","Depreciation And Amortization","Capital Expenditure","Change In Working Capital"])
        record={"revenues":1000,"operating_Income":100,"cash_Flow_From_Operating_Activities":120}
        self.assertEqual(list(annual(primary,income)),[end])
        self.assertEqual(fill(primary,income,annual(primary,income)).at["Operating Income",end],100e6)
        result=trailing(income,cash,record,end,trailing=True)
        self.assertAlmostEqual(result["ebita"],108e6)
        self.assertAlmostEqual(result["fcff"],71.4e6)
        income.at["Total Operating Income As Reported",end]=80e6
        self.assertTrue(annual(primary,income).empty)
        self.assertEqual(trailing(income,cash,record,end,trailing=True),{})

    def test_afry_r12_fcff_requires_reconciled_lease_inclusive_debt(self):
        ns={"math":math,"pd":pd,"re":re}
        load_function("numeric",ns)
        load_function("bd_field",ns)
        calc=load_function("abb_r12_ebita",ns)
        record={"currency":"SEK","revenues":25295,"net_Sales":25276,
                "operating_Income":1448,"cash_Flow_From_Operating_Activities":2153,
                "net_Debt":6285}
        date="2026-09-30"
        result=calc("AFRY.ST",record,date)
        self.assertEqual(result["value"],1613e6)
        self.assertAlmostEqual(result["fcff"],(1448*(1-323/1182)+815+69+512-80)*1e6)
        self.assertIn("nettoinvesteringar",result["fcff_note"])
        self.assertNotIn("fcff",calc("AFRY.ST",{**record,"net_Debt":4843},date))
        self.assertNotIn("fcff",calc("AFRY.ST",{**record,"cash_Flow_From_Operating_Activities":2220},date))
        self.assertIsNone(calc("AFRY.ST",{**record,"currency":"EUR"},date))
        self.assertIsNone(calc("AFRY.ST",{**record,"operating_Income":1387},date))
        self.assertIsNone(calc("AFRY.ST",record,"2026-06-30"))

    def test_loomis_r12_fcff_uses_ifrs_cash_flows(self):
        ns={"math":math,"pd":pd,"re":re}
        load_function("numeric",ns)
        load_function("bd_field",ns)
        calc=load_function("official_r12_fcff",ns)
        record={"currency":"SEK","revenues":31253,"net_Sales":31253,
                "operating_Income":3651,"cash_Flow_From_Operating_Activities":5243}
        end="2026-09-30"
        result=calc("LOOMIS.ST",record,end)
        self.assertAlmostEqual(result["fcff"],(3651*(1-987/3009)+3241+871-645-1425)*1e6)
        self.assertIn("Hermes",result["fcff_note"])
        self.assertNotIn("ebita",result)
        self.assertIsNone(calc("LOOMIS.ST",{**record,"currency":"USD"},end))
        self.assertIsNone(calc("LOOMIS.ST",record,"2026-06-30"))
        self.assertIsNone(calc("LOOMIS.ST",{**record,"cash_Flow_From_Operating_Activities":3613},end))
        self.assertIsNone(calc("LOOMIS.ST",{**record,"operating_Income":4291},end))
        self.assertIsNone(calc("LOOMIS.ST",{**record,"net_Sales":30000},end))

    def test_emission(self):
        calc = load_function("emission_adjustment", {})
        result = calc(100, 200, 10, 20, 500, 1000)
        self.assertEqual(result["Totalt antal aktier"], 120)
        self.assertEqual(result["Justerad nettoskuld"], 320)
        self.assertEqual(result["TERP"], 10)

if __name__ == "__main__":
    unittest.main()
