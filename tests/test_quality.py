import ast
import math
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
    def test_numeric(self):
        numeric = load_function("numeric", {"math": math, "pd": pd})
        for x in (None, float("nan"), float("inf"), -float("inf"), "bad"):
            self.assertIsNone(numeric(x))
        self.assertEqual(numeric("0.97"), 0.97)

    def test_dividend_yield(self):
        records=[{'currency':'SEK','dividend':0.11}]
        calc=load_function("trailing_dividend_yield",{"borsapi_key":lambda:'fixture',"borsapi_r12_records":lambda *args:records,"bd_field":lambda r,k:r.get(k)})
        self.assertAlmostEqual(calc("STOR-B.ST",11.10),0.11/11.10*100)
        records[0]['currency']='EUR'
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

    def test_emission(self):
        calc = load_function("emission_adjustment", {})
        result = calc(100, 200, 10, 20, 500, 1000)
        self.assertEqual(result["Totalt antal aktier"], 120)
        self.assertEqual(result["Justerad nettoskuld"], 320)
        self.assertEqual(result["TERP"], 10)

if __name__ == "__main__":
    unittest.main()
