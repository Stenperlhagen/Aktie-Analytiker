import unittest
from streamlit.testing.v1 import AppTest

class AppSmokeTest(unittest.TestCase):
    def test_startup_without_company(self):
        app = AppTest.from_file("app.py", default_timeout=30).run()
        self.assertFalse(app.exception, str(app.exception))
        self.assertTrue(any(x.label == "Sök eller välj bolag" for x in app.selectbox))
        self.assertTrue(any(x.label == "Sök igenom alla bolag" for x in app.button))

if __name__ == "__main__":
    unittest.main()
