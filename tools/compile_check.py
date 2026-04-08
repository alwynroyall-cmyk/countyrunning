import py_compile
import traceback
p = r'd:/wiltshireroadleague/league_scorer/view_enquiry/enquiry_panel.py'
try:
    py_compile.compile(p, doraise=True)
    print('OK')
except Exception:
    traceback.print_exc()
