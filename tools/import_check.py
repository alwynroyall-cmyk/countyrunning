import sys,traceback
sys.path.insert(0, r'd:/wiltshireroadleague')
try:
    import league_scorer.view_enquiry.enquiry_panel as ep
    print('Imported OK')
except Exception:
    traceback.print_exc()
