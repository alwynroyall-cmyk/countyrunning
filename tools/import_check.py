import sys, traceback
sys.path.insert(0, r'd:/wiltshireroadleague')
try:
    import league_scorer.graphical.qt.dashboard as ep
    print('Imported OK')
except Exception:
    traceback.print_exc()
