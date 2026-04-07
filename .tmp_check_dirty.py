from league_scorer.session_config import config as session_config
from pathlib import Path
out = session_config.output_dir
print('output_dir:', out)
if out is None:
    print('output_dir is None')
else:
    ap = Path(out) / 'autopilot' / 'dirty'
    rp = Path(out) / 'raes' / 'dirty'
    cj = Path(out) / 'raes' / 'changes.json'
    print('autopilot dirty exists:', ap.exists())
    print('raes dirty exists:', rp.exists())
    print('changes.json exists:', cj.exists())
    if cj.exists():
        try:
            txt = cj.read_text(encoding='utf-8')
            print('changes.json len:', len(txt))
            print('changes.json tail:', txt[-400:])
        except Exception as e:
            print('failed reading changes.json', e)
