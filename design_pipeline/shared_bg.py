"""18-mers found in two or more genes, counted over every RefSeq transcript (UTRs and non-coding RNAs included).

Usage: python shared_bg.py <species_code>     reads <sp>_bg.pkl, writes <sp>_shared18.npy
"""
import numpy as np, pickle, sys, os, time
from kmers import enc, kcodes

sp = sys.argv[1]; K = 18; NB = 64; SH = np.uint64(2 * K - 6); t = time.time()
bg = pickle.load(open(f'{sp}_bg.pkl', 'rb'))
d = f'{sp}_bk'; os.makedirs(d, exist_ok=True)
for f in os.listdir(d): os.remove(f'{d}/{f}')
buf = [[] for _ in range(NB)]; nbuf = 0; total = 0
def flush():
    for b in range(NB):
        if buf[b]:
            with open(f'{d}/c{b:02d}', 'ab') as fh: np.concatenate(buf[b]).tofile(fh)
            buf[b].clear()
for seqs in bg.values():
    u = np.unique(np.concatenate([kcodes(enc(s), K) for s in seqs if len(s) >= K] or [np.zeros(0, np.uint64)]))
    if not len(u): continue
    total += len(u)
    bk = (u >> SH).astype(np.int64); cuts = np.searchsorted(bk, np.arange(NB + 1))
    for b in range(NB):
        if cuts[b + 1] > cuts[b]: buf[b].append(u[cuts[b]:cuts[b + 1]])
    nbuf += len(u)
    if nbuf > 30_000_000: flush(); nbuf = 0
flush(); del bg
out = []
for b in range(NB):
    p = f'{d}/c{b:02d}'
    if not os.path.exists(p): continue
    c = np.fromfile(p, np.uint64); c.sort()
    out.append(np.unique(c[1:][c[1:] == c[:-1]]))      # each gene contributes an 18-mer at most once
    os.remove(p)
S = np.concatenate(out); S.sort(); np.save(f'{sp}_shared18.npy', S)
print(sp, 'gene-distinct 18-mers', total, '| shared between genes', len(S), f'{time.time() - t:.0f}s')
