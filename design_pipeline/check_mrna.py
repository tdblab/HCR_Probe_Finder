"""Verify a random sample of designed pairs: position, isoform coverage, transcript-part label, gene specificity,
and that every CDS boundary agrees with the CDS file (the CDS really sits at the recorded place in the mRNA).

Usage: python check_mrna.py <sp> [genes to sample]
"""
import json, pickle, glob, base64, random, sys, statistics as st, collections
import numpy as np
from kmers import enc, kcodes

sp = sys.argv[1]; NS = int(sys.argv[2]) if len(sys.argv) > 2 else 400
G = pickle.load(open(f'{sp}_genes.pkl', 'rb')); S18 = np.load(f'{sp}_shared18.npy')
D = {}
for f in glob.glob(f'{sp}_out/*.jsonl'):
    for l in open(f):
        d = json.loads(l); D[d['k']] = d
def unpack(b):
    v = int.from_bytes(base64.b64decode(b), 'big'); return ''.join('ACGT'[(v >> (2 * (51 - i))) & 3] for i in range(52))
def parts_of(a, e, c):
    m = 0
    if a < c[0]: m |= 1
    if e > c[0] and a < c[1]: m |= 2
    if e > c[1]: m |= 4
    return m
err = collections.Counter(); n = 0
random.seed(7)
for k in random.sample(sorted(D), min(NS, len(D))):
    d, g = D[k], G[k]
    for t, q, mask, i0, p, parts in d['p']:
        t = unpack(t); n += 1; m = int(mask, 16); want = 0
        if g['iso'][i0]['seq'][p:p + 52] != t: err['position'] += 1
        for i, iso in enumerate(g['iso']):
            j = iso['seq'].find(t)
            if bool(m >> i & 1) != (j >= 0): err['isoform coverage'] += 1
            if j >= 0: want |= parts_of(j, j + 52, iso['cds'])
        if want != parts: err['part label'] += 1
        c = kcodes(enc(t), 18); x = np.searchsorted(S18, c); x[x == len(S18)] = 0
        hit = S18[x] == c
        if hit[:8].any() or hit[27:35].any(): err['shared 18-mer in an arm'] += 1
    for iso, rec in zip(g['iso'], d['iso']):
        a, e = rec[6], rec[7]
        if iso['seq'][a:e] != iso['seq'][iso['cds'][0]:iso['cds'][1]]: err['CDS boundary'] += 1
pp = [sum(1 for p in d['p'] if p[5] == 2) for d in D.values()]
allp = [p[5] for d in D.values() for p in d['p']]
c = collections.Counter(allp)
print(f'{sp}: {len(D)} genes; checked {n} pairs in {min(NS, len(D))} genes; errors: {dict(err) or "none"}')
print(f'   CDS pairs per gene: median {st.median(pp)}, none {sum(x == 0 for x in pp)}, >=10 {sum(x >= 10 for x in pp)}, >=20 {sum(x >= 20 for x in pp)}')
print('   pairs by location:', {('5UTR' if b == 1 else 'CDS' if b == 2 else '3UTR' if b == 4 else f'span{b}'): v for b, v in sorted(c.items())})
multi = [d for d in D.values() if len(d['iso']) > 1]
utr_only = sum(1 for d in multi for u in d['u'] if u and all(x[2] in (1, 4) for x in u))
print(f'   multi-isoform genes {len(multi)}; with unique regions {sum(1 for d in multi if any(d["u"]))}; '
      f'isoforms whose unique regions lie only in UTRs {utr_only}')
