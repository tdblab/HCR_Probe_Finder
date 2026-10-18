"""Probe-pair design on full mRNAs, with every pair labelled by the transcript parts it covers.

Usage: python design_mrna.py <sp> <first gene> <stop gene> [time limit s]    (resumes; rerun until all genes are written)

Pair record: [packed 52-nt target, quality 0-100, isoform mask (hex), isoform index, start in that isoform, parts]
  parts is a bit mask of the transcript parts the 52-nt target touches in any isoform: 1 = 5'UTR, 2 = CDS, 4 = 3'UTR.
Isoform record: [protein id, isoform label, length, identical transcripts, partial, transcript id,
                 CDS start (0-based), CDS end (exclusive), UTRs known (0/1), transcript variant]
Unique regions: per isoform, [start, end (exclusive), parts] of stretches whose 25-mers occur in no other isoform.
"""
import numpy as np, pickle, json, sys, os, time, re, base64
from kmers import enc, kcodes

sp = sys.argv[1]; start = int(sys.argv[2]); stop = int(sys.argv[3]); TL = float(sys.argv[4]) if len(sys.argv) > 4 else 240
G = pickle.load(open(f'{sp}_genes.pkl', 'rb')); S18 = np.load(f'{sp}_shared18.npy')
ARM, GAP = 25, 2; W = 2 * ARM + GAP
GCMIN, GCMAX, DGC = 0.30, 0.80, 0.30
MAX_ALL, MAX_SUB = 60, 15          # per isoform subset and per part combination
CAP_CDS, CAP_UTR = 100, 100        # per gene

def winsum(x, k):
    c = np.concatenate(([0], np.cumsum(x))); return c[k:] - c[:-k]

def site_ok(s, b):
    L = len(s); n = L - ARM + 1
    if n <= 0: return None, None
    bad = (b == 255); b = np.where(bad, 0, b)
    gc = winsum(((b == 1) | (b == 2)).astype(np.int32), ARM) / ARM
    eq = (b[1:] == b[:-1]).astype(np.int32)
    run5 = np.zeros(L, np.int32); run5[:L - 4] = winsum(eq, 4) == 4
    run6 = np.zeros(L, np.int32); run6[:L - 5] = winsum(eq, 5) == 5
    g4 = np.zeros(L, np.int32); g5 = np.zeros(L, np.int32)
    for m in re.finditer('(?=GGGG|CCCC)', s): g4[m.start()] = 1
    for m in re.finditer('(?=GGGGG|CCCCC)', s): g5[m.start()] = 1
    ok = (gc >= GCMIN) & (gc <= GCMAX) & (winsum(run6, ARM - 5)[:n] == 0) & (winsum(g5, ARM - 4)[:n] == 0)
    ok &= winsum(bad.astype(np.int32), ARM)[:n] == 0
    pen = 0.1 * (winsum(run5, ARM - 4)[:n] > 0) + 0.1 * (winsum(g4, ARM - 3)[:n] > 0)
    c18 = kcodes(b, 18); i = np.searchsorted(S18, c18); i[i == len(S18)] = 0
    ok &= winsum((S18[i] == c18).astype(np.int32), ARM - 18 + 1)[:n] == 0     # no 18-mer shared with another gene
    return ok, (gc, pen)

def pack(t):
    v = 0
    for ch in t: v = (v << 2) | 'ACGT'.index(ch)
    return base64.b64encode(v.to_bytes(13, 'big')).decode()

def parts_of(a, e, cds):
    s0, s1 = cds; m = 0
    if a < s0: m |= 1
    if e > s0 and a < s1: m |= 2
    if e > s1: m |= 4
    return m

def design(g):
    isos = g['iso']; cand = {}
    for ii, iso in enumerate(isos):
        s = iso['seq']
        if len(s) < W: continue
        b = enc(s); ok, gp = site_ok(s, b)
        if ok is None: continue
        gc, pen = gp; m = len(s) - W + 1
        v = ok[:m] & ok[ARM + GAP:ARM + GAP + m] & (np.abs(gc[:m] - gc[ARM + GAP:ARM + GAP + m]) <= DGC)
        for p in np.nonzero(v)[0]:
            p = int(p); t = s[p:p + W]; e = cand.get(t)
            if e is None:
                g1, g2 = gc[p], gc[p + ARM + GAP]
                cand[t] = e = [abs(g1 - .5) + abs(g2 - .5) + .5 * abs(g1 - g2) + pen[p] + pen[p + ARM + GAP], {}, 0]
            if ii not in e[1]:
                e[1][ii] = p; e[2] |= parts_of(p, p + W, iso['cds'])
    n = len(isos); occ = [np.zeros(len(i['seq']) + 4, bool) for i in isos]
    order = sorted(cand.items(), key=lambda kv: kv[1][0] - 0.05 * len(kv[1][1]) / n)
    out = []; per = {}
    for want_cds, cap in ((True, CAP_CDS), (False, CAP_UTR)):      # CDS pairs first, then pairs touching a UTR
        got = 0
        for t, (sc, pos, parts) in order:
            if (parts == 2) != want_cds: continue
            if got >= cap: break
            cov = frozenset(pos); key = (cov, parts)
            if per.get(key, 0) >= (MAX_ALL if len(cov) == n else MAX_SUB): continue
            if any(occ[i][max(0, p - 2):p + W + 2].any() for i, p in pos.items()): continue
            for i, p in pos.items(): occ[i][max(0, p - 2):p + W + 2] = True
            per[key] = per.get(key, 0) + 1; got += 1
            i0 = min(pos)
            out.append([pack(t), max(0, round(100 * (1 - sc / 1.2))), format(sum(1 << i for i in cov), 'x'), i0, pos[i0], parts])
    # stretches unique to one isoform: 25-mers that occur in exactly one isoform of the gene
    uniq = []
    if n > 1:
        codes = [np.unique(kcodes(enc(i['seq']).astype(np.uint64) & np.uint64(3), ARM)) for i in isos]
        allc = np.concatenate(codes); vals, cnt = np.unique(allc, return_counts=True); single = vals[cnt == 1]
        for ii, iso in enumerate(isos):
            c = kcodes(enc(iso['seq']).astype(np.uint64) & np.uint64(3), ARM)
            j = np.searchsorted(single, c); j[j == len(single)] = 0
            u = (single[j] == c) if len(single) else np.zeros(len(c), bool)
            iv = []; d = np.diff(np.concatenate(([0], u.astype(np.int8), [0])))
            for a, e in zip(np.nonzero(d == 1)[0], np.nonzero(d == -1)[0]):
                a, e = int(a), int(e) - 1 + ARM
                iv.append([a, e, parts_of(a, e, iso['cds'])])
            uniq.append(iv)
    prot = re.sub(r' isoform \S+$', '', isos[0]['prot'])
    iso = [[i['pid'], i['iso'], len(i['seq']), i['same'], int(i['partial']), i['tid'], i['cds'][0], i['cds'][1], int(i['utr']), i['var']]
           for i in isos]
    return dict(s=g['sym'], id=g['gid'], n=g['desc'] or prot, iso=iso, p=out, u=uniq)

os.makedirs(f'{sp}_out', exist_ok=True); t0 = time.time(); last = start
path = f'{sp}_out/{start:06d}.jsonl'
done = {json.loads(l)['k'] for l in open(path)} if os.path.exists(path) else set()
with open(path, 'a') as f:
    for k in range(start, min(stop, len(G))):
        if k in done: continue
        d = design(G[k]); d['k'] = k; f.write(json.dumps(d, separators=(',', ':')) + '\n'); f.flush(); last = k
        if time.time() - t0 > TL: break
print(sp, 'genes', start, '..', last, 'of', len(G), f'{time.time() - t0:.0f}s')
