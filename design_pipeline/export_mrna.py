"""Write tx/<sp>/index.js and data chunks for species designed with design_mrna.py.

Usage: python export_mrna.py <out tx folder> <sp> [<sp> ...]
Index row: [symbol, GeneID, name, chunk, row in chunk, isoforms, CDS pairs, lab short name, lab long name, lab sets,
            search text (synonyms, database IDs, transcript and protein IDs), synonyms shown on the page, pairs touching a UTR]
"""
import json, os, sys, glob, pickle, re, datetime

META = {
    'nfur': ('Nothobranchius furzeri (killifish)', 'NfurGRZ-RIMD1 (GCF_043380555.1), NCBI RefSeq'),
    'dmel': ('Drosophila melanogaster', 'Release 6 plus ISO1 MT (GCF_000001215.4), NCBI RefSeq / FlyBase 6.54'),
    'cele': ('Caenorhabditis elegans', 'WBcel235 (GCF_000002985.6), NCBI RefSeq / WormBase WS298'),
    'mmus': ('Mus musculus', 'GRCm39 (GCF_000001635.27), NCBI RefSeq'),
    'hsap': ('Homo sapiens', 'GRCh38.p14 (GCF_000001405.40), NCBI RefSeq, primary assembly'),
}
RULES = ('Designed on the full mRNA of every protein-coding transcript; positions are on the mRNA. Two 25-nt arms with a 2-nt gap '
         '(HCR v3.0 split initiator). Hard filters: arm GC 30–80%, arms within 30 points of each other, no run of 6 identical bases, '
         'no GGGGG/CCCCC, no 18-mer shared with any transcript of another gene (UTRs and non-coding RNAs included; readthrough genes '
         'left out). Ranked by closeness of each arm to 50% GC, similar GC between arms, and penalties for GGGG/CCCC or runs of 5. '
         'Pairs do not overlap within any isoform. CDS pairs and pairs touching a UTR have separate quotas, so long UTRs do not '
         'crowd out CDS pairs. No genome check was possible for this species.')
out_root = sys.argv[1]
for sp in sys.argv[2:]:
    G = pickle.load(open(f'{sp}_genes.pkl', 'rb'))
    PX = json.load(open(f'{sp}_pxref.json')) if os.path.exists(f'{sp}_pxref.json') else {}
    D = [json.loads(l) for f in sorted(glob.glob(f'{sp}_out/*.jsonl')) for l in open(f)]
    D.sort(key=lambda d: d['k']); assert len(D) == len(G), (sp, len(D), len(G))
    od = f'{out_root}/{sp}'; os.makedirs(od, exist_ok=True)
    for f in glob.glob(f'{od}/*.js'): os.remove(f)
    SH = 100; idx = []
    for si in range(0, len(D), SH):
        chunk = D[si:si + SH]
        for d in chunk: d.pop('k', None)
        open(f'{od}/s{si // SH:04d}.js', 'w').write(f'HCRTX_SHARD("{sp}",{si // SH},' + json.dumps(chunk, separators=(',', ':')) + ');\n')
        for j, d in enumerate(chunk):
            g = G[si + j]
            syn = [x for x in g['syn'] if '\\' not in x]
            xr = [x.split(':', 1)[1] for x in g['xref']]          # FLYBASE:FBgn... -> FBgn...; MGI:MGI:97490 -> MGI:97490
            loc = g['locus']; locs = [loc, loc.split('_', 1)[1]] if '_' in loc else [loc]
            ids = [i[5].split('.')[0] for i in d['iso']] + [i[0].split('.')[0] for i in d['iso']]
            px = sorted({x for i in d['iso'] for x in PX.get(i[0], [])})       # CCDS and Ensembl protein IDs
            search = ' '.join(syn[:40] + xr + locs + ids + px)
            idx.append([d['s'], d['id'], d['n'], si // SH, j, len(d['iso']), sum(1 for p in d['p'] if p[5] == 2),
                        '', '', 0, search, ', '.join(syn[:8]), sum(1 for p in d['p'] if p[5] != 2)])
    name, asm = META[sp]
    meta = dict(name=name, assembly=asm, genome=False, utr=True, rules=RULES, built=datetime.date.today().isoformat(),
                genes=len(D), pairs=sum(len(d['p']) for d in D), cds_pairs=sum(1 for d in D for p in d['p'] if p[5] == 2))
    open(f'{od}/index.js', 'w').write(f'HCRTX_INDEX("{sp}",' + json.dumps(dict(meta=meta, g=idx), separators=(',', ':')) + ');\n')
    sz = sum(os.path.getsize(f) for f in glob.glob(f'{od}/*.js'))
    print(sp, len(D), 'genes', meta['pairs'], 'pairs (', meta['cds_pairs'], 'CDS )', f'{sz / 1e6:.1f} MB',
          '| index', f"{os.path.getsize(f'{od}/index.js') / 1e6:.2f} MB", '| genes with no pairs', sum(1 for d in D if not d['p']))
