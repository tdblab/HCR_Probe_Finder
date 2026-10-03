import json, os, re, sys, collections, glob, datetime
sys.path.insert(0,'/home/claude/hcr'); import gene_names as GN
OUT='/home/claude/hcr/HCR_probe_database/tx'
# lab names per LOC / GeneID from all designers' sets
lab=collections.defaultdict(lambda: dict(short=set(),long=set(),sets=0))
for f in ('../hcr/tables.json','../hcr/jeriel_tables.json','../hcr/suriya_tables.json'):
    for s in json.load(open(f))['SETS']:
        keys=set(re.findall(r'LOC\d+',s.get('gloc','') or '') or re.findall(r'LOC\d+',s.get('loc','') or ''))
        keys|={'GeneID:'+m for m in re.findall(r'GeneID:(\d+)',s.get('loc','') or '')}
        if s.get('gconf')=='Low' and s.get('locsrc')=='genome_cds.fa match': continue
        g=GN.G.get(s['gene'])
        for k in keys:
            e=lab[k]; e['sets']+=1; e['short'].add(g[0] if g else s['gene'])
            if g and g[1]: e['long'].add(g[1])
META={'bany':dict(name='Bicyclus anynana',assembly='ilBicAnyn1.1 (GCF_947172395.1)',genome=True),
      'danio':dict(name='Danio rerio',assembly='GRCz13ab, as provided (Danio_CDS.fna)',genome=False)}
RULES=('Two 25-nt arms with a 2-nt gap (HCR v3.0 split initiator). Hard filters: arm GC 30–80%, arms within 30 points of each other, '
       'no run of 6 identical bases, no GGGGG/CCCCC, no 18-mer shared with another gene\'s CDS{g}. Ranked by closeness of each arm to 50% GC, '
       'similar GC between arms, and penalties for GGGG/CCCC or runs of 5. Pairs do not overlap within any isoform.')
for sp in ('bany','danio'):
    D=[json.loads(l) for f in sorted(glob.glob(f'{sp}_out/*.jsonl')) for l in open(f)]
    D.sort(key=lambda d:d['k'])
    os.makedirs(f'{OUT}/{sp}',exist_ok=True)
    for f in glob.glob(f'{OUT}/{sp}/*.js'): os.remove(f)
    SH=100; idx=[]
    for si in range(0,len(D),SH):
        chunk=D[si:si+SH]
        for d in chunk: d.pop('k',None)
        open(f'{OUT}/{sp}/s{si//SH:04d}.js','w').write(f'HCRTX_SHARD("{sp}",{si//SH},'+json.dumps(chunk,separators=(',',':'))+');\n')
        for j,d in enumerate(chunk):
            key=d['s'] if d['s'].startswith('LOC') else f"GeneID:{d['id']}"
            L=lab.get(d['s']) or lab.get(f"GeneID:{d['id']}")
            idx.append([d['s'],d['id'],d['n'],si//SH,j,len(d['iso']),len(d['p']),
                        ', '.join(sorted(L['short'])) if L else '', '; '.join(sorted(L['long'])) if L else '', L['sets'] if L else 0])
    meta=dict(META[sp],rules=RULES.format(g=', and at most one copy of each arm site in the genome' if META[sp]['genome'] else ''),
              built=datetime.date.today().isoformat(),genes=len(D),pairs=sum(len(d['p']) for d in D))
    open(f'{OUT}/{sp}/index.js','w').write(f'HCRTX_INDEX("{sp}",'+json.dumps(dict(meta=meta,g=idx),separators=(',',':'))+');\n')
    sz=sum(os.path.getsize(f) for f in glob.glob(f'{OUT}/{sp}/*.js'))
    print(sp,len(D),'genes',meta['pairs'],'pairs',len(glob.glob(f'{OUT}/{sp}/s*.js')),'shards',f'{sz/1e6:.1f} MB','| index',f"{os.path.getsize(f'{OUT}/{sp}/index.js')/1e6:.2f} MB",
          '| genes with lab names',sum(1 for x in idx if x[7]))
