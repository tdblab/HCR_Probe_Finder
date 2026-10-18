"""Build gene / transcript tables from an NCBI datasets zip (rna.fna + genomic.gtf + cds_from_genomic.fna).

Usage: python prep_mrna.py <species_code> <ncbi_dataset.zip>

Writes <sp>_genes.pkl   protein-coding genes; each isoform is a full mRNA with the CDS located inside it
       <sp>_bg.pkl      every RefSeq transcript (coding and non-coding) grouped by gene, used for specificity
Isoforms whose CDS cannot be found exactly inside their mRNA are kept as CDS-only (cds=(0,len), utr=False).
"""
import sys, re, zipfile, pickle, collections, io, time

sp, zpath = sys.argv[1], sys.argv[2]
z = zipfile.ZipFile(zpath)
names = z.namelist()
acc = sorted({m.group(1) for n in names for m in [re.search(r'data/(GCF_[^/]+)/', n)] if m})[0]   # RefSeq copy
def member(suffix): return f'ncbi_dataset/data/{acc}/{suffix}'
t0 = time.time()

def attrs(s):
    d = collections.defaultdict(list)
    for k, v in re.findall(r'(\w+) "([^"]*)"', s):
        d[k].append(v)
    return d

# ---------- GTF: genes, transcript -> gene, transcript <-> protein ----------
genes = {}              # GeneID -> info
gene_of_gid = {}        # GTF gene_id -> GeneID
tx_gene = {}            # transcript_id -> GeneID
tx_prot = {}            # transcript_id -> protein_id
tx_seqid = {}
with z.open(member('genomic.gtf')) as fh:
    for raw in io.TextIOWrapper(fh, encoding='utf-8'):
        if raw.startswith('#'): continue
        f = raw.rstrip('\n').split('\t')
        if len(f) < 9: continue
        typ = f[2]
        if typ not in ('gene', 'transcript') and 'gbkey "CDS"' not in f[8]: continue   # WormBase GTFs name CDS lines by CDS id
        a = attrs(f[8])
        if typ == 'gene':
            gid = next((x.split(':', 1)[1] for x in a['db_xref'] if x.startswith('GeneID:')), '')
            gene_of_gid[a['gene_id'][0]] = gid
            if gid in genes and genes[gid]['seqid'].startswith('NC_'): continue
            xr = [x for x in a['db_xref'] if not x.startswith('GeneID:')]
            genes[gid] = dict(gid=gid, sym=(a.get('gene') or a['gene_id'])[0], desc=(a.get('description') or [''])[0],
                              syn=a.get('gene_synonym', []), xref=xr, biotype=(a.get('gene_biotype') or [''])[0],
                              locus=(a.get('locus_tag') or [''])[0], seqid=f[0])
        elif typ == 'transcript':
            tid = a['transcript_id'][0]
            if tid in tx_gene and tx_seqid[tid].startswith('NC_'): continue
            tx_gene[tid] = gene_of_gid.get(a['gene_id'][0], a['gene_id'][0]); tx_seqid[tid] = f[0]
        else:
            tid = a['transcript_id'][0]; pid = (a.get('protein_id') or [''])[0]
            if pid and tid not in tx_prot: tx_prot[tid] = pid
print(sp, acc, 'GTF:', len(genes), 'genes', len(tx_gene), 'transcripts', len(tx_prot), 'coding', f'{time.time()-t0:.0f}s', flush=True)

def fasta(name):
    with z.open(member(name)) as fh:
        h = None; buf = []
        for line in io.TextIOWrapper(fh, encoding='utf-8'):
            if line.startswith('>'):
                if h: yield h, ''.join(buf).upper()
                h = line[1:].strip(); buf = []
            else: buf.append(line.strip())
        if h: yield h, ''.join(buf).upper()

# ---------- CDS: protein_id -> CDS sequence and protein name ----------
cds = {}
for h, s in fasta('cds_from_genomic.fna'):
    pid = re.search(r'\[protein_id=([^\]]+)\]', h)
    if not pid: continue
    pid = pid.group(1)
    if pid in cds and cds[pid]['primary']: continue
    pr = re.search(r'\[protein=([^\]]+)\]', h)
    cds[pid] = dict(seq=s, prot=pr.group(1) if pr else '', partial='partial=' in h,
                    exc='[exception=' in h, primary=h.startswith('lcl|NC_'))
print(sp, 'CDS:', len(cds), 'proteins', f'{time.time()-t0:.0f}s', flush=True)

# ---------- mRNAs ----------
bg = collections.defaultdict(list)       # GeneID -> [seq, ...] (all RefSeq transcripts)
iso = collections.defaultdict(list)
stat = collections.Counter()
for h, s in fasta('rna.fna'):
    tid = h.split()[0]
    g = tx_gene.get(tid)
    if g is None:
        m = re.search(r'\(([^()]+)\),[^,]*$', h); g = 'sym:' + (m.group(1) if m else tid); stat['transcript not in GTF'] += 1
    bg[g].append(s)
    pid = tx_prot.get(tid)
    if not pid or pid not in cds: continue
    c = cds[pid]
    p = s.find(c['seq'])
    var = re.search(r'transcript variant (\S+?)[,)]', h)
    il = re.search(r'isoform (\S+)$', c['prot'])
    rec = dict(tid=tid, pid=pid, prot=c['prot'], partial=c['partial'], var=var.group(1) if var else '',
               iso=il.group(1) if il else '')
    if p >= 0:
        rec.update(seq=s, cds=(p, p + len(c['seq'])), utr=True); stat['CDS located in mRNA'] += 1
    else:
        rec.update(seq=c['seq'], cds=(0, len(c['seq'])), utr=False); stat['CDS not found in mRNA (CDS only)'] += 1
    iso[g].append(rec)

# readthrough genes join the exons of two neighbouring genes: drop them so they do not make the parent genes look shared
rt = {gid for gid, g in genes.items() if 'readthrough' in g['desc'].lower()}
for gid in rt: bg.pop(gid, None); iso.pop(gid, None)
stat['readthrough genes left out'] = len(rt)
out = []
for gid, isos in iso.items():
    info = genes.get(gid, dict(gid=gid, sym=gid, desc='', syn=[], xref=[], biotype='', locus=''))
    seen = {}; keep = []
    for r in isos:
        if r['seq'] in seen: seen[r['seq']]['same'].append(r['tid'])
        else: r['same'] = []; seen[r['seq']] = r; keep.append(r)
    keep.sort(key=lambda r: (r['iso'] == '', r['iso'], r['tid']))
    out.append(dict(info, iso=keep))
out.sort(key=lambda g: g['sym'].lower())
pickle.dump(out, open(f'{sp}_genes.pkl', 'wb'))
pickle.dump(dict(bg), open(f'{sp}_bg.pkl', 'wb'))
n = sum(len(g['iso']) for g in out); mb = sum(len(i['seq']) for g in out for i in g['iso']) / 1e6
print(sp, 'genes', len(out), '| distinct coding isoforms', n, f'| {mb:.0f} Mb', '| background transcripts',
      sum(len(v) for v in bg.values()), dict(stat), f'{time.time()-t0:.0f}s')
