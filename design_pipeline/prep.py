import re,sys,pickle,collections
def parse(path):
    genes=collections.OrderedDict(); h=None; buf=[]
    def flush():
        if h is None: return
        seq=''.join(buf).upper()
        g=re.search(r'\[gene=([^\]]+)\]',h); gid=re.search(r'GeneID:(\d+)',h)
        pid=re.search(r'\[protein_id=([^\]]+)\]',h); pr=re.search(r'\[protein=([^\]]+)\]',h)
        key=g.group(1) if g else h[1:30]
        e=genes.setdefault(key,dict(sym=key,gid=gid.group(1) if gid else '',iso=[]))
        e['iso'].append(dict(pid=pid.group(1) if pid else '',prot=pr.group(1) if pr else '',seq=seq,
                             partial='partial=' in h,pseudo='pseudo=true' in h))
    with open(path) as f:
        for line in f:
            if line.startswith('>'):
                flush(); h=line.strip(); buf=[]
            else: buf.append(line.strip())
    flush()
    # collapse isoforms with identical CDS (keep the first; remember the others as synonyms)
    for g in genes.values():
        seen={}; keep=[]
        for i in g['iso']:
            if i['seq'] in seen: seen[i['seq']]['same'].append(i['pid'])
            else: i['same']=[]; seen[i['seq']]=i; keep.append(i)
        g['iso']=keep
    return list(genes.values())
sp=sys.argv[1]; G=parse(f'{sp}_cds.fa')
pickle.dump(G,open(f'{sp}_genes.pkl','wb'))
niso=sum(len(g['iso']) for g in G); tot=sum(len(i['seq']) for g in G for i in g['iso'])
multi=sum(1 for g in G if len(g['iso'])>1)
print(sp,len(G),'genes',niso,'distinct isoform CDS',f'{tot/1e6:.1f} Mb','| genes with >1 distinct isoform:',multi,
      '| non-ACGT bases:',sum(1 for g in G for i in g['iso'] if re.search('[^ACGT]',i['seq'])))
