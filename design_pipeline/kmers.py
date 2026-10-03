import numpy as np, pickle, sys, os, time
LUT=np.full(256,255,np.uint8)
for i,c in enumerate(b'ACGT'): LUT[c]=i
def enc(s): return LUT[np.frombuffer(s.encode(),np.uint8)].astype(np.uint64)
def kcodes(b,k):
    n=len(b)-k+1
    if n<=0: return np.zeros(0,np.uint64)
    c=np.zeros(n,np.uint64)
    for i in range(k): c=(c<<np.uint64(2))|b[i:i+n]
    return c
if __name__=='__main__':
    sp=sys.argv[1]; K=18; NB=64; SH=np.uint64(2*K-6)
    G=pickle.load(open(f'{sp}_genes.pkl','rb'))
    os.makedirs(f'{sp}_bk',exist_ok=True)
    for f in os.listdir(f'{sp}_bk'): os.remove(f'{sp}_bk/{f}')
    bufc=[[] for _ in range(NB)]; bufg=[[] for _ in range(NB)]; nbuf=0; t=time.time(); total=0
    def flush():
        for b in range(NB):
            if bufc[b]:
                np.concatenate(bufc[b]).tofile(open(f'{sp}_bk/c{b:02d}','ab')); np.concatenate(bufg[b]).tofile(open(f'{sp}_bk/g{b:02d}','ab'))
                bufc[b].clear(); bufg[b].clear()
    for gi,g in enumerate(G):
        u=np.unique(np.concatenate([kcodes(enc(i['seq']),K) for i in g['iso']]))
        total+=len(u)
        bk=(u>>SH).astype(np.int64); order=np.argsort(bk,kind='stable'); u=u[order]; bk=bk[order]
        cuts=np.searchsorted(bk,np.arange(NB+1))
        for b in range(NB):
            if cuts[b+1]>cuts[b]:
                bufc[b].append(u[cuts[b]:cuts[b+1]]); bufg[b].append(np.full(cuts[b+1]-cuts[b],gi,np.int32))
        nbuf+=len(u)
        if nbuf>20_000_000: flush(); nbuf=0
    flush()
    print(sp,'genes',len(G),'gene-distinct 18-mers',total,f'{time.time()-t:.0f}s')
