import numpy as np, pickle, json, sys, os, time, re, base64
from kmers import enc, kcodes
sp=sys.argv[1]; start=int(sys.argv[2]); stop=int(sys.argv[3]); TL=float(sys.argv[4]) if len(sys.argv)>4 else 240
G=pickle.load(open(f'{sp}_genes.pkl','rb')); S18=np.load(f'{sp}_shared18.npy')
GEN = sp=='bany'
if GEN: Q25=np.load('bany_q25.npy'); C25=np.load('bany_c25.npy')
ARM,GAP=25,2; W=2*ARM+GAP
GCMIN,GCMAX,DGC=0.30,0.80,0.30
MAX_ALL,MAX_SUB=60,15
def winsum(x,k):                       # sum over each length-k window
    c=np.concatenate(([0],np.cumsum(x))); return c[k:]-c[:-k]
def site_ok(s,b):
    L=len(s); n=L-ARM+1
    if n<=0: return None,None
    gc=winsum(((b==1)|(b==2)).astype(np.int32),ARM)/ARM
    eq=(b[1:]==b[:-1]).astype(np.int32)
    run5=np.zeros(L,np.int32); run5[:L-4]=winsum(eq,4)==4          # 5 identical bases start here
    run6=np.zeros(L,np.int32); run6[:L-5]=winsum(eq,5)==5          # 6 identical bases start here
    g4=np.zeros(L,np.int32); g5=np.zeros(L,np.int32)
    for m in re.finditer('(?=GGGG|CCCC)',s): g4[m.start()]=1
    for m in re.finditer('(?=GGGGG|CCCCC)',s): g5[m.start()]=1
    ok=(gc>=GCMIN)&(gc<=GCMAX)&(winsum(run6,ARM-5)[:n]==0)&(winsum(g5,ARM-4)[:n]==0)
    pen=0.1*(winsum(run5,ARM-4)[:n]>0)+0.1*(winsum(g4,ARM-3)[:n]>0)          # soft penalties used in ranking
    c18=kcodes(b,18); i=np.searchsorted(S18,c18); i[i==len(S18)]=0; bad18=(S18[i]==c18).astype(np.int32)
    ok&=winsum(bad18,ARM-18+1)[:n]==0                                 # no 18-mer shared with another gene
    if GEN:
        c25=kcodes(b,ARM); j=np.searchsorted(Q25,c25); ok&=C25[j]<=1  # at most one copy in the genome
    return ok,(gc,pen)
def pack(t):
    v=0
    for ch in t: v=(v<<2)|'ACGT'.index(ch)
    return base64.b64encode(v.to_bytes(13,'big')).decode()
def design(g):
    isos=g['iso']; cand={}
    for ii,iso in enumerate(isos):
        s=iso['seq']; b=enc(s); ok,gp=site_ok(s,b)
        if ok is None or len(s)<W: continue
        gc,pen=gp
        m=len(s)-W+1
        v=ok[:m]&ok[ARM+GAP:ARM+GAP+m]&(np.abs(gc[:m]-gc[ARM+GAP:ARM+GAP+m])<=DGC)
        for p in np.nonzero(v)[0]:
            t=s[p:p+W]; e=cand.get(t)
            if e is None:
                g1,g2=gc[p],gc[p+ARM+GAP]
                cand[t]=e=[abs(g1-.5)+abs(g2-.5)+.5*abs(g1-g2)+pen[p]+pen[p+ARM+GAP],{}]
            e[1].setdefault(ii,int(p))
    n=len(isos); occ=[np.zeros(len(i['seq'])+4,bool) for i in isos]
    order=sorted(cand.items(),key=lambda kv: kv[1][0]-0.05*len(kv[1][1])/n)
    out=[]; per={}
    for t,(sc,pos) in order:
        cov=frozenset(pos); full=len(cov)==n
        lim=MAX_ALL if full else MAX_SUB
        if per.get(cov,0)>=lim: continue
        if any(occ[i][max(0,p-2):p+W+2].any() for i,p in pos.items()): continue
        for i,p in pos.items(): occ[i][max(0,p-2):p+W+2]=True
        per[cov]=per.get(cov,0)+1
        mask=sum(1<<i for i in cov); i0=min(pos)
        out.append([pack(t),max(0,round(100*(1-sc/1.2))),format(mask,'x'),i0,pos[i0]])
    # regions unique to each isoform (25-mers absent from every other isoform of the gene)
    uniq=[]
    if n>1:
        codes=[kcodes(enc(i['seq']),ARM) for i in isos]
        for ii in range(n):
            others=np.unique(np.concatenate([c for k,c in enumerate(codes) if k!=ii])) if n>1 else np.zeros(0,np.uint64)
            c=codes[ii]; j=np.searchsorted(others,c); j[j==len(others)]=0
            u=others[j]!=c if len(others) else np.ones(len(c),bool)
            iv=[]; st=None
            for p,f in enumerate(u):
                if f and st is None: st=p
                if not f and st is not None: iv.append([st,p-1+ARM]); st=None
            if st is not None: iv.append([st,len(c)-1+ARM])
            uniq.append(iv)
    prot=re.sub(r' isoform \S+$','',isos[0]['prot'])
    iso=[[i['pid'],(re.search(r'isoform (\S+)$',i['prot']) or [None,''])[1],len(i['seq']),i['same'],int(i['partial'])] for i in isos]
    return dict(s=g['sym'],id=g['gid'],n=prot,iso=iso,p=out,u=uniq)
os.makedirs(f'{sp}_out',exist_ok=True); t0=time.time(); last=start
with open(f'{sp}_out/{start:06d}.jsonl','a') as f:
    done=set()
    if os.path.exists(f'{sp}_out/{start:06d}.jsonl'):
        done={json.loads(l)['k'] for l in open(f'{sp}_out/{start:06d}.jsonl')}
    for k in range(start,min(stop,len(G))):
        if k in done: continue
        d=design(G[k]); d['k']=k; f.write(json.dumps(d,separators=(',',':'))+'\n'); last=k
        if time.time()-t0>TL: break
print(sp,'genes',start,'..',last,f'{time.time()-t0:.0f}s')
