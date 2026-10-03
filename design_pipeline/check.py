import json,pickle,base64,random,sys,statistics as st
sp=sys.argv[1]
G=pickle.load(open(f'{sp}_genes.pkl','rb'))
def unpack(b):
    v=int.from_bytes(base64.b64decode(b),'big'); return ''.join('ACGT'[(v>>(2*(51-i)))&3] for i in range(52))
def rc(s): return s[::-1].translate(str.maketrans('ACGT','TGCA'))
import glob
D={}
for f in sorted(glob.glob(f'{sp}_out/*.jsonl')):
    for l in open(f): d=json.loads(l); D[d['k']]=d
print(sp,'designed genes',len(D),'of',len(G))
bad=0; n=0; random.seed(1)
for k in random.sample(sorted(D),300):
    d=D[k]; g=G[k]
    for t,sc,mask,i0,p in d['p']:
        t=unpack(t); n+=1
        m=int(mask,16)
        for i,iso in enumerate(g['iso']):
            inside=t in iso['seq']
            if bool(m>>i&1)!=inside: bad+=1
        if g['iso'][i0]['seq'][p:p+52]!=t: bad+=1
print('pairs checked',n,'| coverage/position errors',bad)
pp=[len(d['p']) for d in D.values()]
print('pairs per gene: median',st.median(pp),'| 0 pairs:',sum(1 for x in pp if x==0),'| >=10:',sum(1 for x in pp if x>=10),'| >=20:',sum(1 for x in pp if x>=20))
multi=[d for d in D.values() if len(d['iso'])>1]
print('multi-isoform genes',len(multi),'| with >=1 isoform-unique region:',sum(1 for d in multi if any(d['u'])))
