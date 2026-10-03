import json, numpy as np, re, collections, pickle
from kmers import enc, kcodes
S18=np.load('bany_shared18.npy'); Q25=np.load('bany_q25.npy'); C25=np.load('bany_c25.npy')
def rc(s): return s[::-1].translate(str.maketrans('ACGT','TGCA'))
def why(site):
    r=[]; gc=sum(c in 'GC' for c in site)/25
    if not .30<=gc<=.80: r.append('GC')
    if re.search(r'(A{6}|C{6}|G{6}|T{6})',site): r.append('run of 6')
    if 'GGGGG' in site or 'CCCCC' in site: r.append('GGGGG/CCCCC')
    c=kcodes(enc(site),18); i=np.searchsorted(S18,c); i[i==len(S18)]=0
    if (S18[i]==c).any(): r.append('18-mer shared with another gene')
    c=kcodes(enc(site),25)[0]; j=np.searchsorted(Q25,c)
    if j<len(Q25) and Q25[j]==c and C25[j]>1: r.append('>1 genome copy')
    return r,gc
cnt=collections.Counter(); n=0; dgc=0
for f in ('../hcr/tables.json','../hcr/jeriel_tables.json','../hcr/suriya_tables.json'):
    for p in json.load(open(f))['PAIRS']:
        if p.get('gmatch')!='Both halves' or len(p['a1'])!=25 or len(p['a2'])!=25: continue
        n+=1; s1,s2=rc(p['a1']),rc(p['a2']); r1,g1=why(s1); r2,g2=why(s2)
        rs=set(r1)|set(r2)
        if abs(g1-g2)>0.30: rs.add('arm GC difference >30 points')
        cnt['passes all' if not rs else 'fails']+=1
        for r in rs: cnt[r]+=1
print('lab pairs that match the B. anynana genome on both halves:',n)
for k,v in cnt.most_common(): print(f'  {k:34} {v:5}  {v/n:.0%}')
