import numpy as np, pickle, os, time, json
from kmers import enc, kcodes, LUT
K=25; t0=time.time(); ST='g25_state.json'
if not os.path.exists('bany_q25.npy'):
    G=pickle.load(open('bany_genes.pkl','rb'))
    Q=np.unique(np.concatenate([kcodes(enc(i['seq']),K) for g in G for i in g['iso']]))
    np.save('bany_q25.npy',Q); np.save('bany_c25.npy',np.zeros(len(Q),np.int32)); json.dump({'done':[]},open(ST,'w'))
    print('query 25-mers:',len(Q),f'{time.time()-t0:.0f}s')
Q=np.load('bany_q25.npy'); C=np.load('bany_c25.npy'); state=json.load(open(ST)); done=set(state['done'])
HITS=[]
def count_chunk(s):
    b=LUT[np.frombuffer(s,np.uint8)]
    bad=(b==255); b=np.where(bad,0,b).astype(np.uint64)
    n=len(b)-K+1
    if n<=0: return
    badw=np.convolve(bad.astype(np.int32),np.ones(K,np.int32),'valid')>0
    for arr in (b, (np.uint64(3)-b)[::-1]):          # forward strand, then reverse complement
        c=kcodes(arr,K); ok=~(badw if arr is b else badw[::-1])
        c=np.sort(c[ok]); idx=np.searchsorted(Q,c); idx[idx==len(Q)]=0; hit=Q[idx]==c
        HITS.append(idx[hit].astype(np.int64))
def chroms():
    name=None; buf=[]
    with open('bany_genome.fa','rb') as f:
        for line in f:
            if line.startswith(b'>'):
                if name: yield name,b''.join(buf)
                name=line[1:].split()[0].decode(); buf=[]
            else: buf.append(line.strip().upper())
    if name: yield name,b''.join(buf)
for name,s in chroms():
    if name in done: continue
    CH=4_000_000; HITS.clear()
    for st in range(0,len(s),CH): count_chunk(s[st:st+CH+K-1])
    if HITS: C+=np.bincount(np.concatenate(HITS),minlength=len(Q)).astype(np.int32)
    done.add(name); state['done']=sorted(done)
    np.save('bany_c25.npy',C); json.dump(state,open(ST,'w'))
    if time.time()-t0>230: break
print('chromosomes done:',len(done),'of 82;',f'{time.time()-t0:.0f}s')
