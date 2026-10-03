import numpy as np, sys, time
sp=sys.argv[1]; t=time.time(); out=[]
for b in range(64):
    c=np.fromfile(f'{sp}_bk/c{b:02d}',np.uint64)
    c.sort()
    dup=c[1:][c[1:]==c[:-1]]           # an 18-mer seen in two or more genes (each gene contributes it at most once)
    out.append(np.unique(dup))
S=np.concatenate(out); S.sort(); np.save(f'{sp}_shared18.npy',S)
print(sp,'18-mers shared between genes:',len(S),f'{time.time()-t:.0f}s')
