"""Collect the database cross-references (CCDS, Ensembl protein IDs) from the CDS headers of an NCBI datasets zip.

Usage: python protein_xrefs.py <sp> <dataset.zip>     writes <sp>_pxref.json  {protein_id: [CCDS..., ENSP...]}
"""
import sys, re, zipfile, io, json
sp, zp = sys.argv[1], sys.argv[2]
z = zipfile.ZipFile(zp)
m = next(n for n in z.namelist() if re.search(r'data/GCF_[^/]+/cds_from_genomic\.fna$', n))
out = {}
with z.open(m) as fh:
    for line in io.TextIOWrapper(fh, encoding='utf-8'):
        if not line.startswith('>'): continue
        pid = re.search(r'\[protein_id=([^\]]+)\]', line); dx = re.search(r'\[db_xref=([^\]]+)\]', line)
        if not pid or not dx: continue
        ids = [x.split(':', 1)[1].split('.')[0] for x in dx.group(1).split(',') if x.startswith(('CCDS:', 'Ensembl:'))]
        if ids: out.setdefault(pid.group(1), sorted(set(ids)))
json.dump(out, open(f'{sp}_pxref.json', 'w'))
print(sp, len(out), 'proteins with CCDS or Ensembl IDs')
