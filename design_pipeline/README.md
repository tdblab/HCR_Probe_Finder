# Transcriptome-wide probe design

These scripts produced the files in `tx/`. Inputs (not stored in the repository because of their size):
`bany_cds.fa` and `bany_genome.fa` (NCBI ilBicAnyn1.1, GCF_947172395.1) and `danio_cds.fa` (the zebrafish CDS).

Run from a working folder containing those files (each step resumes if interrupted):

1. `python prep.py bany` and `python prep.py danio`: genes and distinct isoform CDS (identical isoforms collapsed).
2. `python kmers.py <sp>` then `python shared.py <sp>`: 18-mers shared between genes.
3. `python genome25.py` (B. anynana only): genome copy number of every CDS 25-mer; rerun until all chromosomes are done.
4. `python design.py <sp> 0 <number of genes>`: rerun until every gene is written to `<sp>_out/`.
5. `python check.py <sp>`: verifies positions and isoform coverage of a random sample of pairs.
6. `python export_tx.py`: writes `tx/<sp>/index.js` and the data chunks (edit the paths at the top first).

Design rules are set at the top of `design.py` (arm length, GC limits, maximum pairs kept per gene).
`labcheck.py` reports how many of the lab's own probe pairs would pass the same filters.

## mRNA-based species (killifish, fly, worm, mouse, human)

Input: the NCBI datasets zip for each species (`rna.fna`, `genomic.gtf`, `cds_from_genomic.fna`, RefSeq annotation).
Species codes: `nfur`, `dmel`, `cele`, `mmus`, `hsap`.

1. `python prep_mrna.py <sp> <dataset.zip>`: links every mRNA to its protein through the GTF and finds the CDS inside the
   mRNA, which gives the 5′UTR / CDS / 3′UTR boundaries. Isoforms whose CDS is not found exactly inside the mRNA are kept
   as CDS-only. Readthrough genes are left out. Also saves every RefSeq transcript (non-coding ones included) by gene.
2. `python shared_bg.py <sp>`: 18-mers found in transcripts of two or more genes.
3. `python design_mrna.py <sp> 0 <number of genes>`: rerun until every gene is written to `<sp>_out/`. Each pair is
   labelled with the transcript parts it touches (1 = 5′UTR, 2 = CDS, 4 = 3′UTR). CDS pairs and pairs touching a UTR
   have separate quotas (100 each per gene).
4. `python check_mrna.py <sp>`: verifies position, isoform coverage, part label and gene specificity of a sample of pairs.
5. `python protein_xrefs.py <sp> <dataset.zip>` (optional): CCDS and Ensembl protein IDs from the CDS headers, added to search.
6. `python export_mrna.py <repo>/tx <sp> ...`: writes `tx/<sp>/index.js` and the data chunks.
