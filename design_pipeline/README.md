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
