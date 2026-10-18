# HCR probe database

All HCR probe sets designed in the lab, one Excel file per person, plus a searchable viewer
(`index.html`) built from them.

<img width="1262" height="1244" alt="hcr_probe_finder" src="https://github.com/user-attachments/assets/a128d790-df3c-4edc-836c-a41bab7a59df" />


```
data/                       one workbook per designer (Dragonaut9000, Jeriel, Suriya, ...)
user_submit/                submission forms emailed in by other people; dropped here, they show on the site
gene_names.xlsx             short and long name for every gene name used in the data files
hairpins.xlsx               HCR amplifier hairpins and initiators, B1-B20
tx/                         pre-computed probe designs for every gene in seven species, loaded by the viewer
design_pipeline/            the scripts that produced tx/
templates/                  HCR_probe_submission.xlsx (the public submission form, and the script that makes it)
                            and new_designer_TEMPLATE.xlsx (the full in-lab format)
build_viewer.py             builds index.html from data/, user_submit/ and gene_names.xlsx
index.html                  the viewer; open it in any browser, works offline
.github/workflows/          rebuilds and publishes the viewer when files are pushed (GitHub Pages)
```

## The viewer

`index.html` has five tabs. **Lab probe sets** lists every set designed in the lab or submitted by others. **Design from transcriptome**
covers every protein-coding gene in seven species: *Bicyclus anynana* (ilBicAnyn1.1) and zebrafish (CDS only), and
*Nothobranchius furzeri* (killifish), *Drosophila melanogaster*, *Caenorhabditis elegans*, mouse (GRCm39) and human
(GRCh38.p14), designed on full NCBI RefSeq mRNAs. Choose a gene, the isoforms to target (all, or only the regions unique
to one variant), the probe location (CDS by default; 5′UTR and 3′UTR can be added for the five mRNA-based species), the
number of pairs, the amplifier (B1-B3) and the fluorophore, then add the set to the order. The isoform table gives the
5′UTR, CDS and 3′UTR coordinates of every transcript and labels each isoform-unique region by the part it lies in. **Submit your probes** offers the submission form and explains how to send it in. **Amplifiers & hairpins** lists the H1, H2, I1 and I2 sequences for B1-B20. The order tray builds an IDT oPool sheet,
an oligo list and a hairpin list, and warns when two genes share an amplifier or a fluorophore. Keep `index.html` and the
`tx/` folder together; the page then works offline as well as on GitHub Pages.

The transcriptome designs are computational and untested. Rules: two 25-nt arms with a 2-nt gap; arm GC 30-80% and within
30 points of each other; no run of 6 identical bases or GGGGG/CCCCC; no 18-mer shared with another gene (its CDS for
B. anynana and zebrafish; any of its transcripts, UTRs and non-coding RNAs included, for the other five species); for
B. anynana each arm site also occurs at most once in the genome. Pairs are ranked toward 50% GC and balanced arms, with penalties
for GGGG/CCCC and runs of 5.

## Submissions from other people

The **Submit your probes** tab offers `templates/HCR_probe_submission.xlsx` for download. People fill in their details,
one row per gene (short and long name, species, Gene ID or LOC, amplifier, whether it worked, and optional details) and
one row per probe pair, then email the file. The address is shown only after a click and is stored encoded in the page,
so address-harvesting bots do not pick it up.

When a form arrives, upload it into `user_submit/` (on GitHub: open the folder, **Add file → Upload files**, commit).
The site rebuilds itself and the sets appear under the submitter's name, usually within two minutes. The build checks
every pair itself and lists anything it skipped in the Actions log; see `user_submit/README.md` for details. To change
the form, edit `templates/make_submission_template.py` and run it.

## Adding probes (lab members)

1. A new person copies `templates/new_designer_TEMPLATE.xlsx` into `data/`, renames it with their
   name (for example `data/Alex.xlsx`), and follows the instructions on its README sheet.
2. Existing designers add rows to their own file in `data/`.
3. New gene names get a row in `gene_names.xlsx` (short name, long name, other names).
4. Save in Excel, then rebuild the viewer: `python build_viewer.py` (needs Python 3 and `openpyxl`).

The build checks that every oligo carries the initiator of its set's amplifier and notes any set
that shares oligos with another person's set.

