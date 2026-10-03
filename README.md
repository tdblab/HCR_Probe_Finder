# HCR probe database

All HCR probe sets designed in the lab, one Excel file per person, plus a searchable viewer
(`index.html`) built from them.

```
data/                       one workbook per designer (Dragonaut9000, Jeriel, Suriya, ...)
gene_names.xlsx             short and long name for every gene name used in the data files
hairpins.xlsx               HCR amplifier hairpins and initiators, B1-B20
tx/                         pre-computed probe designs for every gene (B. anynana and zebrafish), loaded by the viewer
design_pipeline/            the scripts that produced tx/
templates/                  new_designer_TEMPLATE.xlsx, the starting point for a new person
build_viewer.py             builds index.html from data/ and gene_names.xlsx
index.html                  the viewer; open it in any browser, works offline
.github/workflows/          rebuilds and publishes the viewer when files are pushed (GitHub Pages)
```

## The viewer

`index.html` has three tabs. **Lab probe sets** lists every set designed in the lab. **Design from transcriptome**
covers every gene in the B. anynana (ilBicAnyn1.1) and zebrafish CDS: choose a gene, the isoforms to target (all, or only
the regions unique to one variant), the number of pairs, the amplifier (B1-B3) and the fluorophore, then add the set to the
order. **Amplifiers & hairpins** lists the H1, H2, I1 and I2 sequences for B1-B20. The order tray builds an IDT oPool sheet,
an oligo list and a hairpin list, and warns when two genes share an amplifier or a fluorophore. Keep `index.html` and the
`tx/` folder together; the page then works offline as well as on GitHub Pages.

The transcriptome designs are computational and untested. Rules: two 25-nt arms with a 2-nt gap; arm GC 30-80% and within
30 points of each other; no run of 6 identical bases or GGGGG/CCCCC; no 18-mer shared with another gene's CDS; for
B. anynana each arm site occurs at most once in the genome. Pairs are ranked toward 50% GC and balanced arms, with penalties
for GGGG/CCCC and runs of 5.

## Adding probes

1. A new person copies `templates/new_designer_TEMPLATE.xlsx` into `data/`, renames it with their
   name (for example `data/Alex.xlsx`), and follows the instructions on its README sheet.
2. Existing designers add rows to their own file in `data/`.
3. New gene names get a row in `gene_names.xlsx` (short name, long name, other names).
4. Save in Excel, then rebuild the viewer: `python build_viewer.py` (needs Python 3 and `openpyxl`).

The build checks that every oligo carries the initiator of its set's amplifier and notes any set
that shares oligos with another person's set.

## Publishing on GitHub Pages

1. On GitHub, create a new repository named `hcr-probe-finder` (owner `tdblab`).
2. Upload this folder with git (the web uploader takes at most 100 files at a time, and `tx/` has more than 400):
   ```
   cd HCR_probe_database
   git init -b main
   git add .
   git commit -m "HCR probe finder"
   git remote add origin https://github.com/tdblab/hcr-probe-finder.git
   git push -u origin main
   ```
3. In the repository, open **Settings > Pages** and set **Source** to **GitHub Actions**.
   The workflow in `.github/workflows/` builds and publishes the page; watch it on the **Actions** tab.
4. Because tirthadasbanerjee.com is the custom domain of tdblab.github.io, the page appears at
   **https://tirthadasbanerjee.com/hcr-probe-finder/** (also at https://tdblab.github.io/hcr-probe-finder/).

To update: edit the Excel files, then `git add . && git commit -m "update" && git push`. The page rebuilds itself.

GitHub Pages sites are public on the internet, even when the repository is private, unless the
repository belongs to an organization on GitHub Enterprise Cloud, which can publish a site privately
to people with read access to the repository. If the probe sequences should stay within the lab,
keep the repository private without Pages and share `index.html` directly, or use Enterprise Cloud.
