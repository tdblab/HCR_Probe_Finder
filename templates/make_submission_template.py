"""Make templates/HCR_probe_submission.xlsx, the form people download from the "Submit your probes" tab.

Usage: python make_submission_template.py [output.xlsx]

No macros: the sheet works in Excel, Google Sheets, LibreOffice and Numbers. Yellow cells are required, white cells
optional, gray cells are filled in by the sheet. build_viewer.py reads only the typed-in cells, never the formulas,
so a file saved by any program works.
"""
import sys, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.workbook.defined_name import DefinedName

OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), 'HCR_probe_submission.xlsx')
MAXG, MAXP = 30, 1000                     # genes per file, probe-pair rows
G0, P0 = 6, 5                             # first data row on the Genes and Probes sheets
SITE = 'https://tirthadasbanerjee.com/HCR_Probe_Finder/#submit'

INK, MUTED = '18202B', '5B6573'
YEL, GRAY, SOFT, GREEN, ORANGE = 'FFF3B0', 'E6E8EB', 'F5F6F4', 'DDF2E4', 'FFE2C2'
f_title = Font(name='Calibri', size=18, bold=True, color=INK)
f_big = Font(name='Calibri', size=13, color=INK)
f_bigb = Font(name='Calibri', size=13, bold=True, color=INK)
f_head = Font(name='Calibri', size=12, bold=True, color='FFFFFF')
f_norm = Font(name='Calibri', size=12, color=INK)
f_hint = Font(name='Calibri', size=11, italic=True, color=MUTED)
f_seq = Font(name='Consolas', size=11, color=INK)
fill = lambda c: PatternFill('solid', start_color=c, end_color=c)
thin = Side(style='thin', color='C9CED4')
box = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap = Alignment(wrap_text=True, vertical='top')
center = Alignment(horizontal='center', vertical='top')
OPEN = Protection(locked=False)

AMPS = [('B1', 'AF546', 'GAGGAGGGCAGCAAACGG', 'AA', 'TA', 'GAAGAGTCTTCCTTTACG'),
        ('B2', 'AF647', 'CCTCGTAAATCCTCATCA', 'AA', 'AA', 'ATCATCCAGTAAACCGCC'),
        ('B3', 'AF488', 'GTCCCTGCCTCTATATCT', 'TT', 'TT', 'CCACTCAACTTTAACCCG'),
        ('B4', 'AF514', 'CCTCAACCTACCTCCAAC', 'AA', 'AT', 'TCTCACCATATTCGCTTC'),
        ('B5', 'AF594', 'CTCACTCCCAATCTCTAT', 'AA', 'AA', 'CTACCCTACAAATCCAAT')]
SPECIES = ['Bicyclus anynana', 'Danio rerio', 'Nothobranchius furzeri', 'Drosophila melanogaster',
           'Caenorhabditis elegans', 'Mus musculus', 'Homo sapiens', 'Other: type the species name']
DYES = ['AF488', 'AF514', 'AF546', 'AF594', 'AF647', 'Other: type it']
REGION = ['CDS', 'CDS and UTRs', "5′UTR", "3′UTR", 'Intron', 'Other: say in Notes']
METHOD = ['HCR probe finder (this site)', 'Molecular Instruments', 'insitu_probe_generator', 'Own design / other tool']
RESULT = ['Works well', 'Weak signal', 'Did not work', 'Not tested yet']

wb = Workbook()

# ---------------------------------------------------------------- Lists (hidden)
L = wb.active; L.title = 'Lists'
L.append(['Amplifier', 'Lab default dye', 'P1 initiator half', 'P1 spacer', 'P2 spacer', 'P2 initiator half', '',
          'Species', 'Dye', 'Where the probes bind', 'How designed', 'Did it work?', 'Gene #'])
for i in range(max(len(AMPS), len(SPECIES), MAXG)):
    row = list(AMPS[i]) if i < len(AMPS) else [None] * 6
    row += [None, SPECIES[i] if i < len(SPECIES) else None, DYES[i] if i < len(DYES) else None,
            REGION[i] if i < len(REGION) else None, METHOD[i] if i < len(METHOD) else None,
            RESULT[i] if i < len(RESULT) else None, i + 1 if i < MAXG else None]
    L.append(row)
L['O1'] = 'template'; L['O2'] = 'HCR probe submission v1'          # lets the build recognize this form
L.sheet_state = 'hidden'
def lst(col, n): return f"Lists!${col}$2:${col}${n + 1}"
for name, ref in [('AmpList', lst('A', 5)), ('SpeciesList', lst('H', len(SPECIES))), ('DyeList', lst('I', len(DYES))),
                  ('RegionList', lst('J', len(REGION))), ('MethodList', lst('K', len(METHOD))),
                  ('ResultList', lst('L', len(RESULT))), ('GeneNumbers', lst('M', MAXG)),
                  ('P1Halves', lst('C', 5)), ('P2Halves', lst('F', 5))]:
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

def dv(ws, ref, src, prompt=None, strict=True):
    d = DataValidation(type='list', formula1=f'={src}', allow_blank=True, showDropDown=False,
                       errorStyle='stop' if strict else 'warning',
                       error='Please pick from the list.' if strict else 'Not in the list. Click Yes to keep what you typed.',
                       errorTitle='Check this cell')
    if prompt: d.promptTitle, d.prompt, d.showInputMessage = 'Tip', prompt, True
    ws.add_data_validation(d); d.add(ref)

# ---------------------------------------------------------------- Start here
H = wb.create_sheet('Start here', 0)
H.column_dimensions['A'].width = 4; H.column_dimensions['B'].width = 6; H.column_dimensions['C'].width = 100
H['C2'] = 'HCR probe submission form'; H['C2'].font = f_title
H['C3'] = 'Share the HCR probe sets you have designed, so others can find and reuse them.'; H['C3'].font = f_big
steps = [('1', 'Fill in the sheet "1 About you" (five boxes).'),
         ('2', 'In "2 Genes", type how many genes you are sending (up to 30). Then fill one row per gene.'),
         ('3', 'In "3 Probes", add one row per probe pair: pick the gene number, then paste P1 and P2.'),
         ('4', 'Look at the "Check" column. Every pair should say OK ✓. Fix anything it flags.'),
         ('5', 'Save the file and email it to the address on the "Submit your probes" page:')]
r = 5
for k, t in steps:
    H.cell(r, 2, k).font = Font(name='Calibri', size=14, bold=True, color='FFFFFF'); H.cell(r, 2).fill = fill(INK)
    H.cell(r, 2).alignment = Alignment(horizontal='center', vertical='center')
    H.cell(r, 3, t).font = f_big; H.cell(r, 3).alignment = Alignment(vertical='center', wrap_text=True)
    H.row_dimensions[r].height = 26; r += 2
H.cell(r - 1, 3, SITE).font = Font(name='Calibri', size=13, color='1F5FBF', underline='single')
H.cell(r - 1, 3).hyperlink = SITE
r += 1
H.cell(r, 3, 'What the colors mean').font = f_bigb; r += 1
for c, t in [(YEL, 'Yellow: please fill in'), ('FFFFFF', 'White: optional, but helpful'), (GRAY, 'Gray: filled in for you, or not needed')]:
    H.cell(r, 2).fill = fill(c); H.cell(r, 2).border = box
    H.cell(r, 3, t).font = f_big; r += 1
r += 1
H.cell(r, 3, 'Good to know').font = f_bigb; r += 1
for t in ['Everything you send will be shown publicly on the HCR probe finder website: gene names, sequences, your name and lab. '
          'Please do not send sets you want to keep private.',
          'Paste full oligos, 5′ to 3′, exactly as ordered. P1 starts with the initiator half; P2 ends with it. '
          'The sheet recognizes amplifiers B1 to B5 from the sequence.',
          'Have more than 30 genes? Send a second file. Want to fix a set later? Send a corrected file and say which one it replaces.']:
    H.cell(r, 3, '•  ' + t).font = f_norm; H.cell(r, 3).alignment = wrap; H.row_dimensions[r].height = 34; r += 1
H.sheet_view.showGridLines = False

# ---------------------------------------------------------------- 1 About you
A = wb.create_sheet('1 About you', 1)
A.column_dimensions['A'].width = 34; A.column_dimensions['B'].width = 44; A.column_dimensions['C'].width = 56
A['A1'] = 'About you'; A['A1'].font = f_title
A['A2'] = 'This shows on the website next to your probe sets.'; A['A2'].font = f_hint
about = [('Your name', 'e.g. Jane Tan', None),
         ('Initials for your set IDs', 'e.g. JT  (2 to 4 letters; your sets become JT-001, JT-002…)', None),
         ('Lab / PI', 'e.g. Smith lab', None),
         ('Institution', 'e.g. National University of Singapore', None),
         ('Date of this submission', 'e.g. 18/10/2026', 'date')]
for i, (lab, hint, kind) in enumerate(about):
    r = 4 + i
    A.cell(r, 1, lab).font = f_bigb
    c = A.cell(r, 2); c.fill = fill(YEL); c.border = box; c.font = f_big; c.protection = OPEN
    if kind == 'date': c.number_format = 'yyyy-mm-dd'
    A.cell(r, 3, hint).font = f_hint
    A.row_dimensions[r].height = 24
dv_init = DataValidation(type='textLength', operator='between', formula1='2', formula2='4', allow_blank=True,
                         error='Use 2 to 4 letters, for example JT.', errorTitle='Initials')
A.add_data_validation(dv_init); dv_init.add('B5')
A['A11'] = 'Your email address is not needed here: this file becomes public, and you will email it anyway.'; A['A11'].font = f_hint
A.sheet_view.showGridLines = False
A.protection.sheet = True; A.protection.formatColumns = False; A.protection.formatRows = False

# ---------------------------------------------------------------- 2 Genes
S = wb.create_sheet('2 Genes', 2)
cols = [  # header, width, required, hint, list, strict list
    ('Gene #', 8, False, '', None, True),
    ('Gene short name', 16, True, 'e.g. hth', None, True),
    ('Gene long name', 26, True, 'e.g. homothorax', None, True),
    ('Species', 24, True, 'pick, or type another', 'SpeciesList', False),
    ('Gene ID or LOC', 18, True, 'LOC number or NCBI Gene ID', None, True),
    ('Transcript targeted', 18, False, 'XM_ / NM_ accession', None, True),
    ('Amplifier', 11, True, 'B1 to B5', 'AmpList', True),
    ('Dye you used', 13, False, 'e.g. AF546', 'DyeList', False),
    ('Where the probes bind', 17, False, 'CDS, UTR, intron…', 'RegionList', True),
    ('How you designed them', 24, False, 'pick one', 'MethodList', False),
    ('Did it work?', 16, True, 'pick one', 'ResultList', True),
    ('Tissue / stage tested', 22, False, 'e.g. 5th instar wing disc', None, True),
    ('Order date', 13, False, 'e.g. 2026-09-01', None, True),
    ('Vendor / order ID', 18, False, 'e.g. IDT oPool + order no.', None, True),
    ('Notes', 30, False, 'anything else useful', None, True),
    ('Pairs entered', 10, None, 'counted for you', None, True),
    ('Row check', 30, None, 'filled in for you', None, True)]
S['A1'] = 'Genes'; S['A1'].font = f_title
S['A2'] = 'One row per gene. Yellow = required. Rows below the number you type stay gray.'; S['A2'].font = f_hint
S['A3'] = 'How many genes are you sending?'; S['A3'].font = f_bigb
S.merge_cells('A3:C3')
n = S['D3']; n.value = 1; n.fill = fill(YEL); n.border = box; n.font = Font(name='Calibri', size=14, bold=True); n.protection = OPEN
n.alignment = Alignment(horizontal='center')
dv_n = DataValidation(type='whole', operator='between', formula1='1', formula2=str(MAXG), error=f'Type a number from 1 to {MAXG}.',
                      errorTitle='Number of genes', promptTitle='Tip', prompt=f'Type 1 to {MAXG}', showInputMessage=True)
S.add_data_validation(dv_n); dv_n.add('D3')
S['E3'] = f'(1 to {MAXG})'; S['E3'].font = f_hint
for j, (hdr, w, req, hint, src, strict) in enumerate(cols, 1):
    L_ = S.cell(1, j).column_letter
    S.column_dimensions[L_].width = w
    h = S.cell(5, j, hdr + (' *' if req else '')); h.font = f_head; h.fill = fill(INK if req is not None else MUTED); h.alignment = Alignment(wrap_text=True, vertical='center'); h.border = box
    S.cell(4, j, hint).font = f_hint; S.cell(4, j).alignment = Alignment(wrap_text=True, vertical='bottom')
    if src: dv(S, f'{L_}{G0}:{L_}{G0 + MAXG - 1}', src, strict=strict)
S.row_dimensions[4].height = 30; S.row_dimensions[5].height = 34
last = G0 + MAXG - 1
for i in range(MAXG):
    r = G0 + i
    S.cell(r, 1, i + 1).alignment = center; S.cell(r, 1).font = f_bigb
    for j in range(1, len(cols) + 1):
        c = S.cell(r, j); c.border = box
        c.font = f_bigb if j == 1 else f_norm
        if 2 <= j <= 15: c.protection = OPEN
        c.alignment = wrap if j > 1 else center
    S.cell(r, 13).number_format = 'yyyy-mm-dd'
    S.cell(r, 16, f"=IF(B{r}=\"\",\"\",COUNTIF('3 Probes'!$A${P0}:$A${P0 + MAXP - 1},A{r}))").alignment = center
    need = '&'.join(f'IF({c}{r}="","{lab}, ","")' for c, lab in
                    [('B', 'short name'), ('C', 'long name'), ('D', 'species'), ('E', 'Gene ID'), ('G', 'amplifier'), ('K', 'did it work')])
    S.cell(r, 17, f'=IF(A{r}>$D$3,"",IF(({need})<>"","Missing: "&LEFT({need},LEN({need})-2),'
                  f'IF(P{r}=0,"No pairs yet in 3 Probes","Ready ✓")))')
    S.cell(r, 16).fill = fill(SOFT); S.cell(r, 17).fill = fill(SOFT)
# gray out rows past the number of genes; yellow for required cells still empty
S.conditional_formatting.add(f'A{G0}:Q{last}', FormulaRule(formula=[f'$A{G0}>$D$3'], fill=fill(GRAY), font=Font(color='9AA3AD'), stopIfTrue=True))
for c in 'BCDEGK':
    S.conditional_formatting.add(f'{c}{G0}:{c}{last}', FormulaRule(formula=[f'AND($A{G0}<=$D$3,{c}{G0}="")'], fill=fill(YEL)))
S.conditional_formatting.add(f'Q{G0}:Q{last}', FormulaRule(formula=[f'LEFT(Q{G0},5)="Ready"'], fill=fill(GREEN)))
S.conditional_formatting.add(f'Q{G0}:Q{last}', FormulaRule(formula=[f'AND(Q{G0}<>"",LEFT(Q{G0},5)<>"Ready")'], fill=fill(ORANGE)))
S.freeze_panes = 'C6'
S.protection.sheet = True; S.protection.formatColumns = False; S.protection.formatRows = False

# ---------------------------------------------------------------- 3 Probes
P = wb.create_sheet('3 Probes', 3)
pc = [('Gene #', 9, True, 'from 2 Genes'), ('Gene', 16, None, 'filled in for you'), ('Pair #', 8, None, 'counted for you'),
      ('P1 sequence (5′→3′)', 56, True, 'full oligo: initiator half + spacer + arm'),
      ('P2 sequence (5′→3′)', 56, True, 'full oligo: arm + spacer + initiator half'),
      ('Check', 36, None, 'should say OK ✓'),
      ('Your P1 name', 20, False, 'optional, e.g. hth_B1_P1_01'), ('Your P2 name', 20, False, 'optional')]
P['A1'] = 'Probes'; P['A1'].font = f_title
P['A2'] = 'One row per probe pair. Pick the gene number from "2 Genes", then paste P1 and P2. You can paste a whole column at once.'; P['A2'].font = f_hint
P['A3'] = 'Pairs entered:'; P['A3'].font = f_bigb
P['C3'] = f'=COUNTA(D{P0}:D{P0 + MAXP - 1})'; P['C3'].font = f_bigb
P['D3'] = f'=IF(COUNTIF(F{P0}:F{P0 + MAXP - 1},"OK*")=COUNTA(F{P0}:F{P0 + MAXP - 1})-COUNTBLANK(F{P0}:F{P0 + MAXP - 1}),"All pairs look fine",' \
          f'(COUNTA(F{P0}:F{P0 + MAXP - 1})-COUNTBLANK(F{P0}:F{P0 + MAXP - 1})-COUNTIF(F{P0}:F{P0 + MAXP - 1},"OK*"))&" pair(s) need a look: see the Check column")'
P['D3'].font = f_bigb
for j, (hdr, w, req, hint) in enumerate(pc, 1):
    L_ = P.cell(1, j).column_letter; P.column_dimensions[L_].width = w
    h = P.cell(4, j, hdr + (' *' if req else '')); h.font = f_head; h.fill = fill(INK if req is not None else MUTED); h.border = box
    h.alignment = Alignment(wrap_text=True, vertical='center')
P.row_dimensions[4].height = 30
P['F3'] = 'Tip: P1 = initiator half + 2-nt spacer + 25-nt arm (45 nt). P2 = 25-nt arm + 2-nt spacer + initiator half (45 nt).'; P['F3'].font = f_hint
P.merge_cells('F3:H3')
dv(P, f'A{P0}:A{P0 + MAXP - 1}', 'GeneNumbers', prompt='Pick the gene number from sheet 2')
# helper columns (hidden): J amp on sheet 2, K/L cleaned P1/P2, M/N amplifier found in P1/P2, O letters other than ACGT
for c, h in zip('JKLMNO', ['amp chosen', 'P1 clean', 'P2 clean', 'P1 amp', 'P2 amp', 'bad letters']):
    P[f'{c}4'] = h; P.column_dimensions[c].hidden = True
end = P0 + MAXP - 1
for r in range(P0, end + 1):
    for j in range(1, 9):
        c = P.cell(r, j); c.border = box
        c.font = f_seq if j in (4, 5) else f_norm
        if j in (1, 4, 5, 7, 8): c.protection = OPEN
    P.cell(r, 1).alignment = center; P.cell(r, 3).alignment = center
    P.cell(r, 2, f"=IF(A{r}=\"\",\"\",INDEX('2 Genes'!$B${G0}:$B${last},A{r})&\"\")")
    P.cell(r, 3, f'=IF(A{r}="","",COUNTIF(A${P0}:A{r},A{r}))')
    P.cell(r, 10, f"=IF(A{r}=\"\",\"\",INDEX('2 Genes'!$G${G0}:$G${last},A{r})&\"\")")
    P.cell(r, 11, f'=UPPER(SUBSTITUTE(SUBSTITUTE(TRIM(D{r})," ",""),"-",""))')
    P.cell(r, 12, f'=UPPER(SUBSTITUTE(SUBSTITUTE(TRIM(E{r})," ",""),"-",""))')
    P.cell(r, 13, f'=IFERROR(INDEX(AmpList,MATCH(LEFT(K{r},18),P1Halves,0)),"?")')
    P.cell(r, 14, f'=IFERROR(INDEX(AmpList,MATCH(RIGHT(L{r},18),P2Halves,0)),"?")')
    P.cell(r, 15, f'=LEN(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(K{r}&L{r},"A",""),"C",""),"G",""),"T",""))')
    P.cell(r, 6, f'=IF(AND(A{r}="",D{r}="",E{r}=""),"",IF(A{r}="","Pick the gene #",IF(OR(D{r}="",E{r}=""),"P1 or P2 is missing",'
                 f'IF(O{r}>0,"Only A, C, G and T are allowed",IF(M{r}="?","P1 does not start with a B1–B5 initiator",'
                 f'IF(N{r}="?","P2 does not end with a B1–B5 initiator",IF(M{r}<>N{r},"P1 is "&M{r}&" but P2 is "&N{r},'
                 f'IF(AND(J{r}<>"",M{r}<>J{r}),"These are "&M{r}&" probes; 2 Genes says "&J{r},'
                 f'IF(OR(LEN(K{r})<>45,LEN(L{r})<>45),"OK ✓ ("&M{r}&"; lengths "&LEN(K{r})&" and "&LEN(L{r})&" nt, usually 45)",'
                 f'"OK ✓ ("&M{r}&")")))))))))')
    for j in (2, 3, 6): P.cell(r, j).fill = fill(SOFT)
P.conditional_formatting.add(f'F{P0}:F{end}', FormulaRule(formula=[f'LEFT(F{P0},2)="OK"'], fill=fill(GREEN)))
P.conditional_formatting.add(f'F{P0}:F{end}', FormulaRule(formula=[f'AND(F{P0}<>"",LEFT(F{P0},2)<>"OK")'], fill=fill(ORANGE)))
for c in 'ADE':
    P.conditional_formatting.add(f'{c}{P0}:{c}{end}', FormulaRule(formula=[f'AND(COUNTA($A{P0},$D{P0},$E{P0})>0,{c}{P0}="")'], fill=fill(YEL)))
P.freeze_panes = f'D{P0}'
P.protection.sheet = True; P.protection.formatColumns = False; P.protection.formatRows = False

for ws in (H, A, S, P):
    ws.sheet_properties.tabColor = INK
    ws.page_setup.orientation = 'landscape'; ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
wb.active = 0
from openpyxl.workbook.properties import CalcProperties
wb.calculation = CalcProperties(fullCalcOnLoad=True)
wb.save(OUT)
print('wrote', OUT)
