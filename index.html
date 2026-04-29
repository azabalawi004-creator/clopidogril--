from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# Page margins
section = doc.sections[0]
section.top_margin    = Inches(1)
section.bottom_margin = Inches(1)
section.left_margin   = Inches(1)
section.right_margin  = Inches(1)

styles = doc.styles
styles['Normal'].font.name = 'Arial'
styles['Normal'].font.size = Pt(11)

# ── helpers ──────────────────────────────────────────────────────────
def add_rule(color='2E74B5'):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bot  = OxmlElement('w:bottom')
    bot.set(qn('w:val'),   'single')
    bot.set(qn('w:sz'),    '6')
    bot.set(qn('w:space'), '1')
    bot.set(qn('w:color'), color)
    pBdr.append(bot)
    pPr.append(pBdr)
    p.paragraph_format.space_after = Pt(4)

def h1(text, color=(0x1F, 0x49, 0x7D)):
    p = doc.add_heading(text, level=1)
    r = p.runs[0]
    r.font.name = 'Arial'; r.font.size = Pt(14); r.font.bold = True
    r.font.color.rgb = RGBColor(*color)
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after  = Pt(6)

def h2(text, color=(0x2E, 0x74, 0xB5)):
    p = doc.add_heading(text, level=2)
    r = p.runs[0]
    r.font.name = 'Arial'; r.font.size = Pt(12); r.font.bold = True
    r.font.color.rgb = RGBColor(*color)
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after  = Pt(4)

def body(text, italic=False):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.name = 'Arial'; r.font.size = Pt(11); r.italic = italic
    p.paragraph_format.space_after = Pt(6)

def bullet(label, detail=''):
    p = doc.add_paragraph(style='List Bullet')
    if detail:
        rb = p.add_run(label)
        rb.bold = True; rb.font.name = 'Arial'; rb.font.size = Pt(11)
        rd = p.add_run(detail)
        rd.font.name = 'Arial'; rd.font.size = Pt(11)
    else:
        r = p.add_run(label)
        r.font.name = 'Arial'; r.font.size = Pt(11)
    p.paragraph_format.space_after = Pt(3)

# ── TITLE ─────────────────────────────────────────────────────────────
tp = doc.add_paragraph()
tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
tr = tp.add_run('Concept & Sourcing Report')
tr.font.name = 'Arial'; tr.font.size = Pt(18); tr.font.bold = True
tr.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

sp = doc.add_paragraph()
sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
sr = sp.add_run('VR Educational Module: Clopidogrel Activation & P2Y12 Receptor Inhibition')
sr.font.name = 'Arial'; sr.font.size = Pt(13)
sr.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)

mp = doc.add_paragraph()
mp.alignment = WD_ALIGN_PARAGRAPH.CENTER
mr = mp.add_run('Medicinal Chemistry | Group Assignment | Spring 2026')
mr.font.name = 'Arial'; mr.font.size = Pt(10); mr.italic = True
mr.font.color.rgb = RGBColor(0x70, 0x70, 0x70)
mp.paragraph_format.space_after = Pt(10)

add_rule()

# ── SECTION 1: CONCEPT ───────────────────────────────────────────────
h1('1. Chosen Medicinal Chemistry Concept')

h2('Concept Title')
body('Clopidogrel Activation and P2Y12 Receptor Inhibition: A 3D VR Educational Model')

h2('Scientific Mechanism')
body(
    'Clopidogrel (Plavix) is an orally administered antiplatelet thienopyridine prodrug. '
    'It is pharmacologically inactive until it undergoes two-step hepatic oxidation by '
    'CYP450 enzymes (principally CYP2C19), yielding a reactive thiol metabolite. This '
    'active metabolite irreversibly and selectively blocks the P2Y12 purinergic receptor '
    'on platelet surfaces via covalent disulfide bonds with Cys97 and Cys175.'
)
body(
    'Under normal physiology, ADP (released from damaged vessel walls and activated platelets) '
    'binds to P2Y12, activating Gi proteins that suppress cAMP and trigger platelet '
    'activation and aggregation. Clopidogrel interrupts this pathway permanently for the '
    'lifetime of the platelet (~7-10 days), dramatically reducing thrombotic risk.'
)

h2('Why VR Enhances Learning of This Concept')
body(
    'This mechanism spans three distinct anatomical compartments — the oral route, the liver, '
    'and the platelet membrane receptor — making it inherently 3D and spatial. Key reasons '
    'VR is uniquely suited to teaching this topic:'
)
bullet('Prodrug-to-drug transformation: visualising the molecular change inside a 3D liver environment is more intuitive than a 2D reaction arrow')
bullet('GPCR structure: the transmembrane topology of P2Y12 is a 3D concept that flat diagrams approximate poorly')
bullet('Covalent vs non-covalent binding: students can see the irreversible nature of the bond, contrasting with the reversible ADP interaction, in the same spatial scene')
bullet('ADP competition: physically watching ADP get blocked by a pre-bound metabolite is far more memorable than reading "competitive inhibition" on a slide')
bullet('Clinical consequence chain: the linear station layout (prodrug → liver → receptor → ADP blocked → reduced clots) creates a clear spatial narrative')
body(
    'Embodied learning in VR creates procedural memory. Walking through the mechanism '
    'station by station encodes the sequence in a way that 2D slides cannot replicate.',
    italic=True
)

# ── SECTION 2: STORYBOARD ────────────────────────────────────────────
h1('2. Visual Storyboard — User Journey')

body('The VR module is divided into five linear stations arranged left-to-right in a shared 3D space. The user begins at Station 1 and walks right (WASD controls or physical VR movement) to progress through the mechanism.')

from docx.shared import Pt as Ptx
from docx.oxml.ns import qn

stations = [
    ('1', 'Inactive Prodrug',
     'User sees Clopidogrel as a blue rotating molecular model. Click reveals: "Clopidogrel is an antiplatelet prodrug requiring liver activation." Learning goal: understand that inactive prodrugs exist and why.'),
    ('2', 'Liver Activation by CYP450',
     'A large reddish-brown liver shape dominates the station. A green active metabolite floats nearby, bobbing gently. Click on liver reveals CYP450 enzyme detail. Learning goal: understand hepatic biotransformation.'),
    ('3', 'P2Y12 Receptor',
     'A tall purple column with torus rings (GPCR-like) represents the receptor. Binding pocket glows at the top. G-protein sphere at base. Click reveals GPCR signalling cascade. Learning goal: understand receptor architecture and function.'),
    ('4', 'Normal ADP Binding',
     'Yellow ADP molecule animates toward an unblocked P2Y12 receptor. An arrow reinforces the direction. Click on ADP reveals its physiological role. Learning goal: establish the normal baseline before inhibition.'),
    ('5', 'Clopidogrel Blocks Receptor',
     'Green active metabolite is covalently attached. The receptor glows red. A red X stops the dimmed ADP from reaching the site. Click reveals irreversibility. Learning goal: understand how covalent, irreversible inhibition differs from competitive inhibition.'),
]

tbl = doc.add_table(rows=1, cols=3)
tbl.style = 'Table Grid'
from docx.shared import Inches as In2
hdr = tbl.rows[0].cells
for i, h in enumerate(['Scene', 'Title', 'Description & Learning Goal']):
    hdr[i].width = In2([0.7, 2.0, 3.8][i])
    rn = hdr[i].paragraphs[0].add_run(h)
    rn.bold = True; rn.font.name = 'Arial'; rn.font.size = Pt(10)
    rn.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    tcPr = hdr[i]._tc.get_or_add_tcPr()
    shd  = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  '1F497D')
    tcPr.append(shd)

for sc, title, desc in stations:
    row = tbl.add_row().cells
    r0 = row[0].paragraphs[0].add_run(sc)
    r0.bold = True; r0.font.name = 'Arial'; r0.font.size = Pt(10)
    row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = row[1].paragraphs[0].add_run(title)
    r1.bold = True; r1.font.name = 'Arial'; r1.font.size = Pt(10)
    r2 = row[2].paragraphs[0].add_run(desc)
    r2.font.name = 'Arial'; r2.font.size = Pt(9.5)

doc.add_paragraph()

# ── SECTION 3: ASSETS ────────────────────────────────────────────────
h1('3. Molecular Assets — Sources')

h2('Primary Molecular Reference')
bullet('Database: ', 'RCSB Protein Data Bank (rcsb.org)')
bullet('PDB IDs considered: ', '2PXY (P2Y12 with antagonist) and 4NTJ (P2Y12 with AZD1283 — highest resolution human structure at 2.62 A)')
bullet('Use: ', 'Reference for receptor topology and binding pocket geometry; manually recreated as A-Frame primitives')
bullet('Access: ', 'https://www.rcsb.org/structure/4NTJ')

h2('Drug Structure Reference')
bullet('Source: ', 'MolView (molview.org) — CID 2724635 (Clopidogrel) and CID 25151352 (active thiol metabolite)')
bullet('Use: ', 'Atom counts and connectivity reference for the stylised 3D sphere-and-cylinder models built in A-Frame')
bullet('Format: ', 'Viewed online; no file download required (primitives hand-coded in HTML)')

h2('ADP Structure Reference')
bullet('Source: ', 'RCSB PDB Ligand database — ADP (Adenosine 5-diphosphate, PDB ligand ID: ADP)')
bullet('Use: ', 'Atom layout reference for the yellow ADP model at Station 4')

h2('Environment & Platform')
bullet('VR Framework: ', 'A-Frame 1.5.0 (aframe.io) — free, open-source WebXR framework')
bullet('Hosting: ', 'Glitch.com (free tier) — no account or payment required; generates a permanent public URL')
bullet('3D assets: ', 'All geometry built from A-Frame primitives (spheres, cylinders, torus, cones, boxes, planes). No external 3D files required.')
bullet('Compatibility: ', 'Works in any modern browser (desktop and mobile) and in WebXR headsets (Meta Quest, HTC Vive, Valve Index)')

# ── SECTION 4: PLATFORM ──────────────────────────────────────────────
h1('4. Platform & Technical Approach')
body(
    'The complete VR module is a single self-contained index.html file using the A-Frame '
    'CDN. It requires no installation, no paid services, and no external assets that could '
    'break. The file is uploaded to Glitch.com and set to Public visibility.'
)
bullet('A-Frame v1.5.0 loaded from CDN: https://aframe.io/releases/1.5.0/aframe.min.js')
bullet('Five educational stations on a single linear scene (X-axis layout)')
bullet('Click interactions on all five key objects showing pharmacological detail')
bullet('Animations: molecular rotation (prodrug), pulsing glow (blocked receptor), ADP approach/rebound, active metabolite bobbing')
bullet('WASD keyboard navigation + mouse look; full WebXR VR headset support via built-in A-Frame VR button')

h2('Sharing Instructions')
body(
    'Before submission: open the Glitch project, click Share, set to "Public". '
    'Paste the direct URL (e.g. https://your-project.glitch.me) at the top of this '
    'document and in the assignment portal text box.'
)

add_rule()

fp = doc.add_paragraph()
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
fr = fp.add_run('Submitted for Medicinal Chemistry VR Assignment  |  Spring 2026  |  All assets sourced from free, publicly accessible databases.')
fr.font.name = 'Arial'; fr.font.size = Pt(9); fr.italic = True
fr.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

out = '/mnt/user-data/outputs/Clopidogrel_VR_Concept_Report.docx'
doc.save(out)
print('Saved:', out)
