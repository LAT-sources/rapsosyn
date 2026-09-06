#!/usr/bin/env python3
"""
RAPSORD — Générateur RD0 par remplacement dans le template Word original.
Usage: python generate_from_template.py dossier.json
Produit: RAPSORD_<ref>_RD0.docx avec la mise en page exacte du modèle.
"""
import sys, json, os, shutil, zipfile, re
from datetime import datetime

# ── helpers ───────────────────────────────────────────────────────────────────
def val(D, key, default=''):
    v = D.get(key, default)
    return str(v).strip() if v else default

def r_val(D, key):
    return val(D, 'r_' + key)

def fmt_date(s):
    if not s: return '/'
    try: return datetime.strptime(s, '%Y-%m-%d').strftime('%d/%m/%Y')
    except: return s or '/'

def oui_non(D, key):
    v = r_val(D, key)
    if v == 'OUI': return '☒ OUI  ☐ NON'
    if v == 'NON': return '☐ OUI  ☒ NON'
    return '☐ OUI  ☐ NON'

def conception(D):
    v = r_val(D, 'conc')
    if v == 'Conception':
        return ('☒ Conception', '☐ Conception + exécution')
    return ('☐ Conception', '☒ Conception + exécution')

def domaines(D):
    mc = D.get('mc', {})
    sfs = '☒' if mc.get('Sols/Fondations/Structures') else '☐'
    env = '☒' if mc.get('Enveloppe') else '☐'
    stab = '☒' if mc.get('Stabilité') else '☐'
    et   = '☒' if mc.get('Étanchéité toitures') else '☐'
    ef   = '☒' if mc.get('Étanchéité façades') else '☐'
    ess  = '☒' if mc.get('Étanchéité sous-sols') else '☐'
    return sfs, env, stab, et, ef, ess

# ── XML text replacement ──────────────────────────────────────────────────────
def replace_in_cell(cell_elem, ns, old_text, new_text):
    """Replace full text of a cell, preserving the first run's formatting."""
    W = f'{{{ns}}}'
    runs = list(cell_elem.iter(f'{W}r'))
    if not runs:
        return False
    # Get combined text
    full = ''.join((t.text or '') for r in runs for t in r.iter(f'{W}t'))
    if old_text not in full:
        return False
    # Clear all runs except first, set first run text to new value
    first_r = runs[0]
    # Remove all t elements from first run, add new one
    for t in list(first_r.iter(f'{W}t')):
        first_r.remove(t)
    import xml.etree.ElementTree as ET
    new_t = ET.SubElement(first_r, f'{W}t')
    new_t.text = new_text
    if new_text and (new_text[0] == ' ' or new_text[-1] == ' '):
        new_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    # Remove all other runs
    parent = first_r.getparent() if hasattr(first_r, 'getparent') else None
    # Use a different approach - find the paragraph
    for para in cell_elem.iter(f'{W}p'):
        para_runs = list(para.findall(f'{W}r'))
        if first_r in para_runs:
            # Keep first run, remove others
            for r in para_runs[1:]:
                para.remove(r)
            break
    return True

def set_cell_text(cell_elem, ns, new_text):
    """Set the complete text of a cell, preserving formatting of first run."""
    import xml.etree.ElementTree as ET
    W = f'{{{ns}}}'
    # Find first paragraph
    for para in cell_elem.iter(f'{W}p'):
        runs = list(para.findall(f'{W}r'))
        if runs:
            # Keep only first run, set its text
            first_r = runs[0]
            for t in list(first_r.findall(f'{W}t')):
                first_r.remove(t)
            new_t = ET.SubElement(first_r, f'{W}t')
            new_t.text = new_text or ''
            if new_text and (new_text[0] == ' ' or new_text[-1] == ' '):
                new_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            for r in runs[1:]:
                para.remove(r)
        else:
            # No runs - add one
            new_r = ET.SubElement(para, f'{W}r')
            new_t = ET.SubElement(new_r, f'{W}t')
            new_t.text = new_text or ''
        # Remove extra paragraphs in this cell
        break

# ── main replacement logic ────────────────────────────────────────────────────
def process_xml(xml_content, D):
    """Do all text replacements via string substitution (safest approach)."""
    import re
    
    def escape_xml(s):
        s = str(s)
        s = s.replace('&', '&amp;')
        s = s.replace('<', '&lt;')
        s = s.replace('>', '&gt;')
        s = s.replace('"', '&quot;')
        return s

    def replace_val(xml, old_val, new_val):
        """Replace a specific data value string in the XML."""
        old_esc = escape_xml(old_val)
        new_esc = escape_xml(new_val)
        return xml.replace(old_esc, new_esc)

    mc = D.get('mc', {})
    rd_req = D.get('rd_req', {})
    conc_left, conc_right = conception(D)
    sfs, env, stab, et, ef, ess = domaines(D)

    # ── Références Socotec (page de garde) ───────────────────────────────────
    # Référence
    xml_content = replace_val(xml_content, 'JB000-26-070', val(D, 'ref', 'JB000-26-070'))
    # Dossier
    xml_content = replace_val(xml_content, '2502JB000000025', val(D, 'dossier', '2502JB000000025'))
    # Ing dossier
    xml_content = replace_val(xml_content, 'Laïd Atek', val(D, 'ing_doss', 'Laïd Atek'))
    # Ing chantier
    xml_content = replace_val(xml_content, 'Laïd ATEK / Yacine EL KHOUMSSI', val(D, 'ing_chan', 'Laïd ATEK / Yacine EL KHOUMSSI'))
    # Date RD0
    xml_content = replace_val(xml_content, '26/03/2026', fmt_date(val(D, 'date_rd0')))
    # Version
    xml_content = replace_val(xml_content, 'Version 1', f'Version {val(D, "version", "1")}')

    # ── Titre opération (gros titre central) ─────────────────────────────────
    xml_content = replace_val(xml_content, 'NSPA-CP3-NA3&amp;4 - NEWADMIN 3 &amp; 4',
                              escape_xml(val(D, 'op_titre', 'NSPA-CP3-NA3&4 - NEWADMIN 3 & 4')))
    # Also non-escaped version
    xml_content = replace_val(xml_content, 'NSPA-CP3-NA3&4 - NEWADMIN 3 & 4',
                              val(D, 'op_titre', 'NSPA-CP3-NA3&4 - NEWADMIN 3 & 4'))

    # ── TITRE I ───────────────────────────────────────────────────────────────
    # MOA
    xml_content = replace_val(xml_content, 'NSPA 11, rue de la Gare L-8325 Capellen',
                              val(D, 'moa', 'NSPA 11, rue de la Gare L-8325 Capellen').replace('\n', ' '))
    # Réf opération
    xml_content = replace_val(xml_content, 'NSPA-CP3-NA3&4 - NEWADMIN 3 & 4 11, rue de la Gare, L-8325 Capellen',
                              val(D, 'op_ref', '').replace('\n', ' '))
    # Descriptif
    xml_content = replace_val(xml_content, 'CONSTRUCTION DES BATIMENTS 3 ET 4',
                              val(D, 'op_desc', 'CONSTRUCTION DES BATIMENTS 3 ET 4'))
    # Contractant
    xml_content = replace_val(xml_content,
                              'Administration des bâtiments publics 10 rue du St Esprit L-1475 Luxembourg',
                              val(D, 'contractant', '').replace('\n', ' '))
    # Signataire
    xml_content = replace_val(xml_content, 'A.Negretti Directrice f.f.',
                              val(D, 'sig_moa', 'A.Negretti Directrice f.f.'))
    # Conception
    xml_content = replace_val(xml_content, '☐ Conception', conc_left)
    xml_content = replace_val(xml_content, '☒ Conception + exécution', conc_right)
    # Missions
    xml_content = replace_val(xml_content, 'DB F', val(D, 'missions', 'DB F'))
    # Date début
    xml_content = replace_val(xml_content, '04/03/2026', fmt_date(val(D, 'date_debut')))
    # Ctrl depuis début
    ctrl_dep = r_val(D, 'ctrl_dep')
    if ctrl_dep == 'OUI':
        xml_content = replace_val(xml_content, '☒ OUI  ☐ NON', '☒ OUI  ☐ NON')
    elif ctrl_dep == 'NON':
        xml_content = xml_content.replace('☒ OUI  ☐ NON', '☐ OUI  ☒ NON', 1)
    # Domaines
    xml_content = replace_val(xml_content, '☒  Sols/Fondations/Structures',
                              f'{sfs}  Sols/Fondations/Structures')
    xml_content = replace_val(xml_content, '☒  Enveloppe : clos/couvert ______',
                              f'{env}  Enveloppe : clos/couvert ______')
    xml_content = replace_val(xml_content, '☒ Etanchéité ______   ☒  Toitures',
                              f'{et} Etanchéité ______   {et}  Toitures')
    xml_content = replace_val(xml_content, '☒ Stabilité                    ☒  Façades',
                              f'{stab} Stabilité                    {ef}  Façades')
    xml_content = replace_val(xml_content, '☒  Sous-sols', f'{ess}  Sous-sols')

    # Participants
    xml_content = replace_val(xml_content,
                              'HO ARCHITECTES sàrl 7, rue de la Toison d\'Or L-2265 Luxembourg',
                              val(D, 'archi_conc', '').replace('\n', ' '))
    xml_content = replace_val(xml_content, 'Non renseigné', val(D, 'be_sol', 'Non renseigné'), )
    xml_content = replace_val(xml_content,
                              'INCA, ingénieurs-conseils sàrl 47, rue Gabriel Lippmann L-6947 Niederanven',
                              val(D, 'be_struct', '').replace('\n', ' '))
    xml_content = replace_val(xml_content,
                              'Génie technique : RMC Consulting, ingénieurs-conseils sàrl 221, route d\'Esch L-1471 Luxembourg',
                              val(D, 'be_autres', '').replace('\n', ' '))
    xml_content = replace_val(xml_content, 'Non renseigné à ce stade.',
                              val(D, 'gros_oeuvre', 'Non renseigné à ce stade.').replace('\n', ' '))

    # Honoraires
    hon = r_val(D, 'hon')
    hon_str = '☒ OUI  ☐ NON' if hon == 'OUI' else '☐ OUI  ☒ NON' if hon == 'NON' else '☐ OUI  ☐ NON'
    # Replace first occurrence of ☐ OUI  ☐ NON (honoraires row)
    xml_content = xml_content.replace('☐ OUI  ☐ NON', hon_str, 1)

    # Montant
    xml_content = replace_val(xml_content, '… HTA', val(D, 'montant', '… HTA'), )

    # Equipements
    eqp = r_val(D, 'eqp')
    eqp_str = '☒ OUI  ☐ NON' if eqp == 'OUI' else '☐ OUI  ☒ NON' if eqp == 'NON' else '☐ OUI  ☐ NON'
    xml_content = xml_content.replace('☐ OUI  ☐ NON', eqp_str, 1)

    # Montant eqp
    xml_content = replace_val(xml_content, '… HTA', val(D, 'montant_eqp', '… HTA'))

    # Nature eqp (the "/" in the cell)
    # Dates
    xml_content = replace_val(xml_content, 'Date du début des travaux (mois/année) : | /',
                              f'Date du début des travaux (mois/année) : | {val(D, "date_trav", "/")}')

    # Commentaires intervenants
    xml_content = replace_val(xml_content,
                              'Aucun commentaire n\'est à formuler sur les références (non communiquées).',
                              val(D, 'comm_interv',
                                  'Aucun commentaire n\'est à formuler sur les références (non communiquées).'))

    # ── TITRE II ─────────────────────────────────────────────────────────────
    xml_content = replace_val(xml_content, '306 m', f'{val(D, "altitude", "306")} m')
    # Zone cyclonique
    cycl = r_val(D, 'cycl')
    cycl_str = '☒ OUI  ☐ NON' if cycl == 'OUI' else '☐ OUI  ☒ NON'
    xml_content = replace_val(xml_content, '☐ OUI  ☒ NON', cycl_str)
    # Vent
    xml_content = replace_val(xml_content, 'vb,0= 24 m/s', f'vb,0= {val(D, "vent", "24")} m/s')
    # Catégorie terrain
    xml_content = replace_val(xml_content, 'Catégorie du terrain : | II',
                              f'Catégorie du terrain : | {val(D, "cat_terrain", "II")}')
    # Règles vent
    xml_content = replace_val(xml_content,
                              'Réglementaire : suivant EN 1991-1-4 (nov. 2005) + EN 1991-1-4:2005/AN-LU: 2011',
                              val(D, 'reg_vent',
                                  'Réglementaire : suivant EN 1991-1-4 (nov. 2005) + EN 1991-1-4:2005/AN-LU: 2011'))
    # Neige
    neige = r_val(D, 'neige')
    neige_str = '☒ OUI  ☐ NON' if neige == 'OUI' else '☐ OUI  ☒ NON'
    xml_content = replace_val(xml_content, '☒ OUI  ☐ NON', neige_str)
    # Valeur neige
    xml_content = replace_val(xml_content, 'sk= 0,73 KN/m²',
                              f'sk= {val(D, "neige_val", "0,73")} KN/m²')
    # Règles neige
    xml_content = replace_val(xml_content,
                              'Réglementaire : suivant EN 1991-1-3 (avril 2004) + EN 1991-1-3:2003/AN-LU: 2011',
                              val(D, 'reg_neige',
                                  'Réglementaire : suivant EN 1991-1-3 (avril 2004) + EN 1991-1-3:2003/AN-LU: 2011'))
    # Inondation précisions
    xml_content = replace_val(xml_content, 'Information non communiquée',
                              val(D, 'inond_prec', 'Information non communiquée'))
    # Nappe
    xml_content = replace_val(xml_content, 'Néant à ce stade',
                              val(D, 'nappe_rc', 'Néant à ce stade'), )
    # Sol assise
    xml_content = replace_val(xml_content, 'Etude non communiquée',
                              val(D, 'rgt_ref', 'Etude non communiquée'))
    # Fondation retenue
    xml_content = replace_val(xml_content,
                              'Selon les coupes, fondations superficielles de type radier.',
                              val(D, 'fond_sol',
                                  'Selon les coupes, fondations superficielles de type radier.'))

    # ── TITRE III ────────────────────────────────────────────────────────────
    # Structures verticales
    xml_content = replace_val(xml_content,
        'Le bâtiment présente une majorité d\'éléments de superstructure en béton armé.',
        val(D, 'struct_vert',
            'Le bâtiment présente une majorité d\'éléments de superstructure en béton armé.'))
    # Façades légères
    xml_content = replace_val(xml_content,
        'Mur-rideau en aluminium.Ces derniers sont filants verticalement devant les dalles et poteaux porteurs.',
        val(D, 'fac_oss', 'Mur-rideau en aluminium.').replace('\n', ' '))
    xml_content = replace_val(xml_content,
        'Remplissage en triple vitrage.Stores à lamelles en aluminium orientables devant les vitrages.',
        val(D, 'fac_rmp', 'Remplissage en triple vitrage.').replace('\n', ' '))
    xml_content = replace_val(xml_content, '&gt; 95% des façades.',
                              escape_xml(val(D, 'fac_surf', '> 95% des façades.')))
    # Planchers
    xml_content = replace_val(xml_content,
        'Les structures horizontales sont réalisées en béton armé en infrastructure',
        val(D, 'planchers',
            'Les structures horizontales sont réalisées en béton armé en infrastructure').split('\n')[0])
    # Toiture
    xml_content = replace_val(xml_content,
        'Dalle en béton armé, coulées sur place, portée standard.Acrotère est reliefs en maçonnerie.',
        val(D, 'toiture',
            'Dalle en béton armé, coulées sur place, portée standard. Acrotère et reliefs en maçonnerie.').replace('\n', ' '))
    # Étanchéité SS
    xml_content = replace_val(xml_content,
        '(*) A ce stade les risques liées à la présence d\'eau (d\'infiltration, nappe, ...) non pas été renseignés.',
        val(D, 'inond_ss_prec',
            '(*) A ce stade les risques liées à la présence d\'eau n\'ont pas été renseignés.').replace('\n', ' '))
    xml_content = replace_val(xml_content,
        'Systèmes d\'étanchéité à base de membranes renforcées bitumineuses, appliquées sur les parois enterrées',
        val(D, 'etan_vert',
            'Systèmes d\'étanchéité à base de membranes renforcées bitumineuses.').split('\n')[0])
    # Imperm façades
    xml_content = replace_val(xml_content,
        'Absence de système d\'imperméabilisation.',
        val(D, 'imp_fac', 'Absence de système d\'imperméabilisation.'))
    # Étanchéité toiture
    xml_content = replace_val(xml_content,
        'La toiture est multifonctionnelle, avec deux rôles principaux :',
        val(D, 'etan_toit',
            'La toiture est multifonctionnelle, avec deux rôles principaux :').split('\n')[0])
    # Couvertures
    xml_content = replace_val(xml_content,
        'Ouvrages auvent :Les éléments sont composés de tôle (cassettes).',
        val(D, 'couvert', 'Ouvrages auvent :').replace('\n', ' '))
    # Carrelages
    xml_content = replace_val(xml_content,
        'Revêtement de sol en moquette en étage courant.',
        val(D, 'carrelage', 'Revêtement de sol en moquette en étage courant.').split('\n')[0])

    # ── TITRE IV ─────────────────────────────────────────────────────────────
    xml_content = replace_val(xml_content,
        'Structures :Longueur du bâtiment ~60 m sans fractionnement.',
        val(D, 'points_att', '').replace('\n', ' '))

    # ── TITRE V / Conclusions ─────────────────────────────────────────────────
    xml_content = replace_val(xml_content, '23 pages',
                              f'{val(D, "nb_pages", "23")} pages')
    xml_content = replace_val(xml_content, "03 pages d'annexes(Liste des postes contrôlées)",
                              f"{val(D, 'nb_ann', '03')} pages d'annexes (Liste des postes contrôlées)")

    # Avis
    avis = val(D, 'avis', '')
    avis_fav  = '☒' if avis == 'Favorable' else '☐'
    avis_res  = '☒' if 'réserves' in avis.lower() else '☐'
    avis_def  = '☒' if avis == 'Défavorable' else '☐'
    xml_content = replace_val(xml_content, '☒ Avis préalable favorable',
                              f'{avis_fav} Avis préalable favorable')
    xml_content = replace_val(xml_content, '☒ Favorable avec réserves',
                              f'{avis_res} Favorable avec réserves')
    xml_content = replace_val(xml_content, '☒ Défavorable',
                              f'{avis_def} Défavorable')

    # Visites
    xml_content = replace_val(xml_content,
        '3/mois  2/mois  1/mois',
        val(D, 'nb_visites', '3/mois  2/mois  1/mois'))

    # ── Signature ────────────────────────────────────────────────────────────
    xml_content = replace_val(xml_content, 'Livange', val(D, 'fait_a', 'Livange'))
    xml_content = replace_val(xml_content, 'Le 26/03/2026',
                              f'Le {fmt_date(val(D, "fait_le"))}')
    xml_content = replace_val(xml_content, 'Laïd ATEK LA',
                              val(D, 'sig_ing', 'Laïd ATEK LA'))

    return xml_content

# ── build output docx ─────────────────────────────────────────────────────────
def generate(D, template_path, output_path):
    import tempfile

    with tempfile.TemporaryDirectory() as tmpdir:
        # Unpack template
        with zipfile.ZipFile(template_path, 'r') as z:
            z.extractall(tmpdir)

        # Read document.xml
        doc_xml_path = os.path.join(tmpdir, 'word', 'document.xml')
        with open(doc_xml_path, 'r', encoding='utf-8') as f:
            xml = f.read()

        # Apply replacements
        xml = process_xml(xml, D)

        # Write back
        with open(doc_xml_path, 'w', encoding='utf-8') as f:
            f.write(xml)

        # Repack as docx
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for root_dir, dirs, files in os.walk(tmpdir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arcname = os.path.relpath(file_path, tmpdir)
                    zout.write(file_path, arcname)

    print(f'✓ Rapport généré : {output_path}')

# ── CLI ───────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python generate_from_template.py dossier.json [template.docx]')
        print('       Le template par défaut est le fichier RD0 original Socotec.')
        sys.exit(1)

    json_path = sys.argv[1]
    template  = sys.argv[2] if len(sys.argv) > 2 else 'RD0_template.docx'

    if not os.path.exists(template):
        print(f'ERREUR: Template introuvable : {template}')
        print('Placez votre document RD0 Word original dans le même dossier sous le nom RD0_template.docx')
        sys.exit(1)

    with open(json_path, encoding='utf-8') as f:
        D = json.load(f)

    ref = D.get('ref', 'dossier').replace('/', '-').replace(' ', '_')
    out = f'RAPSORD_{ref}_RD0.docx'
    generate(D, template, out)
