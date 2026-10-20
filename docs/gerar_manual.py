"""Gera o manual em Word (Manual_Claude_PE.docx) a partir do MANUAL.md, no padrão visual da Dexco.

O texto do manual vive só no MANUAL.md (fonte única). Este script monta um documento de referência
com as fontes da Dexco (Roboto Light no texto, Roboto Black nos títulos) e converte com o pandoc.

Uso (na pasta do projeto):
  python docs/gerar_manual.py [--saida Manual_Claude_PE.docx]
Requer o pandoc instalado (https://pandoc.org; no Windows: winget install pandoc).
"""
import argparse, io, os, shutil, subprocess, sys, tempfile, zipfile
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
w = lambda t: f'{{{W}}}{t}'
AQUI = os.path.dirname(os.path.abspath(__file__))
PROJETO = os.path.dirname(AQUI)
TEXTO, TITULO = 'Roboto Light', 'Roboto Black'


def el(parent, tag, **att):
    e = parent.find(w(tag))
    if e is None:
        e = etree.SubElement(parent, w(tag))
    for k, v in att.items():
        e.set(w(k), v)
    return e


def fonte(rpr, nome, tam=None, cor=None, negrito=None):
    el(rpr, 'rFonts', ascii=nome, hAnsi=nome, cs=nome, eastAsia=nome)
    for k in ('asciiTheme', 'hAnsiTheme', 'cstheme', 'eastAsiaTheme'):
        rpr.find(w('rFonts')).attrib.pop(w(k), None)
    if tam:
        el(rpr, 'sz', val=str(tam)); el(rpr, 'szCs', val=str(tam))
    if cor:
        el(rpr, 'color', val=cor)
    if negrito is not None:
        b = rpr.find(w('b'))
        if negrito is False and b is not None:
            rpr.remove(b)


def estilo(root, sid):
    s = root.find(f'{w("style")}[@{w("styleId")}="{sid}"]')
    if s is None:
        s = etree.SubElement(root, w('style'), {w('type'): 'paragraph', w('styleId'): sid})
    return s


def referencia(destino):
    ref = subprocess.run(['pandoc', '--print-default-data-file', 'reference.docx'], capture_output=True, check=True).stdout
    zin = zipfile.ZipFile(io.BytesIO(ref))
    root = etree.fromstring(zin.read('word/styles.xml'))
    # padrão do documento
    dd = root.find(w('docDefaults'))
    rpr = el(el(dd, 'rPrDefault'), 'rPr')
    fonte(rpr, TEXTO, 20, '262626')
    # títulos
    for sid, tam, cor in (('Title', 44, '000000'), ('Heading1', 30, '000000'), ('Heading2', 24, '262626'), ('Heading3', 21, '404040')):
        s = estilo(root, sid)
        fonte(el(s, 'rPr'), TITULO, tam, cor, negrito=False)
        ppr = el(s, 'pPr')
        if sid == 'Heading1':
            el(ppr, 'spacing', before='360', after='120')
            bd = el(ppr, 'pBdr'); el(bd, 'bottom', val='single', sz='6', space='4', color='BFBFBF')
    for sid, tam, cor in (('Subtitle', 26, '595959'), ('Date', 20, '7F7F7F'), ('Author', 20, '7F7F7F')):
        fonte(el(estilo(root, sid), 'rPr'), TEXTO, tam, cor)
    # prompts (citações) em caixa cinza
    s = estilo(root, 'BlockText')
    ppr = el(s, 'pPr')
    el(ppr, 'shd', val='clear', color='auto', fill='F2F2F2')
    bd = el(ppr, 'pBdr'); el(bd, 'left', val='single', sz='18', space='8', color='595959')
    el(ppr, 'ind', left='284', right='284'); el(ppr, 'spacing', before='60', after='60')
    r = el(s, 'rPr'); fonte(r, TEXTO, 19, '000000')
    i = r.find(w('i'))
    if i is not None:
        r.remove(i)
    # código / árvore de pastas
    fonte(el(estilo(root, 'VerbatimChar'), 'rPr'), 'Consolas', 18, '262626')
    # tabelas: grade leve, cabeçalho cinza
    t = root.find(f'{w("style")}[@{w("styleId")}="Table"]')
    tpr = el(t, 'tblPr')
    bds = el(tpr, 'tblBorders')
    for lado in ('top', 'bottom', 'insideH'):
        el(bds, lado, val='single', sz='4', space='0', color='BFBFBF')
    cm = el(tpr, 'tblCellMar')
    for lado in ('left', 'right'):
        el(cm, lado, w='80', type='dxa')
    for p in t.findall(w('tblStylePr')):
        if p.get(w('type')) == 'firstRow':
            t.remove(p)
    fr = etree.SubElement(t, w('tblStylePr'), {w('type'): 'firstRow'})
    fonte(el(fr, 'rPr'), TITULO, 19, 'FFFFFF')
    el(el(fr, 'tcPr'), 'shd', val='clear', color='auto', fill='262626')
    # grava a referência
    with zipfile.ZipFile(destino, 'w', zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True) if it.filename == 'word/styles.xml' else zin.read(it.filename)
            zout.writestr(it, data)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--saida', default=os.path.join(PROJETO, 'Manual_Claude_PE.docx'))
    args = ap.parse_args()
    if not shutil.which('pandoc'):
        sys.exit('Instale o pandoc (https://pandoc.org) para gerar o Word.')
    tmp = tempfile.mkdtemp(); ref = os.path.join(tmp, 'referencia.docx')
    referencia(ref)
    subprocess.run(['pandoc', os.path.join(PROJETO, 'MANUAL.md'), '-f', 'markdown+task_lists', '-o', args.saida,
                    '--reference-doc', ref], check=True)
    print(args.saida)


if __name__ == '__main__':
    main()
