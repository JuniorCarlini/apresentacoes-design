"""Estilo padrão Saldus para apresentações PowerPoint (.pptx) — python-pptx.

Uso básico:
    from estilo_pptx import *
    prs = nova_apresentacao("Resultados 2026")
    capa(prs, "Resultados de **2026**", "Diretoria · dezembro de 2026")
    divisor(prs, 1, "Onde estamos")
    slide_cartoes(prs, "O ano em números", [{"rotulo": "Receita", "valor": "R$ 4,2 mi", "destaque": True}, ...])
    slide_grafico(prs, "Receita por trimestre", ["T1", "T2", "T3", "T4"], {"Receita": [0.9, 1.0, 1.1, 1.2]})
    encerramento(prs, "Obrigado", ["contato@saldus.com.br"])
    salvar(prs, "resultados.pptx")

Rodar direto gera uma apresentação de exemplo:  python estilo_pptx.py exemplo.pptx
Requer: pip install python-pptx pillow
"""
import io
import math
import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

# ---------------------------------------------------------------- identidade (igual às outras skills)
COR = {
    "primaria": "12424D", "destaque": "FFB745", "branco": "FFFFFF",
    "primaria_clara": "E7EFF0", "zebra": "F5F8F8", "borda": "D5E0E2",
    "texto": "1E2B2E", "texto_suave": "5F7378", "texto_faixa": "B7CDD2",
    "creme": "E3DCCE",  # cor das letras da logo sobre o petróleo
    "positivo": "2E7D32", "negativo": "C62828", "alerta": "B26A00",
    "petroleo_escuro": "0E3640",  # sombra do petróleo, para formas no fundo escuro
}
CORES_GRAFICO = ["12424D", "FFB745", "3E8A96", "8CC2C9", "D9822B", "A7B4B7"]
FONTE = "Poppins"
FONTE_TITULO = "Sora"
FONTE_FINA = "Poppins Light"

ASSETS = Path(__file__).resolve().parent.parent / "assets"
LOGO_CLARA = ASSETS / "logo.png"
LOGO_ESCURA = ASSETS / "logo-escura.png"
SIMBOLO_SVG = ASSETS / "simbolo-s-contorno.svg"
ICONES = ASSETS / "icones"  # ícones Lucide (traço fino, licença ISC): nome do arquivo sem .svg

# 16:9 (13,333 × 7,5 pol.)
LARGURA, ALTURA = Inches(13.333), Inches(7.5)
MARGEM = Inches(0.7)
UTIL = LARGURA - 2 * MARGEM
TOPO_TITULO = Inches(0.55)
TOPO_CONTEUDO = Inches(1.85)
BASE_CONTEUDO = ALTURA - Inches(0.9)


def _rgb(token):
    return RGBColor.from_string(COR.get(token, token))


# ---------------------------------------------------------------- texto
def _fonte_run(run, tamanho, cor="texto", negrito=False, italico=False, fonte=None):
    f = run.font
    f.size = Pt(tamanho)
    f.bold = negrito
    f.italic = italico
    f.color.rgb = _rgb(cor)
    f.name = fonte or FONTE
    rpr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rpr.find(qn(tag))
        if el is None:
            el = rpr.makeelement(qn(tag), {})
            rpr.append(el)
        el.set("typeface", fonte or FONTE)


def _caixa(slide, x, y, w, h, nome=None, ancora=MSO_ANCHOR.TOP, margem=0):
    tb = slide.shapes.add_textbox(x, y, w, h)
    if nome:
        tb.name = nome
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.vertical_anchor = ancora
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = margem
    return tb


def _par(tf, primeiro=False, alinhamento=PP_ALIGN.LEFT, antes=0, depois=0, linha=None):
    p = tf.paragraphs[0] if primeiro else tf.add_paragraph()
    p.alignment = alinhamento
    p.space_before = Pt(antes)
    p.space_after = Pt(depois)
    if linha:
        p.line_spacing = linha
    return p


def _ricos(p, texto, tamanho, cor="texto", cor_negrito="primaria", fonte=None):
    """Texto com **negrito** (sai em `cor_negrito`)."""
    for i, pedaco in enumerate(re.split(r"\*\*", texto)):
        if pedaco:
            r = p.add_run()
            r.text = pedaco
            _fonte_run(r, tamanho, cor_negrito if i % 2 else cor, negrito=bool(i % 2), fonte=fonte)


def _titulo_misto(p, texto, tamanho, cor_fina="branco"):
    """Título em dois pesos (como na capa dos documentos): normal = fino (Poppins Light),
    **destaque** = grosso (Sora) âmbar. Sem marcação: 1ª palavra fina e o resto em destaque."""
    if "**" not in texto:
        partes = texto.split(" ", 1)
        texto = f"{partes[0]} **{partes[1]}**" if len(partes) == 2 else f"**{texto}**"
    for i, pedaco in enumerate(re.split(r"\*\*", texto)):
        if pedaco:
            r = p.add_run()
            r.text = pedaco
            if i % 2:
                _fonte_run(r, tamanho, "destaque", negrito=True, fonte=FONTE_TITULO)
            else:
                _fonte_run(r, tamanho, cor_fina, fonte=FONTE_FINA)


# ---------------------------------------------------------------- medição do texto (fonte real)
# O PowerPoint não avisa quando o texto estoura a caixa. Por isso o script MEDE cada texto com a fonte de
# verdade (Poppins/Sora são largas e têm linhas altas: ~1,4× o tamanho) e ajusta caixas e tamanhos.
_PASTAS_FONTES = [ASSETS / "fontes", Path.home() / "Library/Fonts", Path("/Library/Fonts"),
                  Path("/usr/share/fonts/truetype"), Path("/usr/share/fonts"), Path("C:/Windows/Fonts"),
                  Path.home() / "AppData/Local/Microsoft/Windows/Fonts"]
_ARQUIVOS = {("Poppins", False): ["Poppins-Regular.ttf"], ("Poppins", True): ["Poppins-Bold.ttf"],
             ("Poppins Light", False): ["Poppins-Light.ttf"], ("Poppins Light", True): ["Poppins-Light.ttf"],
             ("Sora", False): ["Sora[wght].ttf", "Sora-Regular.ttf", "Sora-VariableFont_wght.ttf"],
             ("Sora", True): ["Sora-Bold.ttf", "Sora[wght].ttf", "Sora-VariableFont_wght.ttf"]}
_CACHE = {}


def _fonte_medida(familia, negrito):
    """Fonte PIL (tamanho 100) para medir; None se a fonte não estiver instalada (usa estimativa)."""
    chave = (familia, bool(negrito))
    if chave in _CACHE:
        return _CACHE[chave]
    from PIL import ImageFont
    achada = None
    for nome in _ARQUIVOS.get(chave, []):
        for pasta in _PASTAS_FONTES:
            for arq in (list(pasta.rglob(nome)) if pasta.exists() and pasta.name == "truetype" else [pasta / nome]):
                if arq.exists():
                    try:
                        achada = ImageFont.truetype(str(arq), 100)
                        if "wght" in arq.name:
                            try:
                                achada.set_variation_by_name("Bold" if negrito else "Regular")
                            except Exception:
                                pass
                    except Exception:
                        achada = None
                if achada:
                    break
            if achada:
                break
        if achada:
            break
    _CACHE[chave] = achada
    return achada


def _largura_pt(texto, familia, tamanho, negrito):
    f = _fonte_medida(familia, negrito)
    if f is None:
        return len(texto) * tamanho * (0.62 if negrito else 0.58)
    return f.getlength(texto) * tamanho / 100


def _linha_pt(familia, tamanho, negrito=False):
    """Altura natural de uma linha (o que o PowerPoint usa com entrelinha 1,0)."""
    f = _fonte_medida(familia, negrito)
    if f is None:
        return tamanho * 1.4
    asc, desc = f.getmetrics()
    return (asc + desc) * tamanho / 100


def _segmentos(texto, tamanho, fonte=None, fonte_negrito=None, negrito_tudo=False):
    """Quebra "**negrito**" em pedaços (texto, família, tamanho, negrito) para medir."""
    fonte = fonte or FONTE
    out = []
    for i, pedaco in enumerate(re.split(r"\*\*", texto)):
        if pedaco:
            neg = negrito_tudo or bool(i % 2)
            out.append((pedaco, fonte_negrito if (neg and fonte_negrito) else fonte, tamanho, neg))
    return out


def _altura_pt(segs, largura_pt, linha=1.0):
    """Quantas linhas os pedaços ocupam na largura dada (quebra por palavra) e a altura total em pt."""
    if not segs:
        return 0, 0
    palavras = []
    for txt, fam, tam, neg in segs:
        for parte in re.findall(r"\S+\s*", txt):
            palavras.append((parte, fam, tam, neg))
    linhas, atual = 1, 0.0
    for parte, fam, tam, neg in palavras:
        w = _largura_pt(parte, fam, tam, neg)
        w_sem_espaco = _largura_pt(parte.rstrip(), fam, tam, neg)
        if atual > 0 and atual + w_sem_espaco > largura_pt:
            linhas += 1
            atual = w
        else:
            atual += w
    alt_linha = max(_linha_pt(fam, tam, neg) for _, fam, tam, neg in segs)
    return linhas, linhas * alt_linha * linha


def _altura_texto(texto, tamanho, largura_emu, linha=1.15, fonte=None, negrito=False, prefixo=None):
    segs = ([(prefixo, FONTE, tamanho, False)] if prefixo else []) + _segmentos(texto, tamanho, fonte,
                                                                              negrito_tudo=negrito)
    return Emu(int(_altura_pt(segs, largura_emu / 12700, linha)[1] * 12700))


# ---------------------------------------------------------------- formas
def _sem_sombra(forma):
    """Tira a sombra do tema (efeito padrão do PowerPoint nas formas)."""
    sppr = forma._element.spPr
    for el in sppr.findall(qn("a:effectLst")):
        sppr.remove(el)
    sppr.append(sppr.makeelement(qn("a:effectLst"), {}))
    estilo = forma._element.find(qn("p:style"))
    if estilo is not None:  # a sombra também vem do estilo do tema (effectRef)
        estilo.find(qn("a:effectRef")).set("idx", "0")


def _ret(slide, x, y, w, h, cor, nome=None, arredondado=False, raio=0.08):
    forma = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if arredondado else MSO_SHAPE.RECTANGLE,
                                   x, y, w, h)
    forma.fill.solid()
    forma.fill.fore_color.rgb = _rgb(cor)
    forma.line.fill.background()
    _sem_sombra(forma)
    if arredondado:
        forma.adjustments[0] = raio
    if nome:
        forma.name = nome
    return forma


def _circulo(slide, x, y, d, cor, nome=None):
    c = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, d, d)
    c.fill.solid()
    c.fill.fore_color.rgb = _rgb(cor)
    c.line.fill.background()
    _sem_sombra(c)
    if nome:
        c.name = nome
    return c


def _fundo(slide, cor):
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = _rgb(cor)


def _svg_contorno_png(svg, cor, escala=0.5, traco=18, alfa=255):
    """Contorno do "S" da marca (SVG simples, M/L/C/Z) em PNG transparente, só com a Pillow."""
    from PIL import Image, ImageChops, ImageDraw
    xml = Path(svg).read_text(encoding="utf-8")
    vb = re.search(r'viewBox="[\d.\s-]*?\s([\d.]+)\s([\d.]+)"', xml)
    w, h = float(vb.group(1)), float(vb.group(2))
    W, H = int(w * escala), int(h * escala)
    pre, con = Image.new("L", (W, H), 0), Image.new("L", (W, H), 0)
    dp, dc = ImageDraw.Draw(pre), ImageDraw.Draw(con)
    for d in re.findall(r'<path[^>]*\sd="([^"]+)"', xml):
        tok = re.findall(r"[MLCZ]|-?\d*\.?\d+(?:e-?\d+)?", d)
        figs, pts, i, cmd, at = [], [], 0, None, (0.0, 0.0)
        while i < len(tok):
            t = tok[i]
            if t in "MLCZ":
                cmd = t
                i += 1
                if t == "Z" and pts:
                    figs.append(pts)
                    pts = []
                continue
            if cmd == "M":
                if pts:
                    figs.append(pts)
                at = (float(tok[i]), float(tok[i + 1]))
                pts, i, cmd = [at], i + 2, "L"
            elif cmd == "L":
                at = (float(tok[i]), float(tok[i + 1]))
                pts.append(at)
                i += 2
            elif cmd == "C":
                c = [float(v) for v in tok[i:i + 6]]
                p0, p1, p2, p3 = at, (c[0], c[1]), (c[2], c[3]), (c[4], c[5])
                for k in range(1, 13):
                    u = k / 12
                    a, b, cc, dd = (1 - u) ** 3, 3 * u * (1 - u) ** 2, 3 * u * u * (1 - u), u ** 3
                    pts.append((a * p0[0] + b * p1[0] + cc * p2[0] + dd * p3[0],
                                a * p0[1] + b * p1[1] + cc * p2[1] + dd * p3[1]))
                at, i = p3, i + 6
            else:
                i += 1
        if pts:
            figs.append(pts)
        for f in figs:
            f = [(x * escala, y * escala) for x, y in f]
            dp.polygon(f, fill=255)
            dc.line(f + [f[0], f[1]], fill=255, width=max(2, int(traco * escala)), joint="curve")
    alfa_img = ImageChops.multiply(pre, con)
    if alfa < 255:
        alfa_img = alfa_img.point(lambda v: v * alfa // 255)
    img = Image.new("RGBA", (W, H), "#" + COR.get(cor, cor))
    img.putalpha(alfa_img)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf, w / h


def _simbolo_fundo(slide, cor="texto_faixa", alfa=40, altura=Inches(9.0), x=None, y=None):
    """"S" grande em contorno, vazando pela direita do slide (textura da marca, como no manual)."""
    if not SIMBOLO_SVG.exists():
        return
    png, prop = _svg_contorno_png(SIMBOLO_SVG, cor, alfa=alfa, traco=14)
    larg = int(altura * prop)
    x = x if x is not None else LARGURA - int(larg * 0.72)
    y = y if y is not None else int((ALTURA - altura) / 2)
    pic = slide.shapes.add_picture(png, x, y, larg, altura)
    pic.name = "Símbolo S (fundo)"
    # manda para trás de tudo
    sp = pic._element
    sp.getparent().remove(sp)
    slide.shapes._spTree.insert(2, sp)
    return pic


# ---------------------------------------------------------------- ícones
def icones_disponiveis():
    return sorted(p.stem for p in ICONES.glob("*.svg"))


def _icone_png(nome, cor):
    """Ícone Lucide (SVG de traço) recolorido e convertido em PNG transparente de 512 px (PyMuPDF)."""
    import fitz  # pip install pymupdf
    arq = ICONES / f"{nome}.svg"
    if not arq.exists():
        raise ValueError(f'ícone "{nome}" não existe. Disponíveis: {", ".join(icones_disponiveis())}')
    svg = arq.read_text(encoding="utf-8").replace("currentColor", "#" + COR.get(cor, cor))
    d = fitz.open(stream=svg.encode(), filetype="svg")
    pix = d[0].get_pixmap(matrix=fitz.Matrix(512 / d[0].rect.width, 512 / d[0].rect.width), alpha=True)
    return io.BytesIO(pix.tobytes("png"))


def icone(slide, nome, x, y, diametro=Inches(0.62), escuro=False, sobre_claro=False):
    """Ícone dentro de um círculo (o motivo visual da marca, como os números dos passos).
    Fundo claro: círculo petróleo claro + ícone petróleo. Escuro/destaque: círculo petróleo + ícone âmbar."""
    fundo = "petroleo_escuro" if escuro else ("branco" if sobre_claro else "primaria_clara")
    _circulo(slide, x, y, diametro, fundo, f"Círculo do ícone {nome}")
    lado = int(diametro * 0.52)
    pic = slide.shapes.add_picture(_icone_png(nome, "destaque" if escuro else "primaria"),
                                   x + (diametro - lado) // 2, y + (diametro - lado) // 2, lado, lado)
    pic.name = f"Ícone {nome}"
    return pic


# ---------------------------------------------------------------- apresentação e moldura
def nova_apresentacao(titulo="Apresentação Saldus", autor="Saldus"):
    """Apresentação 16:9 com as fontes da marca no tema (títulos Sora, texto Poppins)."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = LARGURA, ALTURA
    prs.core_properties.title = titulo
    prs.core_properties.author = autor
    _tema(prs)
    prs._saldus = {"titulo": titulo, "n": 0}
    return prs


def _tema(prs):
    """Grava cores e fontes da marca no tema (o que o PowerPoint usa em caixas novas e gráficos)."""
    from lxml import etree
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    master = prs.slide_masters[0]
    tema = master.part.part_related_by(RT.THEME)
    raiz = etree.fromstring(tema.blob)
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    cores = {"dk1": COR["texto"], "lt1": "FFFFFF", "dk2": COR["primaria"], "lt2": COR["primaria_clara"],
             "accent1": CORES_GRAFICO[0], "accent2": CORES_GRAFICO[1], "accent3": CORES_GRAFICO[2],
             "accent4": CORES_GRAFICO[3], "accent5": CORES_GRAFICO[4], "accent6": CORES_GRAFICO[5],
             "hlink": COR["primaria"], "folHlink": COR["texto_suave"]}
    esquema = raiz.find(".//a:clrScheme", ns)
    esquema.set("name", "Saldus")
    for nome, hexa in cores.items():
        el = esquema.find(f"a:{nome}", ns)
        for filho in list(el):
            el.remove(filho)
        etree.SubElement(el, "{%s}srgbClr" % ns["a"]).set("val", hexa)
    fontes = raiz.find(".//a:fontScheme", ns)
    fontes.set("name", "Saldus")
    fontes.find("a:majorFont/a:latin", ns).set("typeface", FONTE_TITULO)
    fontes.find("a:minorFont/a:latin", ns).set("typeface", FONTE)
    tema._blob = etree.tostring(raiz, xml_declaration=True, encoding="UTF-8", standalone=True)


def _novo_slide(prs, escuro=False):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # em branco
    _fundo(slide, "primaria" if escuro else "branco")
    prs._saldus["n"] += 1
    return slide


def _rodape(prs, slide, escuro=False):
    """Rodapé discreto em todo slide de conteúdo: logo pequena à esquerda e número à direita."""
    logo = LOGO_CLARA if escuro else LOGO_ESCURA
    if logo.exists():
        pic = slide.shapes.add_picture(str(logo), MARGEM, ALTURA - Inches(0.56), height=Inches(0.34))
        pic.name = "Logo (rodapé)"
    tb = _caixa(slide, LARGURA - MARGEM - Inches(1.5), ALTURA - Inches(0.52), Inches(1.5), Inches(0.3),
                "Número do slide", MSO_ANCHOR.MIDDLE)
    p = _par(tb.text_frame, True, PP_ALIGN.RIGHT)
    r = p.add_run()
    r.text = str(prs._saldus["n"])
    _fonte_run(r, 11, "texto_faixa" if escuro else "texto_suave")


def _titulo_slide(slide, titulo, sobretitulo=None, escuro=False, subtitulo=None):
    """Título de slide de conteúdo: sempre no mesmo lugar, Sora 30, petróleo (branco no escuro).
    sobretitulo: rótulo curto em MAIÚSCULAS acima do título (o assunto/seção).
    subtitulo: uma frase curta logo abaixo do título (cinza, 16 pt)."""
    y = TOPO_TITULO
    if sobretitulo:
        tb = _caixa(slide, MARGEM, y, UTIL, Inches(0.3), "Sobretítulo")
        r = _par(tb.text_frame, True).add_run()
        r.text = sobretitulo.upper()
        _fonte_run(r, 12, "destaque" if escuro else "texto_suave", negrito=True)
        y += Inches(0.36)
    tb = _caixa(slide, MARGEM, y, UTIL, Inches(0.9), "Título", MSO_ANCHOR.TOP)
    p = _par(tb.text_frame, True, linha=1.0)
    _ricos(p, titulo, 30, "branco" if escuro else "primaria",
           cor_negrito="destaque" if escuro else "primaria", fonte=FONTE_TITULO)
    for r in p.runs:
        r.font.bold = True
    if subtitulo:
        tb = _caixa(slide, MARGEM, y + Inches(0.6), UTIL, Inches(0.4), "Subtítulo do slide")
        r = _par(tb.text_frame, True).add_run()
        r.text = subtitulo
        _fonte_run(r, 16, "texto_faixa" if escuro else "texto_suave")


def _painel(s, x, y, w, h, painel):
    """Painel ao lado de tabela/conteúdo. painel: {"titulo", "texto", "valor" (opcional, número grande),
    "destaque": True (fundo petróleo, valor âmbar), "icone" (opcional)}."""
    forte = painel.get("destaque")
    _ret(s, x, y, w, h, "primaria" if forte else "primaria_clara", "Painel", True, 0.06)
    topo = Inches(0.4)
    if painel.get("icone"):
        icone(s, painel["icone"], x + Inches(0.4), y + Inches(0.4), Inches(0.62), escuro=forte, sobre_claro=not forte)
        topo = Inches(1.2)
    tb = _caixa(s, x + Inches(0.4), y + topo, w - Inches(0.8), h - topo - Inches(0.35), "Texto do painel")
    tf = tb.text_frame
    primeiro = True
    if painel.get("titulo"):
        r = _par(tf, True, depois=8).add_run()
        r.text = painel["titulo"]
        _fonte_run(r, 19, "destaque" if forte else "primaria", negrito=True, fonte=FONTE_TITULO)
        primeiro = False
    if painel.get("valor"):
        r = _par(tf, primeiro, depois=6, linha=1.0).add_run()
        r.text = painel["valor"]
        larg_pol = (w - Inches(0.8)) / 914400  # número sempre numa linha só
        tam = max(24, min(40, int(larg_pol * 72 / (max(len(r.text), 1) * 0.62))))
        _fonte_run(r, tam, "destaque" if forte else "primaria", negrito=True, fonte=FONTE_TITULO)
        primeiro = False
    blocos = [b for b in (painel.get("texto") or "").split("\n") if b.strip()]
    usado = (Pt(19 * 1.4 + 8) if painel.get("titulo") else 0) + (Pt(40 * 1.4 + 6) if painel.get("valor") else 0)
    livre = h - topo - Inches(0.35) - usado
    for tam_t in (15, 14, 13, 12):
        if sum(_altura_texto(b, tam_t, w - Inches(0.8)) + Pt(8) for b in blocos) <= livre:
            break
    for bloco in blocos:
        p = _par(tf, primeiro, depois=8, linha=1.15)
        primeiro = False
        _ricos(p, bloco, tam_t, "creme" if forte else "texto", cor_negrito="destaque" if forte else "primaria")


def _nota(s, texto, y):
    tb = _caixa(s, MARGEM, y, UTIL, Inches(0.5), "Nota")
    _ricos(_par(tb.text_frame, True), texto, 15, "texto_suave")


# ---------------------------------------------------------------- tipos de slide
def capa(prs, titulo, subtitulo=None, rodape_texto=None, foto=None):
    """Capa: petróleo inteiro, logo clara no alto, título em dois pesos (fino branco + destaque âmbar),
    subtítulo creme, "S" da marca em contorno vazando à direita.
    foto: imagem principal ocupando a metade direita do slide (cortada sem distorcer)."""
    s = _novo_slide(prs, escuro=True)
    larg_texto = UTIL
    if foto:
        meia = int(LARGURA * 0.46)
        png = _recortar(foto, meia, ALTURA)
        pic = s.shapes.add_picture(png, LARGURA - meia, 0, meia, ALTURA)
        pic.name = "Foto da capa"
        larg_texto = LARGURA - meia - MARGEM - Inches(0.5)
    else:
        _simbolo_fundo(s, "texto_faixa", alfa=38)
    if LOGO_CLARA.exists():
        s.shapes.add_picture(str(LOGO_CLARA), MARGEM, Inches(0.65), height=Inches(0.45)).name = "Logo"
    # título + subtítulo numa caixa só, centralizada na altura do slide
    tb = _caixa(s, MARGEM, Inches(1.4), larg_texto, ALTURA - Inches(2.8), "Título", MSO_ANCHOR.MIDDLE)
    _titulo_misto(_par(tb.text_frame, True, linha=0.95), titulo, 54)
    if subtitulo:
        r = _par(tb.text_frame, antes=14).add_run()
        r.text = subtitulo
        _fonte_run(r, 18, "creme")
    if rodape_texto:
        tb = _caixa(s, MARGEM, ALTURA - Inches(0.75), larg_texto, Inches(0.3), "Rodapé da capa")
        r = _par(tb.text_frame, True).add_run()
        r.text = rodape_texto
        _fonte_run(r, 11, "texto_faixa")
    return s


def divisor(prs, numero, titulo, subtitulo=None):
    """Divisor de seção: petróleo, número grande âmbar ("01"), título branco, "S" ao fundo."""
    s = _novo_slide(prs, escuro=True)
    _simbolo_fundo(s, "texto_faixa", alfa=30)
    tb = _caixa(s, MARGEM, Inches(2.0), Inches(4), Inches(1.6), "Número da seção", MSO_ANCHOR.BOTTOM)
    r = _par(tb.text_frame, True, linha=0.9).add_run()
    r.text = f"{numero:02d}" if isinstance(numero, int) else str(numero)
    _fonte_run(r, 96, "destaque", negrito=True, fonte=FONTE_TITULO)
    # título e subtítulo na mesma caixa: o subtítulo vem logo abaixo do título
    tb = _caixa(s, MARGEM, Inches(3.75), Inches(8.5), Inches(2.2), "Título da seção")
    _titulo_misto(_par(tb.text_frame, True, linha=1.0), titulo, 40)
    if subtitulo:
        r = _par(tb.text_frame, antes=10).add_run()
        r.text = subtitulo
        _fonte_run(r, 18, "creme")
    _rodape(prs, s, escuro=True)
    return s


def slide_texto(prs, titulo, itens, sobretitulo=None, subtitulo=None, imagem=None, destaque=None):
    """Título + lista curta (marcadores âmbar). imagem: ocupa a coluna da direita (cantos arredondados
    não existem em foto no PowerPoint, então vai reta). destaque: frase grande à direita, se não houver
    imagem (ex.: o número ou a ideia principal)."""
    s = _novo_slide(prs)
    _titulo_slide(s, titulo, sobretitulo, subtitulo=subtitulo)
    col = int(UTIL * (0.52 if (imagem or destaque) else 1))
    disp = BASE_CONTEUDO - TOPO_CONTEUDO - (Inches(0.55) if subtitulo else 0)
    topo = TOPO_CONTEUDO + (Inches(0.35) if subtitulo else 0)
    com_icones = any(isinstance(i, (tuple, list)) for i in itens)
    if com_icones:  # linhas "ícone + texto": (nome_do_ícone, "texto"), alturas medidas com a fonte real
        d, gap = Inches(0.62), Inches(0.3)
        larg_txt = col - d - Inches(0.3)
        pares = [i if isinstance(i, (tuple, list)) else (None, i) for i in itens]
        for tam in (18, 17, 16, 15, 14):
            alturas = [max(d, _altura_texto(t, tam, larg_txt) + Inches(0.12)) for _, t in pares]
            if sum(alturas) + gap * (len(pares) - 1) <= disp:
                break
        y = topo
        for i, ((nome, txt), h) in enumerate(zip(pares, alturas)):
            if nome:
                icone(s, nome, MARGEM, y, d)
            tb = _caixa(s, MARGEM + d + Inches(0.3), y, larg_txt, h, f"Texto {i + 1}", MSO_ANCHOR.TOP)
            tb.text_frame.margin_top = Inches(0.1)
            _ricos(_par(tb.text_frame, True, linha=1.15), txt, tam)
            y += h + gap
    else:
        for tam in (18, 17, 16, 15, 14):
            total = sum(_altura_texto(t, tam, col, prefixo="●  ") + Pt(12) for t in itens)
            if total <= disp:
                break
        tb = _caixa(s, MARGEM, topo, col, disp, "Texto")
        for i, item in enumerate(itens):
            p = _par(tb.text_frame, i == 0, depois=12, linha=1.15)
            r = p.add_run()
            r.text = "●  "
            _fonte_run(r, 12, "destaque")
            _ricos(p, item, tam)
    xd = MARGEM + col + Inches(0.5)
    wd = LARGURA - MARGEM - xd
    if imagem:
        png = _recortar(imagem, wd, BASE_CONTEUDO - TOPO_CONTEUDO)
        s.shapes.add_picture(png, xd, TOPO_CONTEUDO, wd, BASE_CONTEUDO - TOPO_CONTEUDO).name = "Imagem"
    elif destaque:
        caixa = _ret(s, xd, topo, wd, disp, "primaria_clara", "Destaque", arredondado=True, raio=0.06)
        tf = caixa.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = Inches(0.45)
        for tam_d in (26, 24, 22, 20, 18):
            if _altura_texto(destaque, tam_d, wd - Inches(0.9), 1.1, FONTE_TITULO, True) <= disp - Inches(0.8):
                break
        _ricos(_par(tf, True, linha=1.1), destaque, tam_d, "primaria", cor_negrito="primaria", fonte=FONTE_TITULO)
        for r in tf.paragraphs[0].runs:
            r.font.bold = True
    _rodape(prs, s)
    return s


def slide_cartoes(prs, titulo, cartoes, sobretitulo=None, subtitulo=None, nota=None):
    """Números grandes em cartões (2 a 4). Só UM cartão com destaque=True (fundo petróleo, número âmbar)."""
    s = _novo_slide(prs)
    _titulo_slide(s, titulo, sobretitulo, subtitulo=subtitulo)
    n = len(cartoes)
    gap = Inches(0.35)
    w = int((UTIL - gap * (n - 1)) / n)
    tem_icone = any(k.get("icone") for k in cartoes)
    larg_txt = int((UTIL - Inches(0.35) * (len(cartoes) - 1)) / len(cartoes)) - Inches(0.7)
    maior_det = max((_altura_texto(k.get("detalhe", ""), 15, larg_txt) for k in cartoes if k.get("detalhe")),
                    default=0)
    conteudo = Pt(13 * 1.4 + 6) + Pt(44 * 1.4 + 8) + maior_det
    h = max(Inches(2.5), (Inches(1.1) if tem_icone else Inches(0.3)) + conteudo + Inches(0.4))
    y = TOPO_CONTEUDO + Inches(0.35)
    for i, k in enumerate(cartoes):
        x = MARGEM + i * (w + gap)
        forte = k.get("destaque")
        _ret(s, x, y, w, h, "primaria" if forte else "primaria_clara", f"Cartão {i + 1}", True, 0.06)
        topo_txt = Inches(0.3)
        if k.get("icone"):
            icone(s, k["icone"], x + Inches(0.35), y + Inches(0.35), Inches(0.62), escuro=forte, sobre_claro=not forte)
            topo_txt = Inches(1.1)
        tb = _caixa(s, x + Inches(0.35), y + topo_txt, w - Inches(0.7), h - topo_txt - Inches(0.3),
                    f"Texto cartão {i + 1}", MSO_ANCHOR.TOP if tem_icone else MSO_ANCHOR.MIDDLE)
        tf = tb.text_frame
        r = _par(tf, True, depois=6).add_run()
        r.text = k["rotulo"].upper()
        _fonte_run(r, 13, "texto_faixa" if forte else "texto_suave", negrito=True)
        r = _par(tf, linha=1.0, depois=8).add_run()
        r.text = str(k["valor"])
        # número grande, mas sempre numa linha só: diminui conforme o tamanho do texto e do cartão
        larg_pol = (w - Inches(0.7)) / 914400
        tam = min(44 if n <= 3 else 36, int(larg_pol * 72 / (max(len(r.text), 1) * 0.62)))
        _fonte_run(r, max(tam, 24), "destaque" if forte else "primaria", negrito=True, fonte=FONTE_TITULO)
        if k.get("detalhe"):
            r = _par(tf).add_run()
            r.text = k["detalhe"]
            _fonte_run(r, 15, "creme" if forte else "texto_suave")
    if nota:
        tb = _caixa(s, MARGEM, y + h + Inches(0.3), UTIL, Inches(0.5), "Nota")
        _ricos(_par(tb.text_frame, True), nota, 14, "texto_suave")
    _rodape(prs, s)
    return s


def slide_grafico(prs, titulo, categorias, series, sobretitulo=None, subtitulo=None, tipo="colunas", formato='#,##0',
                  comentario=None, fonte_dados=None):
    """Gráfico nativo do PowerPoint (editável), no padrão: sem grade, sem eixo de valores, rótulo do
    valor na barra, cores da marca. tipo: "colunas", "barras" (horizontais) ou "linhas".
    series: {"Nome": [valores]}. comentario: lista de 1–3 frases curtas à direita do gráfico."""
    s = _novo_slide(prs)
    _titulo_slide(s, titulo, sobretitulo, subtitulo=subtitulo)
    dados = CategoryChartData()
    dados.categories = categorias
    for nome, valores in series.items():
        dados.add_series(nome, valores)
    tipos = {"colunas": XL_CHART_TYPE.COLUMN_CLUSTERED, "barras": XL_CHART_TYPE.BAR_CLUSTERED,
             "linhas": XL_CHART_TYPE.LINE_MARKERS}
    larg = int(UTIL * (0.64 if comentario else 1))
    alt = BASE_CONTEUDO - TOPO_CONTEUDO - (Inches(0.35) if fonte_dados else 0)
    gf = s.shapes.add_chart(tipos[tipo], MARGEM, TOPO_CONTEUDO, larg, alt, dados)
    gf.name = "Gráfico"
    g = gf.chart
    g.font.name = FONTE
    g.font.size = Pt(13)
    g.font.color.rgb = _rgb("texto")
    g.has_title = False
    g.has_legend = len(series) > 1
    if g.has_legend:
        from pptx.enum.chart import XL_LEGEND_POSITION
        g.legend.position = XL_LEGEND_POSITION.TOP
        g.legend.include_in_layout = False
    va = g.value_axis
    va.has_major_gridlines = False
    va.visible = False
    ca = g.category_axis
    ca.format.line.color.rgb = _rgb("borda")
    ca.has_major_gridlines = False
    ca.tick_labels.font.color.rgb = _rgb("texto_suave")
    if tipo == "barras":
        ca.reverse_order = True  # barras na mesma ordem da lista (a 1ª em cima)
    pl = g.plots[0]
    if tipo != "linhas":
        pl.gap_width = 60
        pl.overlap = -10 if len(series) > 1 else 0
    pl.has_data_labels = True
    dl = pl.data_labels
    dl.number_format = formato
    dl.number_format_is_linked = False
    dl.font.size = Pt(13)
    dl.font.bold = True
    dl.font.color.rgb = _rgb("texto")
    dl.position = XL_LABEL_POSITION.OUTSIDE_END if tipo != "linhas" else XL_LABEL_POSITION.ABOVE
    for i, serie in enumerate(pl.series):
        cor = _rgb(CORES_GRAFICO[i % len(CORES_GRAFICO)])
        if tipo == "linhas":
            serie.format.line.color.rgb = cor
            serie.format.line.width = Pt(3)
            serie.smooth = False
        else:
            serie.format.fill.solid()
            serie.format.fill.fore_color.rgb = cor
    if fonte_dados:
        tb = _caixa(s, MARGEM, TOPO_CONTEUDO + alt + Inches(0.05), larg, Inches(0.3), "Fonte dos dados")
        r = _par(tb.text_frame, True).add_run()
        r.text = "Fonte: " + fonte_dados
        _fonte_run(r, 11, "texto_suave")
    if comentario:
        xc = MARGEM + larg + Inches(0.5)
        wc = LARGURA - MARGEM - xc
        for tam_c in (17, 16, 15, 14):
            if sum(_altura_texto(f, tam_c, wc) + Pt(16) for f in comentario) <= alt - Inches(0.3):
                break
        tb = _caixa(s, xc, TOPO_CONTEUDO + Inches(0.3), wc, alt - Inches(0.3), "Comentário")
        for i, frase in enumerate(comentario):
            p = _par(tb.text_frame, i == 0, depois=16, linha=1.15)
            _ricos(p, frase, tam_c)
    _rodape(prs, s)
    return s


def slide_tabela(prs, titulo, cabecalho, linhas, sobretitulo=None, subtitulo=None, alinhar_direita=None, total=None,
                 fonte_dados=None, painel=None, nota=None, largura_tabela=0.58):
    """Tabela no padrão da planilha: cabeçalho petróleo com texto branco, zebra, total em petróleo
    claro e negrito. alinhar_direita: nomes das colunas numéricas. Valores já formatados (texto).
    Até ~8 linhas por slide; mais que isso, divida em dois slides."""
    s = _novo_slide(prs)
    _titulo_slide(s, titulo, sobretitulo, subtitulo=subtitulo)
    alinhar_direita = set(alinhar_direita or [])
    n_lin = len(linhas) + 1 + (1 if total else 0)
    alt_lin = Inches(0.48)
    larg_tab = int(UTIL * largura_tabela) if painel else UTIL  # fração da largura quando há painel
    gf = s.shapes.add_table(n_lin, len(cabecalho), MARGEM, TOPO_CONTEUDO, larg_tab, alt_lin * n_lin)
    gf.name = "Tabela"
    t = gf.table
    # largura das colunas proporcional ao texto mais longo de cada uma (nada de colunas iguais)
    todas = [cabecalho] + [list(l) for l in linhas] + ([list(total)] if total else [])
    # +4 de folga por coluna (margens da célula) e mínimo de 10 caracteres, para valores curtos não quebrarem
    pesos = [min(max(max(len(str(l[j])) for l in todas if j < len(l)), 10), 45) + 4 for j in range(len(cabecalho))]
    soma = sum(pesos)
    for j, peso in enumerate(pesos):
        t.columns[j].width = int(larg_tab * peso / soma)
    tblpr = gf._element.graphic.graphicData.tbl.tblPr
    tblpr.set("firstRow", "0")
    tblpr.set("bandRow", "0")
    estilo = tblpr.find(qn("a:tableStyleId"))
    if estilo is not None:
        estilo.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"  # "Sem estilo, sem grade"

    def celula(c, texto, cor_fundo, cor_texto, negrito=False, direita=False, fonte=None):
        c.fill.solid()
        c.fill.fore_color.rgb = _rgb(cor_fundo)
        c.margin_left = c.margin_right = Inches(0.15)
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = c.text_frame
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT if direita else PP_ALIGN.LEFT
        r = p.add_run()
        r.text = str(texto)
        _fonte_run(r, 15, cor_texto, negrito=negrito, fonte=fonte)
        # sem bordas (a zebra separa as linhas). As bordas têm de vir ANTES do preenchimento e na ordem
        # lnL, lnR, lnT, lnB — fora dessa ordem o PowerPoint ignora e desenha a grade.
        tcpr = c._tc.get_or_add_tcPr()
        for i, lado in enumerate(("a:lnL", "a:lnR", "a:lnT", "a:lnB")):
            ln = tcpr.makeelement(qn(lado), {"w": "0"})
            ln.append(ln.makeelement(qn("a:noFill"), {}))
            tcpr.insert(i, ln)

    for j, nome in enumerate(cabecalho):
        celula(t.cell(0, j), nome, "primaria", "branco", True, nome in alinhar_direita, FONTE_TITULO)
    for i, linha in enumerate(linhas, 1):
        for j, v in enumerate(linha):
            celula(t.cell(i, j), v, "zebra" if i % 2 == 0 else "branco", "texto",
                   direita=cabecalho[j] in alinhar_direita)
    if total:
        for j, v in enumerate(total):
            celula(t.cell(n_lin - 1, j), v, "primaria_clara", "primaria", True, cabecalho[j] in alinhar_direita)
    for i in range(n_lin):
        t.rows[i].height = alt_lin
    fim_tab = TOPO_CONTEUDO + alt_lin * n_lin
    if fonte_dados:
        tb = _caixa(s, MARGEM, fim_tab + Inches(0.15), larg_tab, Inches(0.3), "Fonte dos dados")
        r = _par(tb.text_frame, True).add_run()
        r.text = "Fonte: " + fonte_dados
        _fonte_run(r, 11, "texto_suave")
        fim_tab += Inches(0.4)
    if painel:
        # o painel ocupa a altura toda disponível (o texto nunca vaza); a nota fica no pé
        xp = MARGEM + larg_tab + Inches(0.45)
        alt_p = BASE_CONTEUDO - TOPO_CONTEUDO - (Inches(0.7) if nota else 0)
        _painel(s, xp, TOPO_CONTEUDO, LARGURA - MARGEM - xp, alt_p, painel)
        fim_tab = max(fim_tab, TOPO_CONTEUDO + alt_p)
    if nota:
        _nota(s, nota, fim_tab + Inches(0.2))
    _rodape(prs, s)
    return s


def slide_comparacao(prs, titulo, esquerda, direita, sobretitulo=None, subtitulo=None):
    """Duas colunas lado a lado (antes × depois, opção A × B). esquerda/direita: (rótulo, [itens]).
    A coluna da direita é a recomendada/destacada (fundo petróleo)."""
    s = _novo_slide(prs)
    _titulo_slide(s, titulo, sobretitulo, subtitulo=subtitulo)
    gap = Inches(0.4)
    w = int((UTIL - gap) / 2)
    disp = BASE_CONTEUDO - TOPO_CONTEUDO
    larg_txt = w - Inches(0.9)
    for tam in (17, 16, 15, 14):
        altura_txt = max(sum(_altura_texto(t, tam, larg_txt, prefixo="●  ") + Pt(10) for t in col[1])
                         for col in (esquerda, direita))
        h = Inches(1.2) + altura_txt + Inches(0.35)
        if h <= disp:
            break
    h = min(h, disp)
    for k, coluna in enumerate((esquerda, direita)):
        rotulo, itens = coluna[0], coluna[1]
        nome_icone = coluna[2] if len(coluna) > 2 else None
        x = MARGEM + k * (w + gap)
        forte = k == 1
        _ret(s, x, TOPO_CONTEUDO, w, h, "primaria" if forte else "primaria_clara", f"Coluna {k + 1}", True, 0.05)
        dx = Inches(0.45)
        if nome_icone:
            icone(s, nome_icone, x + Inches(0.45), TOPO_CONTEUDO + Inches(0.32), Inches(0.62), escuro=forte,
                  sobre_claro=not forte)
            dx = Inches(1.3)
        tb = _caixa(s, x + dx, TOPO_CONTEUDO + Inches(0.4), w - dx - Inches(0.45), Inches(0.6), f"Rótulo coluna {k + 1}")
        r = _par(tb.text_frame, True).add_run()
        r.text = rotulo
        _fonte_run(r, 22, "destaque" if forte else "primaria", negrito=True, fonte=FONTE_TITULO)
        tb = _caixa(s, x + Inches(0.45), TOPO_CONTEUDO + Inches(1.2), larg_txt, h - Inches(1.35), f"Texto coluna {k + 1}")
        tf = tb.text_frame
        for j, item in enumerate(itens):
            p = _par(tf, j == 0, depois=10, linha=1.15)
            r = p.add_run()
            r.text = "●  "
            _fonte_run(r, 11, "destaque")
            _ricos(p, item, tam, "branco" if forte else "texto", cor_negrito="destaque" if forte else "primaria")
    _rodape(prs, s)
    return s


def slide_passos(prs, titulo, passos, sobretitulo=None, subtitulo=None, nota=None):
    """Processo em etapas (2 a 5): círculos numerados petróleo (o mesmo marcador dos manuais), ligados
    por uma linha fina, com título e texto curto embaixo de cada um. passos: [(titulo, texto), ...]."""
    s = _novo_slide(prs)
    _titulo_slide(s, titulo, sobretitulo, subtitulo=subtitulo)
    n = len(passos)
    w = int(UTIL / n)
    d = Inches(0.9)
    y = TOPO_CONTEUDO + Inches(0.5)
    linha = s.shapes.add_connector(1, MARGEM + w // 2, y + d // 2, MARGEM + w * (n - 1) + w // 2, y + d // 2)
    linha.line.color.rgb = _rgb("borda")
    linha.line.width = Pt(2.5)
    linha.name = "Linha dos passos"
    for i, (tit, txt) in enumerate(passos):
        cx = MARGEM + i * w + w // 2
        c = _circulo(s, cx - d // 2, y, d, "primaria", f"Passo {i + 1}")
        tf = c.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = _par(tf, True, PP_ALIGN.CENTER)
        r = p.add_run()
        r.text = str(i + 1)
        _fonte_run(r, 28, "destaque", negrito=True, fonte=FONTE_TITULO)
        tb = _caixa(s, MARGEM + i * w + Inches(0.15), y + d + Inches(0.3), w - Inches(0.3), Inches(2.6), f"Texto passo {i + 1}")
        p = _par(tb.text_frame, True, PP_ALIGN.CENTER, depois=6)
        r = p.add_run()
        r.text = tit
        _fonte_run(r, 19, "primaria", negrito=True, fonte=FONTE_TITULO)
        p = _par(tb.text_frame, alinhamento=PP_ALIGN.CENTER, linha=1.15)
        _ricos(p, txt, 15, "texto_suave")
    if nota:
        _nota(s, nota, BASE_CONTEUDO - Inches(0.55))
    _rodape(prs, s)
    return s


def slide_frase(prs, frase, autor=None):
    """Frase de impacto ou citação em slide escuro: aspas âmbar grandes, texto em dois pesos e autor,
    tudo numa caixa centralizada na altura do slide."""
    s = _novo_slide(prs, escuro=True)
    _simbolo_fundo(s, "texto_faixa", alfa=26)
    larg, alt = Inches(10), ALTURA - Inches(2.0)
    tb = _caixa(s, MARGEM, Inches(0.9), larg, alt, "Frase", MSO_ANCHOR.MIDDLE)
    r = _par(tb.text_frame, True, linha=0.75).add_run()
    r.text = "“"
    _fonte_run(r, 110, "destaque", negrito=True, fonte=FONTE_TITULO)
    for tam in (36, 32, 28, 26):
        if _altura_texto(frase.replace("**", ""), tam, larg, 1.1, FONTE_TITULO, True) + Pt(110 * 1.05 + 60) <= alt:
            break
    _titulo_misto(_par(tb.text_frame, linha=1.1), frase, tam)
    if autor:
        r = _par(tb.text_frame, antes=18).add_run()
        r.text = "— " + autor
        _fonte_run(r, 16, "creme")
    _rodape(prs, s, escuro=True)
    return s


def slide_imagem(prs, titulo, imagem, legenda=None, sobretitulo=None, subtitulo=None):
    """Imagem grande (print, foto, mapa) ocupando a área de conteúdo, com legenda curta embaixo."""
    s = _novo_slide(prs)
    _titulo_slide(s, titulo, sobretitulo, subtitulo=subtitulo)
    alt = BASE_CONTEUDO - TOPO_CONTEUDO - (Inches(0.45) if legenda else 0)
    from PIL import Image
    im = Image.open(imagem)
    prop = im.width / im.height
    w = min(UTIL, int(alt * prop))
    h = int(w / prop)
    x = MARGEM + (UTIL - w) // 2
    s.shapes.add_picture(str(imagem), x, TOPO_CONTEUDO, w, h).name = "Imagem"
    if legenda:
        tb = _caixa(s, MARGEM, TOPO_CONTEUDO + h + Inches(0.12), UTIL, Inches(0.35), "Legenda")
        p = _par(tb.text_frame, True, PP_ALIGN.CENTER)
        _ricos(p, legenda, 13, "texto_suave")
    _rodape(prs, s)
    return s


def encerramento(prs, titulo="Obrigado", contatos=None, subtitulo=None):
    """Último slide: petróleo, título em dois pesos e os contatos — tudo numa caixa centralizada na
    altura do slide (o título pode ter 2 linhas sem empurrar nada)."""
    s = _novo_slide(prs, escuro=True)
    _simbolo_fundo(s, "texto_faixa", alfa=38)
    if LOGO_CLARA.exists():
        s.shapes.add_picture(str(LOGO_CLARA), MARGEM, Inches(0.65), height=Inches(0.45)).name = "Logo"
    larg = Inches(9)
    alt = ALTURA - Inches(2.4)
    tb = _caixa(s, MARGEM, Inches(1.4), larg, alt, "Título", MSO_ANCHOR.MIDDLE)
    extra = (Pt(18 * 1.4 + 14) if subtitulo else 0) + Pt((16 * 1.4 + 4) * len(contatos or []) + (16 if contatos else 0))
    for tam in (60, 54, 48, 44, 40):
        if _altura_texto(titulo.replace("**", ""), tam, larg, 0.95, FONTE_TITULO, True) + extra <= alt:
            break
    _titulo_misto(_par(tb.text_frame, True, linha=0.95), titulo, tam)
    if subtitulo:
        r = _par(tb.text_frame, antes=14).add_run()
        r.text = subtitulo
        _fonte_run(r, 18, "creme")
    for i, c in enumerate(contatos or []):
        r = _par(tb.text_frame, antes=16 if i == 0 else 4).add_run()
        r.text = c
        _fonte_run(r, 16, "branco")
    return s


def _recortar(imagem, w_emu, h_emu):
    """Corta a imagem pelo centro na proporção da área (sem distorcer) e devolve PNG/JPEG em memória."""
    from PIL import Image
    im = Image.open(imagem).convert("RGB")
    alvo = w_emu / h_emu
    fw, fh = im.size
    if fw / fh > alvo:
        nw = int(fh * alvo)
        im = im.crop(((fw - nw) // 2, 0, (fw - nw) // 2 + nw, fh))
    else:
        nh = int(fw / alvo)
        im = im.crop((0, (fh - nh) // 2, fw, (fh - nh) // 2 + nh))
    buf = io.BytesIO()
    im.save(buf, format="JPEG", quality=90)
    buf.seek(0)
    return buf


def notas(slide, texto):
    """Anotações do apresentador (aparecem só no modo apresentador)."""
    slide.notes_slide.notes_text_frame.text = texto


def revisar(prs):
    """Confere cada slide antes de salvar: texto maior que a caixa, caixas de texto se sobrepondo e
    elementos fora do slide. Devolve a lista de problemas (texto). Mede com a fonte real."""
    problemas = []
    for n, slide in enumerate(prs.slides, 1):
        caixas = []
        for sh in slide.shapes:
            if sh.left is not None and (sh.left < -Emu(5000) or sh.top < -Emu(5000)
                                        or sh.left + sh.width > LARGURA + Emu(5000)
                                        or sh.top + sh.height > ALTURA + Emu(5000)):
                if "fundo" not in (sh.name or "").lower() and "Foto" not in (sh.name or ""):
                    problemas.append(f"slide {n}: \"{sh.name}\" sai do slide")
            if not sh.has_text_frame or not sh.text_frame.text.strip():
                continue
            tf = sh.text_frame
            larg = (sh.width - (tf.margin_left or 0) - (tf.margin_right or 0)) / 12700
            total = 0.0
            for p in tf.paragraphs:
                segs = [(r.text, r.font.name or FONTE, r.font.size.pt if r.font.size else 18, bool(r.font.bold))
                        for r in p.runs if r.text]
                if not segs:
                    continue
                linha = p.line_spacing if isinstance(p.line_spacing, float) else 1.0
                total += _altura_pt(segs, larg, linha)[1]
                total += (p.space_after.pt if p.space_after else 0) + (p.space_before.pt if p.space_before else 0)
            disp = (sh.height - (tf.margin_top or 0) - (tf.margin_bottom or 0)) / 12700
            if total > disp + 4:
                problemas.append(f"slide {n}: o texto de \"{sh.name}\" não cabe na caixa "
                                 f"({total:.0f} pt de texto em {disp:.0f} pt)")
            if sh.shape_type == 17:  # caixa de texto: guarda a área que o texto ocupa de verdade
                caixas.append((sh.name, sh.left, sh.top, sh.left + sh.width, sh.top + Emu(int(total * 12700))))
        for i in range(len(caixas)):
            for j in range(i + 1, len(caixas)):
                a, b = caixas[i], caixas[j]
                if a[1] < b[3] and b[1] < a[3] and a[2] < b[4] and b[2] < a[4]:
                    problemas.append(f"slide {n}: \"{a[0]}\" e \"{b[0]}\" se sobrepõem")
    return problemas


def salvar(prs, caminho, revisao=True):
    """Salva. Antes, roda revisar() e mostra o que estiver quebrado (texto estourando, sobreposição)."""
    if revisao:
        problemas = revisar(prs)
        for pr in problemas:
            print("⚠ REVISAR:", pr)
        if not problemas:
            print("✓ revisão: nenhum texto estourando nem sobreposto")
    prs.save(caminho)
    return caminho


# ---------------------------------------------------------------- exemplo
def _exemplo(saida):
    """Exemplo com dados FICTÍCIOS (Empresa Exemplo). Nunca coloque dados reais da empresa aqui: o
    repositório é público."""
    prs = nova_apresentacao("Despesas operacionais 2026")
    capa(prs, "Despesas **operacionais**", "Resultado das mudanças · janeiro de 2026", rodape_texto="Empresa Exemplo")
    divisor(prs, 1, "Onde **estamos**", "A despesa mensal depois das mudanças")
    slide_cartoes(prs, "O mês em **números**", [
        {"rotulo": "Despesa mensal", "valor": "R$ 1.800", "detalhe": "5 itens ativos", "destaque": True, "icone": "wallet"},
        {"rotulo": "Despesa anual", "valor": "R$ 21,6 mil", "detalhe": "projeção 12 meses", "icone": "calendar"},
        {"rotulo": "Economia / mês", "valor": "-25%", "detalhe": "R$ 600 a menos", "icone": "trending-down"},
    ], sobretitulo="Resumo", nota="Dados fictícios, para exemplo.")
    slide_grafico(prs, "Onde está o **custo**", ["Gestão", "Atendimento", "Site", "E-mail", "Nuvem"],
                  {"Custo mensal (R$)": [900, 450, 250, 120, 80]}, sobretitulo="Itens ativos", tipo="barras",
                  formato='"R$" #,##0', comentario=["**50%** do custo está no sistema de gestão.",
                                                    "Os outros 4 itens somam **R$ 900** por mês."],
                  fonte_dados="dados fictícios, para exemplo.")
    slide_tabela(prs, "O que **saiu**", ["Item", "Motivo", "Custo mensal"],
                 [["Sistema antigo de atendimento", "Substituído", "R$ 400,00"],
                  ["Ferramenta de design", "Pouco usada", "R$ 120,00"],
                  ["Plano extra de armazenamento", "Desnecessário", "R$ 80,00"]],
                 sobretitulo="Cancelamentos", alinhar_direita=["Custo mensal"],
                 total=["Total", "", "R$ 600,00"], fonte_dados="dados fictícios, para exemplo.")
    slide_comparacao(prs, "Antes × **depois**", ("Antes", ["R$ 2.400 por mês", "Dois sistemas de atendimento",
                                                           "Assinaturas sem uso"], "clock"),
                     ("Depois", ["**R$ 1.800** por mês", "Um sistema de atendimento",
                                 "Economia de **R$ 7,2 mil** por ano"], "rocket"), sobretitulo="Comparação")
    divisor(prs, 2, "Próximos **passos**")
    slide_passos(prs, "Como vamos **seguir**", [
        ("Revisar", "Assinaturas a cada trimestre."),
        ("Registrar", "Todo serviço novo na planilha."),
        ("Comparar", "Preços na renovação."),
        ("Avisar", "A equipe sobre cada mudança."),
    ], sobretitulo="Plano")
    slide_texto(prs, "O que muda para a **equipe**", [
        ("mail", "O atendimento passa a usar um **único e-mail** da empresa."),
        ("calculator", "Novos sistemas só entram depois de passar pela planilha de custos."),
        ("calendar", "Cada área revisa as assinaturas que usa a cada 3 meses."),
    ], sobretitulo="Combinados", destaque="**25%** a menos de custo, sem perder ferramentas.")
    slide_frase(prs, "Gastar **menos** para investir **melhor**", "Diretoria")
    encerramento(prs, "Obrigado", ["contato@empresaexemplo.com.br"], "Dúvidas e sugestões")
    salvar(prs, saida)
    return saida


def modelo(caminho):
    """Modelo em branco: um slide de cada tipo, com textos de instrução para trocar."""
    prs = nova_apresentacao("Modelo de apresentação Saldus")
    capa(prs, "Título da **apresentação**", "Subtítulo · área · mês/ano", rodape_texto="Saldus · Nome da área")
    divisor(prs, 1, "Nome da **seção**", "Uma frase sobre o que vem nesta parte")
    slide_texto(prs, "Título do **slide**", ["Uma ideia por linha, com até 2 linhas.",
                                             "Use **negrito** só no essencial.", "No máximo 4 a 5 itens."],
                sobretitulo="Assunto", destaque="A ideia principal deste slide em uma frase.")
    slide_cartoes(prs, "Os **números**", [
        {"rotulo": "Número principal", "valor": "R$ 0", "detalhe": "explicação curta", "destaque": True},
        {"rotulo": "Segundo número", "valor": "0", "detalhe": "explicação curta"},
        {"rotulo": "Variação", "valor": "+0%", "detalhe": "vs. período anterior"}], sobretitulo="Resumo")
    slide_grafico(prs, "Título do **gráfico**", ["Item A", "Item B", "Item C", "Item D"],
                  {"Valor": [40, 30, 20, 10]}, sobretitulo="Análise",
                  comentario=["Escreva aqui **o que o gráfico mostra**.", "Uma conclusão por frase."],
                  fonte_dados="indique a origem dos dados.")
    slide_tabela(prs, "Título da **tabela**", ["Item", "Responsável", "Valor"],
                 [["Item 1", "Nome", "R$ 0,00"], ["Item 2", "Nome", "R$ 0,00"]], sobretitulo="Detalhe",
                 alinhar_direita=["Valor"], total=["Total", "", "R$ 0,00"], fonte_dados="indique a origem.")
    slide_comparacao(prs, "Opção A × **opção B**", ("Opção A", ["Ponto 1", "Ponto 2"]),
                     ("Opção B (recomendada)", ["Ponto 1", "Ponto 2"]), sobretitulo="Comparação")
    slide_passos(prs, "Próximos **passos**", [("Passo 1", "O que fazer."), ("Passo 2", "O que fazer."),
                                               ("Passo 3", "O que fazer.")], sobretitulo="Plano")
    slide_frase(prs, "Uma frase de **impacto** aqui", "Autor")
    encerramento(prs, "Obrigado", ["email@saldus.com.br", "saldus.com.br"], "Dúvidas e sugestões")
    salvar(prs, caminho)
    return caminho


if __name__ == "__main__":
    if len(sys.argv) > 2 and sys.argv[1] == "--modelo":
        print("Gerado:", modelo(sys.argv[2]))
    else:
        print("Gerado:", _exemplo(sys.argv[1] if len(sys.argv) > 1 else "exemplo_saldus.pptx"))
