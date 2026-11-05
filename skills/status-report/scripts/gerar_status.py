#!/usr/bin/env python3
"""Gera o status report em um slide PPTX. Uso: gerar_status.py entrada.json saida.pptx"""
import json
import sys
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

A_DEFINIR = "a definir"
FONTE = "Arial"
PRIMARIA, CLARA, BORDA, TEXTO, SUAVE = "1F3A5F", "DCE6F1", "BFBFBF", "262626", "7F7F7F"
CRITICIDADE = {"BAIXO": 15, "MÉDIO": 10, "ALTO": 5}  # limite de escalonamento, em %
ACOES = {"BAIXO": "acionar o gerente de projetos", "MÉDIO": "acionar o sponsor e o gerente",
         "ALTO": "convocar gerentes e diretoria"}
FAROL = {"Verde": ("2E8B57", "No prazo"), "Amarelo": ("E0A800", "Atenção"),
         "Vermelho": ("C0392B", "Risco de atraso"), "Cinza": ("9E9E9E", A_DEFINIR)}
LIMITES = {"proximas_atividades": 5, "atividades_atrasadas": 5, "riscos": 4}


def cor(hexa):
    return RGBColor.from_string(hexa)


def numero(v):
    return type(v) in (int, float)


def validar(d):
    erros = []
    for campo in ("previsto", "realizado"):
        v = d.get(campo)
        if v is not None and (not numero(v) or not 0 <= v <= 100):
            erros.append(f"{campo} deve ser um número de 0 a 100")
    if d.get("criticidade") is not None and d["criticidade"] not in CRITICIDADE:
        erros.append(f"criticidade deve ser um destes: {', '.join(CRITICIDADE)}")
    curva = d.get("curva_s")
    if curva:
        n = len(curva.get("periodos", []))
        for serie in ("previsto", "realizado"):
            valores = curva.get(serie, [])
            if len(valores) > n:
                erros.append(f"curva_s.{serie} tem mais valores que períodos")
            preenchidos = [v for v in valores if v is not None]
            if any(not numero(v) or not 0 <= v <= 100 for v in preenchidos):
                erros.append(f"curva_s.{serie} deve ter números de 0 a 100")
            elif preenchidos != sorted(preenchidos):
                erros.append(f"curva_s.{serie} deve ser acumulada (não pode diminuir)")
    for i, r in enumerate(d.get("riscos", []), 1):
        for campo in ("probabilidade", "severidade"):
            v = r.get(campo)
            if v is not None and (type(v) is not int or not 1 <= v <= 5):
                erros.append(f"riscos[{i}].{campo} deve ser um inteiro de 1 a 5")
    if erros:
        sys.exit("Entrada inválida:\n- " + "\n- ".join(erros))


def calcular_farol(d):
    prev, real, crit = d.get("previsto"), d.get("realizado"), d.get("criticidade")
    if prev is None or real is None or crit is None:
        return "Cinza", None
    if prev == 0:
        return "Verde", 0.0
    desvio = (prev - real) / prev * 100
    limite = CRITICIDADE[crit]
    if desvio > limite:
        return "Vermelho", desvio
    return ("Amarelo" if desvio > limite / 2 else "Verde"), desvio


def texto(slide, x, y, w, h, valor, tamanho=9, negrito=False, cor_texto=TEXTO, alinhar=PP_ALIGN.LEFT):
    caixa = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = caixa.text_frame
    tf.margin_left = tf.margin_right = Inches(0.03)
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = alinhar
    r = p.add_run()
    r.text = str(valor)
    r.font.name, r.font.size, r.font.bold = FONTE, Pt(tamanho), negrito
    r.font.color.rgb = cor(cor_texto)
    return caixa


def retangulo(slide, x, y, w, h, preenchimento=None, contorno=None, forma=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(forma, Inches(x), Inches(y), Inches(w), Inches(h))
    if preenchimento:
        s.fill.solid()
        s.fill.fore_color.rgb = cor(preenchimento)
    else:
        s.fill.background()
    if contorno:
        s.line.color.rgb = cor(contorno)
        s.line.width = Pt(0.75)
    else:
        s.line.fill.background()
    s.shadow.inherit = False
    return s


def painel(slide, x, y, w, h, titulo):
    retangulo(slide, x, y, w, h, contorno=BORDA)
    texto(slide, x + 0.1, y + 0.02, w - 0.2, 0.26, titulo, tamanho=11, negrito=True, cor_texto=PRIMARIA)
    retangulo(slide, x + 0.1, y + 0.3, w - 0.2, 0.012, preenchimento=PRIMARIA)


def tabela(slide, x, y, w, colunas, linhas, chave_lista, avisos):
    """colunas: (título, fração da largura, função que extrai o valor)."""
    limite = LIMITES[chave_lista]
    if len(linhas) > limite:
        avisos.append(f"{chave_lista}: {len(linhas) - limite} item(ns) não couberam no slide")
    linhas = linhas[:limite]
    if not linhas:
        texto(slide, x, y, w, 0.25, A_DEFINIR, cor_texto=SUAVE)
        return
    forma = slide.shapes.add_table(len(linhas) + 1, len(colunas), Inches(x), Inches(y), Inches(w), Inches(0.22 * (len(linhas) + 1)))
    t = forma.table
    t.first_row = True
    for j, (titulo, fracao, _) in enumerate(colunas):
        t.columns[j].width = Inches(w * fracao)
    for i in range(len(linhas) + 1):
        t.rows[i].height = Inches(0.22)
        for j, (titulo, _, extrair) in enumerate(colunas):
            c = t.cell(i, j)
            valor = titulo if i == 0 else extrair(linhas[i - 1])
            c.text = A_DEFINIR if valor in (None, "") else str(valor)
            c.margin_left = c.margin_right = Inches(0.04)
            c.margin_top = c.margin_bottom = Inches(0.01)
            c.fill.solid()
            c.fill.fore_color.rgb = cor(PRIMARIA if i == 0 else ("FFFFFF" if i % 2 else "F5F8FC"))
            for p in c.text_frame.paragraphs:
                for r in p.runs:
                    r.font.name, r.font.size, r.font.bold = FONTE, Pt(8), i == 0
                    r.font.color.rgb = cor("FFFFFF" if i == 0 else TEXTO)


def curva_s(slide, x, y, w, h, curva):
    if not curva or not curva.get("periodos"):
        texto(slide, x, y, w, 0.25, A_DEFINIR, cor_texto=SUAVE)
        return
    dados = CategoryChartData(number_format='0"%"')
    dados.categories = curva["periodos"]
    n = len(curva["periodos"])
    for nome, serie in (("Previsto", "previsto"), ("Realizado", "realizado")):
        valores = list(curva.get(serie, []))
        dados.add_series(nome, valores + [None] * (n - len(valores)))
    grafico = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(x), Inches(y), Inches(w), Inches(h), dados).chart
    grafico.has_legend = True
    grafico.legend.position = XL_LEGEND_POSITION.TOP
    grafico.legend.include_in_layout = False
    grafico.legend.font.size, grafico.legend.font.name = Pt(8), FONTE
    for eixo in (grafico.category_axis, grafico.value_axis):
        eixo.tick_labels.font.size, eixo.tick_labels.font.name = Pt(8), FONTE
        eixo.format.line.color.rgb = cor(BORDA)
    grafico.value_axis.maximum_scale, grafico.value_axis.minimum_scale = 100, 0
    grafico.value_axis.major_unit = 25
    grafico.value_axis.major_gridlines.format.line.color.rgb = cor("E7E7E7")
    for serie, hexa, tracejado in zip(grafico.plots[0].series, (SUAVE, PRIMARIA), (True, False)):
        serie.smooth = False
        serie.format.line.color.rgb = cor(hexa)
        serie.format.line.width = Pt(2)
        if tracejado:
            serie.format.line.dash_style = 4  # traço
        serie.marker.format.fill.solid()
        serie.marker.format.fill.fore_color.rgb = cor(hexa)
        serie.marker.format.line.color.rgb = cor(hexa)


def gerar(d, destino):
    ident = d.get("identificacao", {})
    v = lambda chave: ident.get(chave) or A_DEFINIR
    pct = lambda chave: f"{d[chave]:g}%" if d.get(chave) is not None else A_DEFINIR
    farol, desvio = calcular_farol(d)
    avisos = []

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(10), Inches(5.625)
    s = prs.slides.add_slide(prs.slide_layouts[6])

    retangulo(s, 0, 0, 10, 0.62, preenchimento=PRIMARIA)
    texto(s, 0.2, 0.06, 6.3, 0.5, f"STATUS REPORT | {v('titulo')}", tamanho=15, negrito=True, cor_texto="FFFFFF")
    for i, (rotulo, valor) in enumerate((("Código:", v("codigo")), ("Data do status:", v("data_status")))):
        texto(s, 6.6, 0.11 + i * 0.2, 1.3, 0.18, rotulo, tamanho=9, cor_texto="FFFFFF", alinhar=PP_ALIGN.RIGHT)
        texto(s, 7.95, 0.11 + i * 0.2, 1.95, 0.18, valor, tamanho=9, negrito=True, cor_texto="FFFFFF")

    retangulo(s, 0, 0.62, 10, 0.45, preenchimento=CLARA)
    for i, (rotulo, valor) in enumerate((("Início:", v("data_inicio")), ("Fim:", v("data_fim")))):
        texto(s, 0.2, 0.66 + i * 0.18, 0.55, 0.18, rotulo, negrito=True)
        texto(s, 0.75, 0.66 + i * 0.18, 1.0, 0.18, valor)
    texto(s, 1.8, 0.66, 0.85, 0.36, "Criticidade:", negrito=True)
    texto(s, 2.62, 0.66, 0.8, 0.36, d.get("criticidade") or A_DEFINIR)
    for x, rotulo, valor in ((3.5, "Previsto:", pct("previsto")), (4.85, "Realizado:", pct("realizado"))):
        texto(s, x, 0.66, 0.75, 0.36, rotulo, negrito=True)
        texto(s, x + 0.75, 0.66, 0.6, 0.36, valor, tamanho=12 if valor != A_DEFINIR else 9, negrito=True, cor_texto=PRIMARIA)
    texto(s, 6.2, 0.66, 0.75, 0.36, "Desvio:", negrito=True)
    texto(s, 6.85, 0.66, 0.7, 0.36, f"{desvio:.1f}%".replace(".", ",") if desvio is not None else A_DEFINIR,
          tamanho=12 if desvio is not None else 9, negrito=True, cor_texto=PRIMARIA)
    hexa, rotulo = FAROL[farol]
    texto(s, 7.65, 0.66, 0.55, 0.36, "Farol:", negrito=True)
    retangulo(s, 8.2, 0.73, 0.22, 0.22, preenchimento=hexa, forma=MSO_SHAPE.OVAL)
    texto(s, 8.48, 0.66, 1.45, 0.36, rotulo, negrito=True, cor_texto=hexa)

    painel(s, 0.2, 1.15, 4.75, 2.05, "Curva S do projeto")
    curva_s(s, 0.3, 1.48, 4.55, 1.68, d.get("curva_s"))

    painel(s, 5.05, 1.15, 4.75, 2.05, "Próximas atividades")
    tabela(s, 5.15, 1.55, 4.45, [("Atividade", 0.55, lambda a: a.get("atividade")), ("Responsável", 0.25, lambda a: a.get("responsavel")),
                                  ("Data fim", 0.2, lambda a: a.get("data_fim"))], d.get("proximas_atividades", []), "proximas_atividades", avisos)

    painel(s, 0.2, 3.3, 4.75, 2.15, "Atividades atrasadas ou com risco de atraso")
    tabela(s, 0.3, 3.7, 4.45, [("Atividade", 0.55, lambda a: a.get("atividade")), ("Área", 0.25, lambda a: a.get("area")),
                                ("Data fim", 0.2, lambda a: a.get("data_fim"))], d.get("atividades_atrasadas", []), "atividades_atrasadas", avisos)

    painel(s, 5.05, 3.3, 4.75, 2.15, "Riscos do projeto")
    tabela(s, 5.15, 3.7, 4.45, [("Risco", 0.38, lambda r: r.get("descricao")), ("Prob.", 0.12, lambda r: r.get("probabilidade")),
                                 ("Sev.", 0.11, lambda r: r.get("severidade")), ("Plano de ação", 0.39, lambda r: r.get("plano_acao"))],
           d.get("riscos", []), "riscos", avisos)

    prs.core_properties.author = prs.core_properties.last_modified_by = "pmo-ai-skills"
    prs.core_properties.title = "Status Report"
    prs.save(destino)
    return farol, desvio, avisos


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    dados = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    validar(dados)
    farol, desvio, avisos = gerar(dados, sys.argv[2])
    resumo = f"Status: {sys.argv[2]} | farol {farol}"
    if desvio is not None:
        resumo += f" | desvio {desvio:.1f}%".replace(".", ",")
    if farol == "Vermelho":
        resumo += f" | escalonar: {ACOES[dados['criticidade']]}"
    print(resumo)
    for aviso in avisos:
        print(f"Aviso: {aviso}")


if __name__ == "__main__":
    main()
