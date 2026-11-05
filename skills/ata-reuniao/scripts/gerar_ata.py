#!/usr/bin/env python3
"""Gera a ata de reunião em DOCX. Uso: gerar_ata.py [entrada.json] saida.docx (sem entrada: modelo em branco)."""
import json
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

A_DEFINIR = "a definir"
FONTE = "Arial"
PRIMARIA, CLARA, DICA = "1F3A5F", "DCE6F1", "7F7F7F"
STATUS = ["Pendente", "Em andamento", "Concluído"]
LARGURA = 17.0  # cm úteis em A4 com margens de 2 cm
ORDEM_TBLPR = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW",
               "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook"]


def elemento(tag, **atributos):
    el = OxmlElement(f"w:{tag}")
    for k, v in atributos.items():
        el.set(qn(f"w:{k}"), v)
    return el


def sombrear(celula, cor):
    celula._tc.get_or_add_tcPr().append(elemento("shd", val="clear", color="auto", fill=cor))


def formatar(run, negrito=False, cor=None, italico=False, tamanho=10):
    run.font.name, run.font.size, run.bold, run.italic = FONTE, Pt(tamanho), negrito, italico
    if cor:
        run.font.color.rgb = RGBColor.from_string(cor)


def escrever(celula, texto, negrito=False, cor=None, dica=False):
    celula.text = ""
    p = celula.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    formatar(p.add_run(str(texto)), negrito, DICA if dica else cor, italico=dica)


def tabela(doc, linhas, larguras):
    """Tabela com largura fixa e bordas explícitas, para não depender do estilo de quem abre."""
    t = doc.add_table(rows=linhas, cols=len(larguras))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    tblPr = t._tbl.tblPr
    bordas = elemento("tblBorders")
    for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
        bordas.append(elemento(lado, val="single", sz="4", space="0", color="BFBFBF"))
    filhos = {el.tag.split("}")[1]: el for el in tblPr}
    filhos.update({"tblW": elemento("tblW", w=str(int(sum(larguras) * 567)), type="dxa"),
                   "tblLayout": elemento("tblLayout", type="fixed"), "tblBorders": bordas})
    for el in list(tblPr):
        tblPr.remove(el)
    for tag in ORDEM_TBLPR:  # o Word exige esta ordem
        if tag in filhos:
            tblPr.append(filhos[tag])
    for gc, w in zip(t._tbl.tblGrid.findall(qn("w:gridCol")), larguras):
        gc.set(qn("w:w"), str(int(w * 567)))
    for j, w in enumerate(larguras):
        for celula in t.columns[j].cells:
            celula.width = Cm(w)
    return t


def cabecalho_tabela(t, titulos):
    for j, titulo in enumerate(titulos):
        escrever(t.cell(0, j), titulo, negrito=True)
        sombrear(t.cell(0, j), CLARA)


def secao(doc, titulo):
    p = doc.add_paragraph()
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(10), Pt(4)
    formatar(p.add_run(titulo.upper()), negrito=True, cor=PRIMARIA, tamanho=11)
    p_pr = p._p.get_or_add_pPr()
    borda = elemento("pBdr")
    borda.append(elemento("bottom", val="single", sz="8", space="1", color=PRIMARIA))
    p_pr.append(borda)


def marcador(doc, texto):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(2)
    formatar(p.add_run(texto))


def paragrafo(doc, texto, dica=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    formatar(p.add_run(texto), cor=DICA if dica else None, italico=dica)


def numero_pagina(paragrafo_rodape):
    formatar(paragrafo_rodape.add_run("Página "), cor=DICA, tamanho=8)
    run = paragrafo_rodape.add_run()
    formatar(run, cor=DICA, tamanho=8)
    run._r.append(elemento("fldChar", fldCharType="begin"))
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    run._r.append(instr)
    run._r.append(elemento("fldChar", fldCharType="separate"))
    resultado = OxmlElement("w:t")
    resultado.text = "1"  # valor provisório para visualizadores que não calculam campos; o Word atualiza
    run._r.append(resultado)
    run._r.append(elemento("fldChar", fldCharType="end"))


def gerar(d, destino, preencher):
    def v(valor, dica=""):
        if not preencher:
            return dica, True
        return (valor, False) if valor not in (None, "", []) else (A_DEFINIR, False)

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Cm(2)
    numero_pagina(sec.footer.paragraphs[0])

    titulo = doc.add_paragraph()
    formatar(titulo.add_run("ATA DE REUNIÃO"), negrito=True, cor=PRIMARIA, tamanho=16)

    ident = d.get("identificacao", {})
    participantes = d.get("participantes", [])
    texto_participantes = "; ".join(f"{p['nome']} ({p['funcao']})" if p.get("funcao") else p["nome"] for p in participantes)
    campos = [("Título", ident.get("titulo")), ("Projeto ou área", ident.get("projeto")), ("Data", ident.get("data")),
              ("Horário", ident.get("horario")), ("Duração", ident.get("duracao")), ("Local ou plataforma", ident.get("local")),
              ("Participantes", texto_participantes)]
    t = tabela(doc, len(campos), [4.5, LARGURA - 4.5])
    for i, (rotulo, valor) in enumerate(campos):
        escrever(t.cell(i, 0), rotulo, negrito=True)
        sombrear(t.cell(i, 0), CLARA)
        texto, dica = v(valor)
        escrever(t.cell(i, 1), texto, dica=dica)

    secao(doc, "Objetivo da reunião")
    paragrafo(doc, *v(d.get("objetivo"), "Por que a reunião aconteceu e o que precisava sair dela."))

    secao(doc, "Resumo executivo")
    resumo = d.get("resumo")
    if preencher and isinstance(resumo, list) and resumo:
        for item in resumo:
            marcador(doc, item)
    else:
        paragrafo(doc, *v(resumo, "Principais pontos discutidos, em poucas linhas."))

    iniciativas = d.get("iniciativas", [])
    if iniciativas or not preencher:
        secao(doc, "Iniciativas e projetos mapeados")
        linhas = iniciativas or [{}] * 2
        t = tabela(doc, len(linhas) + 1, [5.0, LARGURA - 5.0])
        cabecalho_tabela(t, ["Categoria", "Iniciativa ou projeto"])
        for i, item in enumerate(linhas, 1):
            escrever(t.cell(i, 0), item.get("categoria") or (A_DEFINIR if preencher else ""))
            escrever(t.cell(i, 1), item.get("nome") or (A_DEFINIR if preencher else ""))

    secao(doc, "Decisões")
    decisoes = d.get("decisoes", []) if preencher else []
    linhas = decisoes or ([] if preencher else [""] * 3)
    if preencher and not decisoes:
        paragrafo(doc, "Nenhuma decisão registrada nesta reunião.")
    else:
        t = tabela(doc, len(linhas) + 1, [1.2, LARGURA - 1.2])
        cabecalho_tabela(t, ["#", "Decisão"])
        for i, decisao in enumerate(linhas, 1):
            escrever(t.cell(i, 0), str(i))
            escrever(t.cell(i, 1), decisao)

    secao(doc, "Itens de ação")
    acoes = d.get("acoes", []) if preencher else []
    if preencher and not acoes:
        paragrafo(doc, "Nenhum item de ação registrado nesta reunião.")
    else:
        linhas = acoes or [{}] * 3
        t = tabela(doc, len(linhas) + 1, [1.2, 7.8, 3.4, 2.2, 2.4])
        cabecalho_tabela(t, ["#", "Item de ação", "Responsável", "Prazo", "Status"])
        for i, item in enumerate(linhas, 1):
            escrever(t.cell(i, 0), str(i))
            for j, chave in enumerate(("descricao", "responsavel", "prazo", "status"), 1):
                escrever(t.cell(i, j), item.get(chave) or (A_DEFINIR if preencher else ""))

    secao(doc, "Próximos passos")
    passos = d.get("proximos_passos", [])
    if preencher and passos:
        for passo in passos:
            marcador(doc, passo)
    else:
        paragrafo(doc, *v(None, "Próxima reunião, entregas esperadas e encaminhamentos."))

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    texto, dica = v(d.get("elaborado_por"), "[nome]")
    formatar(p.add_run(f"Ata elaborada por {texto}."), cor=DICA, italico=True, tamanho=9)

    for estilo in ("Normal", "List Bullet"):
        doc.styles[estilo].font.name = FONTE
        doc.styles[estilo].font.size = Pt(10)
    doc.core_properties.author = doc.core_properties.last_modified_by = "pmo-ai-skills"
    doc.core_properties.title = "Ata de Reunião"
    doc.save(destino)


def validar(d):
    erros = []
    for i, p in enumerate(d.get("participantes", []), 1):
        if not p.get("nome"):
            erros.append(f"participantes[{i}].nome é obrigatório")
    for i, a in enumerate(d.get("acoes", []), 1):
        if not a.get("descricao"):
            erros.append(f"acoes[{i}].descricao é obrigatória")
        if a.get("status") is not None and a["status"] not in STATUS:
            erros.append(f"acoes[{i}].status deve ser um destes: {', '.join(STATUS)}")
    for i, dec in enumerate(d.get("decisoes", []), 1):
        if not isinstance(dec, str) or not dec.strip():
            erros.append(f"decisoes[{i}] deve ser um texto")
    if erros:
        sys.exit("Entrada inválida:\n- " + "\n- ".join(erros))


def main():
    if len(sys.argv) not in (2, 3):
        sys.exit(__doc__)
    if len(sys.argv) == 2:
        gerar({}, sys.argv[1], preencher=False)
        print(f"Modelo em branco: {sys.argv[1]}")
        return
    dados = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    validar(dados)
    gerar(dados, sys.argv[2], preencher=True)
    acoes = dados.get("acoes", [])
    lacunas = sum(1 for a in acoes if not a.get("responsavel") or not a.get("prazo"))
    print(f"Ata: {sys.argv[2]} | {len(dados.get('decisoes', []))} decisões | {len(acoes)} ações"
          + (f" | {lacunas} ação(ões) sem responsável ou prazo" if lacunas else ""))


if __name__ == "__main__":
    main()
