#!/usr/bin/env python3
"""Gera o termo de encerramento em DOCX. Uso: gerar_termo.py [entrada.json] saida.docx (sem entrada: modelo em branco)."""
import json
import sys
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

A_DEFINIR = "a definir"
FONTE = "Arial"
PRIMARIA, CLARA, DICA = "1F3A5F", "DCE6F1", "7F7F7F"
MOTIVOS = ["CONCLUSÃO", "SUBSTITUIÇÃO", "CANCELAMENTO"]
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
    borda = elemento("pBdr")
    borda.append(elemento("bottom", val="single", sz="8", space="1", color=PRIMARIA))
    p._p.get_or_add_pPr().append(borda)


def paragrafo(doc, texto, dica=False, negrito=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    formatar(p.add_run(texto), negrito=negrito, cor=DICA if dica else None, italico=dica)


def data(texto):
    return datetime.strptime(texto, "%d/%m/%Y")


def moeda(v, sinal=False):
    texto = f"{v:+,.2f}" if sinal else f"{v:,.2f}"
    return "R$ " + texto.replace(",", "_").replace(".", ",").replace("_", ".")


def pct(v):
    return f"{v:+.1f}%".replace(".", ",")


def comparativo(d):
    """Calcula prazo (dias corridos) e custo: original, mudanças, atualizado, realizado e desvio."""
    base, real = d.get("linha_base", {}), d.get("realizado", {})
    mudancas = d.get("mudancas_aprovadas", [])
    prazo_mud = sum(m.get("prazo_dias") or 0 for m in mudancas)
    custo_mud = sum(m.get("custo") or 0 for m in mudancas)
    r = {"prazo": {}, "custo": {}}
    if base.get("data_inicio") and base.get("data_fim"):
        original = (data(base["data_fim"]) - data(base["data_inicio"])).days
        r["prazo"] = {"original": original, "mudancas": prazo_mud, "atualizado": original + prazo_mud}
        if real.get("data_fim"):
            r["prazo"]["realizado"] = (data(real["data_fim"]) - data(base["data_inicio"])).days
    if base.get("custo") is not None:
        r["custo"] = {"original": base["custo"], "mudancas": custo_mud, "atualizado": base["custo"] + custo_mud}
        if real.get("custo") is not None:
            r["custo"]["realizado"] = real["custo"]
    for x in r.values():
        if "realizado" in x:
            x["desvio"] = x["realizado"] - x["atualizado"]
            x["desvio_pct"] = x["desvio"] / x["atualizado"] * 100 if x["atualizado"] else None
    return r


def gerar(d, destino, preencher):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Cm(2)
    for estilo in ("Normal", "List Bullet"):
        doc.styles[estilo].font.name = FONTE
        doc.styles[estilo].font.size = Pt(10)

    def v(valor):
        if not preencher:
            return ""
        return valor if valor not in (None, "", []) else A_DEFINIR

    titulo = doc.add_paragraph()
    formatar(titulo.add_run("TERMO DE ENCERRAMENTO DO PROJETO"), negrito=True, cor=PRIMARIA, tamanho=16)

    ident = d.get("identificacao", {})
    campos = [("Código do projeto", "codigo"), ("Nome do projeto", "titulo"), ("Data", "data"), ("Responsável", "responsavel")]
    t = tabela(doc, len(campos), [4.5, LARGURA - 4.5])
    for i, (rotulo, chave) in enumerate(campos):
        escrever(t.cell(i, 0), rotulo, negrito=True)
        sombrear(t.cell(i, 0), CLARA)
        escrever(t.cell(i, 1), v(ident.get(chave)))

    secao(doc, "Objetivo do documento")
    nome = ident.get("titulo") if preencher and ident.get("titulo") else "[nome do projeto]"
    paragrafo(doc, f"Este documento formaliza o encerramento do projeto {nome}, consolidando os resultados obtidos, "
                   "os principais aprendizados, as entregas realizadas e a comparação entre o planejado e o executado. "
                   "Com a validação das entregas, encerra-se a atuação da equipe do projeto e as responsabilidades "
                   "relacionadas ao escopo definido.")

    secao(doc, "Aprendizados do projeto")
    licoes = d.get("licoes", []) if preencher else []
    linhas = licoes or ([{}] if preencher else [{}] * 3)
    t = tabela(doc, len(linhas) + 1, [1.5, 8.0, 7.5])
    cabecalho_tabela(t, ["ID", "Lição aprendida", "Plano de ação"])
    for i, item in enumerate(linhas, 1):
        escrever(t.cell(i, 0), item.get("id") or (f"L{i:02d}" if item else ""))
        escrever(t.cell(i, 1), v(item.get("licao")))
        escrever(t.cell(i, 2), v(item.get("plano_acao")))

    secao(doc, "Comparativo previsto x realizado")
    comp = comparativo(d) if preencher else {"prazo": {}, "custo": {}}
    t = tabela(doc, 3, [2.3, 2.9, 2.6, 2.9, 3.0, 3.3])
    cabecalho_tabela(t, ["Indicador", "Previsto original", "Mudanças aprovadas", "Previsto atualizado", "Realizado", "Desvio"])
    formatos = {"prazo": lambda x, s=False: f"{x:+d} dias" if s else f"{x} dias", "custo": moeda}
    for i, (rotulo, chave) in enumerate((("Prazo", "prazo"), ("Custo", "custo")), 1):
        escrever(t.cell(i, 0), rotulo, negrito=True)
        x, fmt = comp[chave], formatos[chave]
        for j, campo in enumerate(("original", "mudancas", "atualizado", "realizado"), 1):
            if campo == "mudancas" and campo in x and not d.get("mudancas_aprovadas"):
                escrever(t.cell(i, j), "Nenhuma informada", dica=True)
            else:
                escrever(t.cell(i, j), fmt(x[campo], campo == "mudancas") if campo in x else v(None))
        if "desvio" in x:
            texto = fmt(x["desvio"], True) + (f" ({pct(x['desvio_pct'])})" if x["desvio_pct"] is not None else "")
            escrever(t.cell(i, 5), texto, negrito=True)
        else:
            escrever(t.cell(i, 5), v(None))
    mudancas = d.get("mudancas_aprovadas", []) if preencher else []
    if mudancas:
        paragrafo(doc, "Mudanças aprovadas consideradas: " + "; ".join(
            f"{m.get('codigo', A_DEFINIR)} ({m.get('prazo_dias') or 0:+d} dias, {moeda(m.get('custo') or 0, True)})" for m in mudancas) + ".", dica=True)

    secao(doc, "Status do escopo")
    escopo = d.get("escopo", {})
    perguntas = [("Escopo planejado concluído?", "concluido", False, "Caso não, justificar"),
                 ("Entregas homologadas?", "homologado", False, "Caso não, justificar"),
                 ("Pendências remanescentes?", "pendencias", True, "Caso sim, descrever")]
    t = tabela(doc, 3, [5.0, 2.6, 9.4])
    for i, (rotulo, chave, exige_se, dica) in enumerate(perguntas):
        escrever(t.cell(i, 0), rotulo, negrito=True)
        sombrear(t.cell(i, 0), CLARA)
        resposta = escopo.get(chave)
        if not preencher:
            escrever(t.cell(i, 1), "☐ Sim  ☐ Não")
            escrever(t.cell(i, 2), dica, dica=True)
            continue
        escrever(t.cell(i, 1), {True: "Sim", False: "Não", None: A_DEFINIR}[resposta], negrito=resposta is not None)
        detalhe = escopo.get(f"{chave}_detalhe")
        escrever(t.cell(i, 2), detalhe or "-", dica=not detalhe)

    secao(doc, "Motivo do encerramento")
    motivo = d.get("motivo") if preencher else None
    detalhes = {"SUBSTITUIÇÃO": "indicar o projeto que o substitui", "CANCELAMENTO": "descrever o motivo"}
    for op in MOTIVOS:
        marca = "☒" if motivo == op else "☐"
        texto = f"{marca}  {op.capitalize()}"
        if op in detalhes:
            texto += f": {d.get('motivo_detalhe')}" if motivo == op and d.get("motivo_detalhe") else f" ({detalhes[op]})"
        paragrafo(doc, texto, negrito=motivo == op)
    if motivo == "CONCLUSÃO" and escopo.get("concluido") is True:
        paragrafo(doc, "Declara-se que as entregas previstas no escopo do projeto foram concluídas e que a documentação "
                       "aplicável foi elaborada, armazenada e disponibilizada às áreas responsáveis.")

    secao(doc, "Considerações finais")
    paragrafo(doc, v(d.get("consideracoes")) or " ")

    secao(doc, "Assinaturas")
    assinaturas = d.get("assinaturas", {})
    papeis = [("Gerente do projeto", "gerente"), ("Sponsor", "sponsor"), ("PMO", "pmo")]
    t = tabela(doc, len(papeis) + 1, [4.0, 6.0, 4.5, 2.5])
    cabecalho_tabela(t, ["Papel", "Nome", "Assinatura", "Data"])
    for i, (papel, chave) in enumerate(papeis, 1):
        escrever(t.cell(i, 0), papel, negrito=True)
        escrever(t.cell(i, 1), v(assinaturas.get(chave)))
        t.rows[i].height = Cm(1.0)

    doc.core_properties.author = doc.core_properties.last_modified_by = "pmo-ai-skills"
    doc.core_properties.title = "Termo de Encerramento do Projeto"
    doc.save(destino)
    return comp


def validar(d):
    erros = []
    base, real = d.get("linha_base", {}), d.get("realizado", {})
    for rotulo, valor in (("linha_base.data_inicio", base.get("data_inicio")), ("linha_base.data_fim", base.get("data_fim")),
                          ("realizado.data_fim", real.get("data_fim"))):
        if valor is not None:
            try:
                data(valor)
            except (TypeError, ValueError):
                erros.append(f"{rotulo} deve estar no formato dd/mm/aaaa")
    if not erros and base.get("data_inicio") and base.get("data_fim") and data(base["data_fim"]) < data(base["data_inicio"]):
        erros.append("linha_base.data_fim não pode ser anterior à data de início")
    for rotulo, valor in (("linha_base.custo", base.get("custo")), ("realizado.custo", real.get("custo"))):
        if valor is not None and (type(valor) not in (int, float) or valor < 0):
            erros.append(f"{rotulo} deve ser um número não negativo")
    for i, m in enumerate(d.get("mudancas_aprovadas", []), 1):
        if m.get("prazo_dias") is not None and type(m["prazo_dias"]) is not int:
            erros.append(f"mudancas_aprovadas[{i}].prazo_dias deve ser um número inteiro")
        if m.get("custo") is not None and type(m["custo"]) not in (int, float):
            erros.append(f"mudancas_aprovadas[{i}].custo deve ser um número")
    escopo = d.get("escopo", {})
    for chave, exige_quando in (("concluido", False), ("homologado", False), ("pendencias", True)):
        valor = escopo.get(chave)
        if valor is not None and not isinstance(valor, bool):
            erros.append(f"escopo.{chave} deve ser true ou false")
        elif valor is exige_quando and not escopo.get(f"{chave}_detalhe"):
            erros.append(f"escopo.{chave}_detalhe é obrigatório quando escopo.{chave} é {str(exige_quando).lower()}")
    motivo = d.get("motivo")
    if motivo is not None and motivo not in MOTIVOS:
        erros.append(f"motivo deve ser um destes: {', '.join(MOTIVOS)}")
    elif motivo in ("SUBSTITUIÇÃO", "CANCELAMENTO") and not d.get("motivo_detalhe"):
        erros.append(f"motivo_detalhe é obrigatório quando o motivo é {motivo}")
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
    comp = gerar(dados, sys.argv[2], preencher=True)
    partes = []
    for rotulo, chave, fmt in (("prazo", "prazo", lambda x: f"{x:+d} dias"), ("custo", "custo", lambda x: moeda(x, True))):
        x = comp[chave]
        partes.append(f"{rotulo} {fmt(x['desvio'])} ({pct(x['desvio_pct'])})" if x.get("desvio_pct") is not None else f"{rotulo} {A_DEFINIR}")
    print(f"Termo: {sys.argv[2]} | desvio sobre o previsto atualizado: " + ", ".join(partes))


if __name__ == "__main__":
    main()
