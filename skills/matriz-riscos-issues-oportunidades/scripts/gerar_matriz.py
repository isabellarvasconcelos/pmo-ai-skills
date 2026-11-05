#!/usr/bin/env python3
"""Gera a matriz de riscos em XLSX. Uso: gerar_matriz.py [entrada.json] saida.xlsx (sem entrada: modelo em branco)."""
import json
import sys
from collections import Counter
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

A_DEFINIR = "a definir"
PRIMARIA, CLARA = "1F3A5F", "DCE6F1"
FONTE = "Arial"
BORDA = Border(*[Side(style="thin", color="BFBFBF")] * 4)
MOEDA = '"R$" #,##0.00'
SIGNIFICATIVO, IMPORTANTE = 15, 5  # notas mínimas de cada grau
CORES = {"Não Significativo": "C6E0B4", "Importante": "FFE699", "Significativo": "F4B0A5"}
TIPOS = ["RISCO", "ISSUE", "OPORTUNIDADE"]
STATUS = ["ABERTO", "CONCLUÍDO", "CANCELADO"]
MEDIDAS_AMEACA = ["Mitigar", "Compartilhar", "Transferir", "Eliminar", "Justificar"]
MEDIDAS_OPORTUNIDADE = ["Explorar", "Compartilhar", "Melhorar", "Aceitar"]
MEDIDAS = MEDIDAS_AMEACA + [m for m in MEDIDAS_OPORTUNIDADE if m not in MEDIDAS_AMEACA]
ESCALA = ["1", "2", "3", "4", "5"]
LINHA_INICIAL = 8
LINHAS_LIVRES = 15
COLUNAS = [("ID", 10), ("Descrição", 42), ("Tipo", 15), ("Consequência", 34), ("Nível (1-5)", 9),
           ("Probabilidade (1-5)", 13), ("Grau", 17), ("Medida", 14), ("Plano de ação", 42),
           ("Custo", 14), ("Status", 13), ("Observações", 30)]


def grau(nota):
    if nota >= SIGNIFICATIVO:
        return "Significativo"
    return "Importante" if nota >= IMPORTANTE else "Não Significativo"


def formula_nota(ref):
    return f'IF({ref}>={SIGNIFICATIVO},"Significativo",IF({ref}>={IMPORTANTE},"Importante","Não Significativo"))'


def fonte(**k):
    return Font(name=FONTE, size=10, **k)


def validar(d):
    erros, vistos = [], set()
    for i, item in enumerate(d.get("itens", []), 1):
        rotulo = f"itens[{i}]"
        if not item.get("descricao"):
            erros.append(f"{rotulo}.descricao é obrigatória")
        if item.get("tipo") not in TIPOS:
            erros.append(f"{rotulo}.tipo deve ser um destes: {', '.join(TIPOS)}")
        for campo in ("consequencia_nivel", "probabilidade"):
            v = item.get(campo)
            if v is not None and (type(v) is not int or not 1 <= v <= 5):
                erros.append(f"{rotulo}.{campo} deve ser um inteiro de 1 a 5")
        permitidas = MEDIDAS_OPORTUNIDADE if item.get("tipo") == "OPORTUNIDADE" else MEDIDAS_AMEACA
        if item.get("medida") is not None and item["medida"] not in permitidas:
            erros.append(f"{rotulo}.medida deve ser um destes: {', '.join(permitidas)}")
        if item.get("status") is not None and item["status"] not in STATUS:
            erros.append(f"{rotulo}.status deve ser um destes: {', '.join(STATUS)}")
        c = item.get("custo")
        if c is not None and (type(c) not in (int, float) or c < 0):
            erros.append(f"{rotulo}.custo deve ser um número não negativo")
        if item.get("id"):
            if item["id"] in vistos:
                erros.append(f"{rotulo}.id repetido: {item['id']}")
            vistos.add(item["id"])
    if erros:
        sys.exit("Entrada inválida:\n- " + "\n- ".join(erros))


def numerar(itens):
    """Mantém os IDs informados e dá aos novos o próximo número livre, sem reaproveitar lacunas."""
    numeros = [int(i["id"][4:]) for i in itens if str(i.get("id", "")).startswith("RISC") and i["id"][4:].isdigit()]
    prox = max(numeros, default=0) + 1
    saida = []
    for item in itens:
        item = dict(item)
        if not item.get("id"):
            item["id"] = f"RISC{prox:03d}"
            prox += 1
        saida.append(item)
    return saida, prox


def aba_matriz(ws, d, preencher):
    ident = d.get("identificacao", {})
    ws.merge_cells("A1:L1")
    ws["A1"] = "MATRIZ DE RISCOS, ISSUES E OPORTUNIDADES"
    ws["A1"].font = Font(name=FONTE, size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=PRIMARIA)
    ws.row_dimensions[1].height = 24
    for lin, (rotulo, chave) in enumerate((("Código do projeto", "codigo"), ("Nome do projeto", "titulo"), ("Data", "data"), ("Responsável", "responsavel")), 2):
        ws.cell(lin, 1, rotulo).font = fonte(bold=True)
        ws.merge_cells(start_row=lin, start_column=1, end_row=lin, end_column=2)
        ws.merge_cells(start_row=lin, start_column=3, end_row=lin, end_column=4)
        ws.cell(lin, 3, ident.get(chave) or (A_DEFINIR if preencher else None)).font = fonte()

    cab = LINHA_INICIAL - 1
    for j, (titulo, largura) in enumerate(COLUNAS, 1):
        c = ws.cell(cab, j, titulo)
        c.font, c.border = fonte(bold=True, color="FFFFFF"), BORDA
        c.fill = PatternFill("solid", fgColor=PRIMARIA)
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.column_dimensions[c.column_letter].width = largura
    ws.row_dimensions[cab].height = 32

    itens, prox = numerar(d.get("itens", []) if preencher else [])
    for i in range(len(itens) + LINHAS_LIVRES):
        lin = LINHA_INICIAL + i
        item = itens[i] if i < len(itens) else None
        ident_item = item["id"] if item else f"RISC{prox + i - len(itens):03d}"
        vazio = A_DEFINIR if item else None
        pegar = (lambda k: item.get(k) if item.get(k) is not None else vazio) if item else (lambda k: None)
        nota = f"E{lin}*F{lin}"
        valores = [ident_item, pegar("descricao"), pegar("tipo"), pegar("consequencia"),
                   pegar("consequencia_nivel"), pegar("probabilidade"),
                   f'=IF(COUNT(E{lin},F{lin})=2,{formula_nota(nota)},IF(B{lin}="","","{A_DEFINIR}"))',
                   pegar("medida"), pegar("plano_acao"), pegar("custo"), pegar("status"), item.get("observacoes") if item else None]
        for j, v in enumerate(valores, 1):
            c = ws.cell(lin, j, v)
            c.font, c.border = fonte(), BORDA
            c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(lin, 10).number_format = MOEDA
        if item:
            textos = [(item.get("descricao"), 42), (item.get("consequencia"), 34), (item.get("plano_acao"), 42), (item.get("observacoes"), 30)]
            ws.row_dimensions[lin].height = 14 * max(1, *[len(t or "") // l + 1 for t, l in textos]) + 2
    fim = LINHA_INICIAL + len(itens) + LINHAS_LIVRES - 1

    for coluna, opcoes in (("C", TIPOS), ("E", ESCALA), ("F", ESCALA), ("H", MEDIDAS), ("K", STATUS)):
        dv = DataValidation(type="list", formula1='"' + ",".join(opcoes) + '"', allow_blank=True)
        dv.add(f"{coluna}{LINHA_INICIAL}:{coluna}{fim}")
        ws.add_data_validation(dv)
    for nome, cor in CORES.items():
        ws.conditional_formatting.add(f"G{LINHA_INICIAL}:G{fim}", CellIsRule(
            operator="equal", formula=[f'"{nome}"'], fill=PatternFill("solid", bgColor=cor)))
    ws.freeze_panes = ws.cell(LINHA_INICIAL, 3)
    ws.auto_filter.ref = f"A{cab}:L{fim}"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def aba_grau(ws):
    ws.merge_cells("A1:F1")
    ws["A1"] = "GRAU = CONSEQUÊNCIA x PROBABILIDADE"
    ws["A1"].font = fonte(bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=PRIMARIA)
    ws["A2"] = "Consequência / Probabilidade"
    ws["A2"].font, ws["A2"].alignment = fonte(bold=True), Alignment(wrap_text=True)
    ws.column_dimensions["A"].width = 18
    for n in range(1, 6):
        ws.column_dimensions[chr(65 + n)].width = 18
        for c in (ws.cell(2, n + 1, n), ws.cell(n + 2, 1, n)):
            c.font, c.fill, c.border = fonte(bold=True), PatternFill("solid", fgColor=CLARA), BORDA
    for cons in range(1, 6):
        for prob in range(1, 6):
            rotulo = grau(cons * prob)
            c = ws.cell(cons + 2, prob + 1, f"{cons * prob} - {rotulo}")
            c.font, c.border = fonte(), BORDA
            c.fill = PatternFill("solid", fgColor=CORES[rotulo])


def gerar(d, destino, preencher):
    wb = Workbook()
    wb.active.title = "Matriz"
    aba_matriz(wb.active, d, preencher)
    aba_grau(wb.create_sheet("Grau"))
    wb.properties.creator = wb.properties.lastModifiedBy = "pmo-ai-skills"
    wb.properties.title = "Matriz de Riscos, Issues e Oportunidades"
    wb.calculation.fullCalcOnLoad = True
    wb.save(destino)


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
    itens = dados.get("itens", [])
    graus = Counter(
        grau(i["consequencia_nivel"] * i["probabilidade"]) if i.get("consequencia_nivel") and i.get("probabilidade") else A_DEFINIR
        for i in itens)
    resumo = ", ".join(f"{n}: {graus[n]}" for n in [*CORES, A_DEFINIR] if graus[n])
    print(f"Matriz: {sys.argv[2]} | {len(itens)} itens | {resumo or 'sem itens'}")


if __name__ == "__main__":
    main()
