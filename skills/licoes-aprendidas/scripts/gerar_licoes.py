#!/usr/bin/env python3
"""Gera a matriz de lições aprendidas em XLSX. Uso: gerar_licoes.py [entrada.json] saida.xlsx (sem entrada: modelo em branco)."""
import json
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

A_DEFINIR = "a definir"
PRIMARIA, CLARA = "1F3A5F", "DCE6F1"
FONTE = "Arial"
BORDA = Border(*[Side(style="thin", color="BFBFBF")] * 4)
TIPOS = ["POSITIVO", "NEGATIVO"]
STATUS = ["ABERTO", "EM ANDAMENTO", "CONCLUÍDO"]
MOMENTOS = ["DURANTE O PROJETO", "NO ENCERRAMENTO", "PRÓXIMOS PROJETOS"]
CORES = {"POSITIVO": "C6E0B4", "NEGATIVO": "F4B0A5"}
LINHA_INICIAL = 10
LINHAS_LIVRES = 15
COLUNAS = [("ID", "id", 9), ("Data", "data", 11), ("Positivo / negativo", "tipo", 12), ("Detalhamento", "detalhamento", 44),
           ("Área relacionada", "area", 15), ("Plano de ação", "plano_acao", 40), ("Responsável", "responsavel", 17),
           ("Data de encerramento", "data_encerramento", 13), ("Status", "status", 14), ("Momento da tratativa", "momento", 19),
           ("Observações", "observacoes", 28)]
OPCIONAIS = {"data_encerramento", "observacoes"}  # vazios não viram "a definir"


def fonte(**k):
    return Font(name=FONTE, size=10, **k)


def validar(d):
    erros, vistos = [], set()
    for i, item in enumerate(d.get("licoes", []), 1):
        rotulo = f"licoes[{i}]"
        if not item.get("detalhamento"):
            erros.append(f"{rotulo}.detalhamento é obrigatório")
        for campo, validos in (("tipo", TIPOS), ("status", STATUS), ("momento", MOMENTOS)):
            if item.get(campo) is not None and item[campo] not in validos:
                erros.append(f"{rotulo}.{campo} deve ser um destes: {', '.join(validos)}")
        for campo in ("data", "data_encerramento"):
            if item.get(campo) is not None:
                try:
                    datetime.strptime(item[campo], "%d/%m/%Y")
                except (TypeError, ValueError):
                    erros.append(f"{rotulo}.{campo} deve estar no formato dd/mm/aaaa")
        if item.get("data_encerramento") and item.get("status") != "CONCLUÍDO":
            erros.append(f"{rotulo}.data_encerramento só vale para status CONCLUÍDO")
        if item.get("id"):
            if item["id"] in vistos:
                erros.append(f"{rotulo}.id repetido: {item['id']}")
            vistos.add(item["id"])
    if erros:
        sys.exit("Entrada inválida:\n- " + "\n- ".join(erros))


def numerar(itens):
    """Mantém os IDs informados e dá aos novos o próximo número livre, sem reaproveitar lacunas."""
    numeros = [int(i["id"][3:]) for i in itens if str(i.get("id", "")).startswith("LIC") and i["id"][3:].isdigit()]
    prox = max(numeros, default=0) + 1
    saida = []
    for item in itens:
        item = dict(item)
        if not item.get("id"):
            item["id"] = f"LIC{prox:03d}"
            prox += 1
        saida.append(item)
    return saida, prox


def gerar(d, destino, preencher):
    wb = Workbook()
    ws = wb.active
    ws.title = "Lições aprendidas"
    ultima_col = chr(64 + len(COLUNAS))
    ws.merge_cells(f"A1:{ultima_col}1")
    ws["A1"] = "LIÇÕES APRENDIDAS"
    ws["A1"].font = Font(name=FONTE, size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=PRIMARIA)
    ws.row_dimensions[1].height = 24

    ident = d.get("identificacao", {})
    for lin, (rotulo, chave) in enumerate((("Código do projeto", "codigo"), ("Nome do projeto", "titulo"),
                                           ("Data", "data"), ("Responsável", "responsavel")), 2):
        ws.merge_cells(start_row=lin, start_column=1, end_row=lin, end_column=2)
        ws.merge_cells(start_row=lin, start_column=3, end_row=lin, end_column=4)
        ws.cell(lin, 1, rotulo).font = fonte(bold=True)
        ws.cell(lin, 3, ident.get(chave) or (A_DEFINIR if preencher else None)).font = fonte()

    itens, prox = numerar(d.get("licoes", []) if preencher else [])
    fim = LINHA_INICIAL + len(itens) + LINHAS_LIVRES - 1
    resumo = [("Positivas", f'=COUNTIF(C{LINHA_INICIAL}:C{fim},"POSITIVO")', sum(i.get("tipo") == "POSITIVO" for i in itens)),
              ("Negativas", f'=COUNTIF(C{LINHA_INICIAL}:C{fim},"NEGATIVO")', sum(i.get("tipo") == "NEGATIVO" for i in itens)),
              ("Ações abertas", f'=COUNTIF(I{LINHA_INICIAL}:I{fim},"ABERTO")+COUNTIF(I{LINHA_INICIAL}:I{fim},"EM ANDAMENTO")',
               sum(i.get("status") in ("ABERTO", "EM ANDAMENTO") for i in itens))]
    calculados = {}
    for lin, (rotulo, formula, valor) in enumerate(resumo, 2):
        ws.cell(lin, 5, rotulo).font = fonte(bold=True)
        c = ws.cell(lin, 6, formula)
        c.alignment = Alignment(horizontal="left")
        c.font, c.fill, c.border = fonte(bold=True), PatternFill("solid", fgColor=CLARA), BORDA
        calculados[c.coordinate] = valor

    cab = LINHA_INICIAL - 1
    for j, (titulo, _, largura) in enumerate(COLUNAS, 1):
        c = ws.cell(cab, j, titulo)
        c.font, c.border = fonte(bold=True, color="FFFFFF"), BORDA
        c.fill = PatternFill("solid", fgColor=PRIMARIA)
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.column_dimensions[c.column_letter].width = largura
    ws.row_dimensions[cab].height = 32

    for i in range(len(itens) + LINHAS_LIVRES):
        lin = LINHA_INICIAL + i
        item = itens[i] if i < len(itens) else None
        for j, (_, chave, largura) in enumerate(COLUNAS, 1):
            if chave == "id":
                valor = item["id"] if item else f"LIC{prox + i - len(itens):03d}"
            elif item is None:
                valor = None
            else:
                valor = item.get(chave) or (None if chave in OPCIONAIS else A_DEFINIR)
            c = ws.cell(lin, j, valor)
            c.font, c.border = fonte(), BORDA
            c.alignment = Alignment(wrap_text=True, vertical="top")
        if item:
            textos = [(item.get(k) or "", w) for _, k, w in COLUNAS if k in ("detalhamento", "plano_acao", "observacoes")]
            ws.row_dimensions[lin].height = 14 * max(len(t) // w + 1 for t, w in textos) + 2

    for coluna, opcoes in (("C", TIPOS), ("I", STATUS), ("J", MOMENTOS)):
        dv = DataValidation(type="list", formula1='"' + ",".join(opcoes) + '"', allow_blank=True)
        dv.add(f"{coluna}{LINHA_INICIAL}:{coluna}{fim}")
        ws.add_data_validation(dv)
    for nome, cor in CORES.items():
        ws.conditional_formatting.add(f"C{LINHA_INICIAL}:C{fim}", CellIsRule(
            operator="equal", formula=[f'"{nome}"'], fill=PatternFill("solid", bgColor=cor)))
    ws.freeze_panes = ws.cell(LINHA_INICIAL, 4)
    ws.auto_filter.ref = f"A{cab}:{ultima_col}{fim}"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    wb.properties.creator = wb.properties.lastModifiedBy = "pmo-ai-skills"
    wb.properties.title = "Lições Aprendidas"
    wb.calculation.fullCalcOnLoad = True
    wb.save(destino)
    gravar_valores(destino, {1: calculados})
    return itens


def gravar_valores(destino, valores):
    """Grava os resultados já calculados ao lado das fórmulas.

    O Excel recalcula ao abrir, mas pré-visualizações (e-mail, Quick Look, navegadores)
    só mostram o valor salvo. valores: {número da aba: {"G10": valor}}.
    """
    import re
    import zipfile
    from xml.sax.saxutils import escape

    with zipfile.ZipFile(destino) as z:
        arquivos = {n: z.read(n) for n in z.namelist()}
    for aba, celulas in valores.items():
        nome = f"xl/worksheets/sheet{aba}.xml"
        xml = arquivos[nome].decode("utf-8")
        for ref, valor in celulas.items():
            if valor is None or valor == "":
                continue
            tipo = ' t="str"' if isinstance(valor, str) else ""
            conteudo = escape(valor) if isinstance(valor, str) else repr(float(valor))
            padrao = re.compile(rf'<c r="{ref}"( s="\d+")?><f>(.*?)</f>(?:<v\s*/>|<v></v>)</c>')
            xml, n = padrao.subn(lambda m: f'<c r="{ref}"{m.group(1) or ""}{tipo}><f>{m.group(2)}</f><v>{conteudo}</v></c>', xml)
            if n != 1:
                raise RuntimeError(f"Fórmula não encontrada em {ref} da aba {aba}")
        arquivos[nome] = xml.encode("utf-8")
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as z:
        for n, conteudo in arquivos.items():
            z.writestr(n, conteudo)


def main():
    if len(sys.argv) not in (2, 3):
        sys.exit(__doc__)
    if len(sys.argv) == 2:
        gerar({}, sys.argv[1], preencher=False)
        print(f"Modelo em branco: {sys.argv[1]}")
        return
    dados = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    validar(dados)
    itens = gerar(dados, sys.argv[2], preencher=True)
    tipos = Counter(i.get("tipo") or A_DEFINIR for i in itens)
    abertas = [i["id"] for i in itens if i.get("status") in ("ABERTO", "EM ANDAMENTO")]
    print(f"Lições: {sys.argv[2]} | {len(itens)} lições | positivas: {tipos['POSITIVO']}, negativas: {tipos['NEGATIVO']}"
          + (f", sem tipo: {tipos[A_DEFINIR]}" if tipos[A_DEFINIR] else "")
          + f" | ações abertas: {', '.join(abertas) or 'nenhuma'}")


if __name__ == "__main__":
    main()
