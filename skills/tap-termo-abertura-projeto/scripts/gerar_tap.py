#!/usr/bin/env python3
"""Gera um TAP em XLSX. Uso: gerar_tap.py [entrada.json] saida.xlsx (sem entrada: modelo em branco)."""
import json
import math
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col

A_DEFINIR = "a definir"
PRIMARIA, CLARA, DICA = "1F3A5F", "DCE6F1", "7F7F7F"
FONTE = "Arial"
LARGURA_COLUNA = 16
COLUNAS = 8
BORDA = Border(*[Side(style="thin", color="BFBFBF")] * 4)
MOEDA = '"R$" #,##0.00'
AREAS = ["Operações", "Qualidade", "Planejamento", "Logística", "Segurança", "Fiscal",
         "Tecnologia", "Compras", "Comercial", "Financeiro", "Pessoas", "Outros"]
STATUS = ["Aprovado", "Reprovado", "Em análise"]


class Folha:
    def __init__(self, ws, preencher):
        self.ws, self.preencher, self.linha = ws, preencher, 1
        for i in range(1, COLUNAS + 1):
            ws.column_dimensions[col(i)].width = LARGURA_COLUNA

    def celula(self, ini, span, valor, negrito=False, cor=None, fundo=None, dica=False, formato=None):
        fim = ini + span - 1
        for c in range(ini, fim + 1):
            self.ws.cell(self.linha, c).border = BORDA
        if span > 1:
            self.ws.merge_cells(start_row=self.linha, start_column=ini, end_row=self.linha, end_column=fim)
        cel = self.ws.cell(self.linha, ini, valor)
        cel.font = Font(name=FONTE, size=10, bold=negrito, italic=dica, color=DICA if dica else cor)
        cel.alignment = Alignment(wrap_text=True, vertical="top")
        if fundo:
            cel.fill = PatternFill("solid", fgColor=fundo)
        if formato:
            cel.number_format = formato
        return cel

    def altura(self, celulas):
        """Células mescladas não ajustam a altura sozinhas, então estimamos pelo texto."""
        linhas = 1
        for texto, span in celulas:
            if isinstance(texto, str) and not texto.startswith("="):
                largura = span * LARGURA_COLUNA * 1.05
                total = sum(max(1, math.ceil(len(p) / largura)) for p in texto.split("\n"))
                linhas = max(linhas, total)
        self.ws.row_dimensions[self.linha].height = 14 * linhas + 2

    def secao(self, texto):
        self.celula(1, COLUNAS, texto, negrito=True, cor="FFFFFF", fundo=PRIMARIA)
        self.linha += 1

    def rotulos(self, itens):
        ini = 1
        for texto, span in itens:
            self.celula(ini, span, texto, negrito=True, fundo=CLARA)
            ini += span
        self.altura(itens)
        self.linha += 1

    def texto(self, valor, dica=""):
        self.campos([(valor, COLUNAS, dica)])

    def campos(self, itens):
        """itens: (valor, span, dica). No modelo em branco, mostra a dica no lugar do valor."""
        ini, medidas = 1, []
        for valor, span, dica in itens:
            if isinstance(valor, list):
                valor = "\n".join(f"• {v}" for v in valor)
            if not self.preencher:
                self.celula(ini, span, dica or None, dica=True)
                medidas.append((dica, span))
            else:
                valor = valor if valor not in (None, "", []) else A_DEFINIR
                self.celula(ini, span, valor)
                medidas.append((valor, span))
            ini += span
        self.altura(medidas)
        self.linha += 1

    def tabela(self, colunas, linhas, minimo=3):
        """colunas: (título, span, chave)."""
        self.rotulos([(t, s) for t, s, _ in colunas])
        if not self.preencher:
            linhas = [{}] * minimo
        elif not linhas:
            linhas = [{}]
        for item in linhas:
            ini, medidas = 1, []
            for _, span, chave in colunas:
                valor = item.get(chave) or (A_DEFINIR if self.preencher else None)
                self.celula(ini, span, valor)
                medidas.append((valor, span))
                ini += span
            self.altura(medidas)
            self.linha += 1

    def marcadores(self, opcoes, marcadas, por_linha=4):
        span = COLUNAS // por_linha
        for i in range(0, len(opcoes), por_linha):
            for j, opcao in enumerate(opcoes[i:i + por_linha]):
                marca = "☒" if opcao in marcadas else "☐"
                self.celula(1 + j * span, span, f"{marca}  {opcao}")
            self.linha += 1


def validar(d):
    erros = []
    for lista, campo in (("custos", "valor"), ("aquisicoes", "valor_unitario"), ("aquisicoes", "quantidade")):
        for i, item in enumerate(d.get(lista, []), 1):
            v = item.get(campo)
            if v is not None and (isinstance(v, bool) or not isinstance(v, (int, float)) or v < 0):
                erros.append(f"{lista}[{i}].{campo} deve ser um número não negativo")
    for lista in ("custos", "aquisicoes"):
        for i, item in enumerate(d.get(lista, []), 1):
            if item.get("tipo") not in (None, "CAPEX", "OPEX"):
                erros.append(f"{lista}[{i}].tipo deve ser CAPEX ou OPEX")
    for chave, validas in (("impactos", AREAS), ("status", STATUS)):
        valores = d.get(chave) or []
        for v in [valores] if isinstance(valores, str) else valores:
            if v not in validas:
                erros.append(f"{chave}: valor '{v}' inválido. Use: {', '.join(validas)}")
    if erros:
        sys.exit("Entrada inválida:\n- " + "\n- ".join(erros))


def aba_tap(ws, d, preencher, calculados):
    f = Folha(ws, preencher)
    ident = d.get("identificacao", {})
    f.celula(1, COLUNAS, "TERMO DE ABERTURA DO PROJETO", negrito=True, cor="FFFFFF", fundo=PRIMARIA).font = Font(
        name=FONTE, size=14, bold=True, color="FFFFFF")
    ws.row_dimensions[1].height = 24
    f.linha += 1
    f.secao("IDENTIFICAÇÃO DO PROJETO")
    for grupo in (
        (("Título do projeto", "titulo"), ("Código", "codigo"), ("Data de abertura", "data_abertura"), ("Duração prevista", "duracao")),
        (("Gerente do projeto", "gestor"), ("Sponsor", "sponsor"), ("Área responsável", "area"), ("PMO", "pmo")),
    ):
        f.rotulos([(rotulo, 2) for rotulo, _ in grupo])
        f.campos([(ident.get(chave), 2, "") for _, chave in grupo])

    f.secao("1. JUSTIFICATIVA E OBJETIVO")
    f.rotulos([("Justificativa (business case)", COLUNAS)])
    f.texto(d.get("justificativa"), "Problema, oportunidade ou exigência que motiva o projeto.")
    f.rotulos([("Objetivo (SMART)", COLUNAS)])
    f.texto(d.get("objetivo"), "Verbo no infinitivo + resultado mensurável + prazo.")

    f.secao("2. PROCESSO ATUAL x PROCESSO FUTURO")
    f.rotulos([("Processo atual (AS IS)", 4), ("Processo futuro (TO BE)", 4)])
    f.campos([(d.get("as_is"), 4, "Como funciona hoje, com dores e limitações."),
              (d.get("to_be"), 4, "Como passará a funcionar após o projeto.")])

    f.secao("3. ESCOPO")
    f.rotulos([("Escopo do produto, serviço ou resultado", COLUNAS)])
    f.texto(d.get("escopo_produto"), "Características do entregável quando pronto.")
    f.rotulos([("Contempla", 4), ("Não contempla", 4)])
    f.campos([(d.get("contempla"), 4, "Principais entregas."), (d.get("nao_contempla"), 4, "O que explicitamente não será feito.")])
    f.rotulos([("Critérios de aceitação", COLUNAS)])
    f.texto(d.get("criterios_aceitacao"), "Condições objetivas para aceitar o entregável.")
    f.rotulos([("Premissas", 4), ("Restrições", 4)])
    f.campos([(d.get("premissas"), 4, "Suposições tidas como verdadeiras."), (d.get("restricoes"), 4, "Limites de orçamento, prazo, tecnologia e recursos.")])

    f.secao("4. REQUISITOS")
    f.tabela([("ID", 1, "id"), ("Requisito funcional", 4, "descricao"), ("Área", 2, "area"), ("Prioridade", 1, "prioridade")], d.get("requisitos_funcionais"))
    f.tabela([("ID", 1, "id"), ("Requisito técnico", 4, "descricao"), ("Responsável", 2, "responsavel"), ("Prioridade", 1, "prioridade")], d.get("requisitos_tecnicos"))
    f.tabela([("ID", 1, "id"), ("Regra de negócio", 7, "descricao")], d.get("regras_negocio"))

    f.secao("5. SOLUÇÃO TÉCNICA")
    sol = d.get("solucao", {})
    f.rotulos([("Descrição da solução", COLUNAS)])
    f.texto(sol.get("descricao"), "O que será implementado, desenvolvido ou configurado.")
    f.rotulos([("Sistemas envolvidos", 3), ("Interfaces", 3), ("Infraestrutura", 2)])
    f.campos([(sol.get("sistemas"), 3, ""), (sol.get("interfaces"), 3, ""), (sol.get("infraestrutura"), 2, "")])
    f.tabela([("Origem", 2, "origem"), ("Destino", 2, "destino"), ("Tipo", 1, "tipo"), ("Descrição da integração", 3, "descricao")], d.get("integracoes"))

    f.secao("6. IMPACTOS OPERACIONAIS")
    f.marcadores(AREAS, d.get("impactos") or [])

    f.secao("7. RECURSOS")
    f.tabela([("Responsável ou recurso", 3, "nome"), ("Área", 2, "area"), ("Papel e alocação", 3, "papel")], d.get("recursos"))

    f.secao("8. ORÇAMENTO")
    f.rotulos([("Categoria de custo (fora de aquisições)", 4), ("Tipo", 2), ("Valor", 2)])
    custos = d.get("custos") or []
    primeira = f.linha
    for item in custos if preencher and custos else [{}] * (3 if not preencher else 1):
        f.celula(1, 4, item.get("categoria") or (A_DEFINIR if preencher else None))
        f.celula(5, 2, item.get("tipo") or (A_DEFINIR if preencher else None))
        v = item.get("valor")
        f.celula(7, 2, v if v is not None else (A_DEFINIR if preencher else None), formato=MOEDA)
        f.linha += 1
    ultima = f.linha - 1
    f.celula(1, 6, "Subtotal", negrito=True, fundo=CLARA)
    f.celula(7, 2, f"=SUM(G{primeira}:G{ultima})", negrito=True, formato=MOEDA)
    subtotal = sum(i.get("valor") or 0 for i in custos) if preencher else 0
    compras = sum((i.get("quantidade") or 0) * (i.get("valor_unitario") or 0) for i in d.get("aquisicoes", [])) if preencher else 0
    calculados.update({f"G{f.linha}": subtotal, f"G{f.linha + 1}": compras, f"G{f.linha + 2}": subtotal + compras})
    f.linha += 1
    f.celula(1, 6, 'Total da aba "Aquisições"', negrito=True, fundo=CLARA)
    f.celula(7, 2, "='Aquisições'!D4", negrito=True, formato=MOEDA)
    f.linha += 1
    f.celula(1, 6, "ORÇAMENTO TOTAL PREVISTO", negrito=True, cor="FFFFFF", fundo=PRIMARIA)
    f.celula(7, 2, f"=G{f.linha - 2}+G{f.linha - 1}", negrito=True, formato=MOEDA)
    f.linha += 1

    f.secao("9. STATUS E APROVAÇÃO")
    status = d.get("status")
    f.marcadores(STATUS, [status], por_linha=3)
    f.tabela([("Papel", 2, "papel"), ("Nome e cargo", 3, "nome"), ("Data", 1, "data"), ("Assinatura", 2, "assinatura")],
             [{"papel": p} for p in ("Gerente do projeto", "Sponsor", "PMO")] if preencher else [{}] * 3)
    if not preencher:
        for i, p in enumerate(("Gerente do projeto", "Sponsor", "PMO")):
            ws.cell(f.linha - 3 + i, 1, p)
    f.rotulos([("Observações", COLUNAS)])
    f.texto(d.get("observacoes"))

    ws.sheet_view.showGridLines = False
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def aba_aquisicoes(ws, d, preencher, calculados):
    larguras = [38, 8, 16, 16, 9, 15, 20, 14]
    for i, w in enumerate(larguras, 1):
        ws.column_dimensions[col(i)].width = w
    titulo = PatternFill("solid", fgColor=PRIMARIA)
    claro = PatternFill("solid", fgColor=CLARA)
    fonte = lambda **k: Font(name=FONTE, size=10, **k)
    ws.merge_cells("A1:H1")
    ws["A1"] = "RELAÇÃO DE AQUISIÇÕES"
    ws["A1"].font, ws["A1"].fill = fonte(bold=True, color="FFFFFF"), titulo
    ident = d.get("identificacao", {})
    ws["A2"], ws["B2"] = "Projeto", ident.get("titulo") or (A_DEFINIR if preencher else None)
    ws["A3"], ws["B3"] = "Código", ident.get("codigo") or (A_DEFINIR if preencher else None)
    ws.merge_cells("B2:H2")
    ws.merge_cells("B3:H3")
    ws["A2"].font = ws["A3"].font = fonte(bold=True)
    ws["B2"].font = ws["B3"].font = fonte()
    ws.merge_cells("A4:C4")
    ws["A4"] = "VALOR TOTAL PREVISTO"
    ws["A4"].font, ws["A4"].fill = fonte(bold=True), claro
    for i, h in enumerate(["Objeto", "Qtd", "Valor unitário", "Valor total", "Tipo", "Necessidade", "Responsável", "Status"], 1):
        c = ws.cell(5, i, h)
        c.font, c.fill, c.border = fonte(bold=True), claro, BORDA
    itens = (d.get("aquisicoes") or []) if preencher else []
    n = max(len(itens), 10)
    for i in range(n):
        lin = 6 + i
        item = itens[i] if i < len(itens) else {}
        def resto(chave):
            return item.get(chave) or (A_DEFINIR if item else None)
        qtd, unit = item.get("quantidade"), item.get("valor_unitario")
        if qtd is not None and unit is not None:
            calculados[f"D{lin}"] = qtd * unit
        valores = [resto("objeto"), qtd, unit, f'=IF(OR(B{lin}="",C{lin}=""),"",B{lin}*C{lin})',
                   resto("tipo"), resto("data_necessidade"), resto("responsavel"), resto("status")]
        for j, v in enumerate(valores, 1):
            c = ws.cell(lin, j, v)
            c.font, c.border = fonte(), BORDA
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if j in (3, 4):
                c.number_format = MOEDA
    ws["D4"] = f"=SUM(D6:D{5 + n})"
    calculados["D4"] = sum(v for k, v in calculados.items() if k != "D4")
    ws["D4"].font, ws["D4"].number_format = fonte(bold=True), MOEDA


def gerar(d, destino, preencher):
    wb = Workbook()
    tap, aquisicoes = {}, {}
    aba_tap(wb.active, d, preencher, tap)
    wb.active.title = "TAP"
    aba_aquisicoes(wb.create_sheet("Aquisições"), d, preencher, aquisicoes)
    wb.properties.creator = "pmo-ai-skills"
    wb.properties.lastModifiedBy = "pmo-ai-skills"
    wb.properties.title = "Termo de Abertura do Projeto"
    wb.calculation.fullCalcOnLoad = True
    wb.save(destino)
    gravar_valores(destino, {1: tap, 2: aquisicoes})


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


def total_previsto(d):
    custos = sum(i.get("valor") or 0 for i in d.get("custos", []))
    compras = sum((i.get("quantidade") or 0) * (i.get("valor_unitario") or 0) for i in d.get("aquisicoes", []))
    return custos + compras


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
    total = f"{total_previsto(dados):,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")
    print(f"TAP: {sys.argv[2]} | orçamento previsto: R$ {total}")


if __name__ == "__main__":
    main()
