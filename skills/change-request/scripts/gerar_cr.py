#!/usr/bin/env python3
"""Gera o Change Request em DOCX. Uso: gerar_cr.py [entrada.json] saida.docx (sem entrada: modelo em branco)."""
import json
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

A_DEFINIR = "a definir"
FONTE = "Arial"
PRIMARIA, CLARA, DICA = "1F3A5F", "DCE6F1", "7F7F7F"
ASPECTOS = ["Cronograma", "Requisitos", "Recursos", "Custos", "Escopo", "Técnico", "Funcional"]
DECISOES = ["APROVADO", "REPROVADO"]
LARGURA = 17.0  # cm úteis em A4 com margens de 2 cm


def sombrear(celula, cor):
    tcPr = celula._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), cor)
    tcPr.append(shd)


def escrever(celula, texto, negrito=False, cor=None, dica=False, tamanho=10):
    celula.text = ""
    p = celula.paragraphs[0]
    linhas = texto if isinstance(texto, list) else [texto]
    for i, linha in enumerate(linhas):
        if i:
            p = celula.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(("• " if isinstance(texto, list) else "") + str(linha))
        run.font.name, run.font.size, run.bold, run.italic = FONTE, Pt(tamanho), negrito, dica
        if cor or dica:
            run.font.color.rgb = RGBColor.from_string(cor or DICA)


class Formulario:
    def __init__(self, preencher):
        self.preencher = preencher
        self.doc = Document()
        sec = self.doc.sections[0]
        sec.page_width, sec.page_height = Cm(21), Cm(29.7)
        sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Cm(2)
        estilo = self.doc.styles["Normal"]
        estilo.font.name, estilo.font.size = FONTE, Pt(10)
        estilo.element.rPr.rFonts.set(qn("w:eastAsia"), FONTE)

    def valor(self, v, dica=""):
        """Retorna (texto, é_dica). No modelo em branco, mostra a dica."""
        if not self.preencher:
            return dica, True
        if v in (None, "", []):
            return A_DEFINIR, False
        return v, False

    def tabela(self, linhas, colunas, larguras):
        """Tabela com largura fixa e bordas explícitas, para não depender do estilo de quem abre."""
        t = self.doc.add_table(rows=linhas, cols=colunas)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        tblPr = t._tbl.tblPr
        novos = {}
        for tag, atributos in (("tblW", {"w:w": str(int(sum(larguras) * 567)), "w:type": "dxa"}),
                               ("tblLayout", {"w:type": "fixed"})):
            novos[tag] = OxmlElement(f"w:{tag}")
            for k, v in atributos.items():
                novos[tag].set(qn(k), v)
        novos["tblBorders"] = OxmlElement("w:tblBorders")
        for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
            b = OxmlElement(f"w:{lado}")
            for k, v in (("w:val", "single"), ("w:sz", "4"), ("w:space", "0"), ("w:color", "BFBFBF")):
                b.set(qn(k), v)
            novos["tblBorders"].append(b)
        # O Word exige os filhos de tblPr nesta ordem; fora dela, o arquivo pode ser recusado.
        ordem = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW",
                 "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook"]
        filhos = {el.tag.split("}")[1]: el for el in tblPr}
        filhos.update(novos)
        for el in list(tblPr):
            tblPr.remove(el)
        for tag in ordem:
            if tag in filhos:
                tblPr.append(filhos[tag])
        for gc, w in zip(t._tbl.tblGrid.findall(qn("w:gridCol")), larguras):
            gc.set(qn("w:w"), str(int(w * 567)))
        for j, w in enumerate(larguras):
            for celula in t.columns[j].cells:
                celula.width = Cm(w)
        return t

    def titulo_tabela(self, t, texto):
        linha = t.rows[0].cells
        celula = linha[0].merge(linha[-1]) if len(linha) > 1 else linha[0]
        escrever(celula, texto, negrito=True, cor="FFFFFF")
        sombrear(celula, PRIMARIA)

    def cabecalho(self, celulas):
        for c in celulas:
            sombrear(c, CLARA)
            for p in c.paragraphs:
                for r in p.runs:
                    r.bold = True

    def espaco(self):
        self.doc.add_paragraph().paragraph_format.space_after = Pt(2)

    def bloco_texto(self, titulo, valor, dica):
        t = self.tabela(2, 1, [LARGURA])
        self.titulo_tabela(t, titulo)
        texto, e_dica = self.valor(valor, dica)
        escrever(t.cell(1, 0), texto, dica=e_dica)
        self.espaco()


def gerar(d, destino, preencher):
    f = Formulario(preencher)
    doc = f.doc

    titulo = doc.add_paragraph()
    titulo.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = titulo.add_run("REQUISIÇÃO DE MUDANÇA (CHANGE REQUEST)")
    r.font.name, r.font.size, r.bold = FONTE, Pt(16), True
    r.font.color.rgb = RGBColor.from_string(PRIMARIA)

    ident = d.get("identificacao", {})
    campos = [("Código do projeto", "codigo"), ("Nome do projeto", "titulo"), ("Código da mudança", "codigo_mudanca"),
              ("Solicitante", "solicitante"), ("Data da solicitação", "data_solicitacao")]
    t = f.tabela(len(campos), 2, [5.0, LARGURA - 5.0])
    for i, (rotulo, chave) in enumerate(campos):
        escrever(t.cell(i, 0), rotulo, negrito=True)
        sombrear(t.cell(i, 0), CLARA)
        texto, e_dica = f.valor(ident.get(chave))
        escrever(t.cell(i, 1), texto, dica=e_dica)
    f.espaco()

    p = doc.add_paragraph()
    p.add_run("Este documento formaliza a solicitação de mudança do projeto e a avaliação dos impactos que ela causará.").italic = True

    f.bloco_texto("Descrição da mudança", d.get("descricao"), "O que muda, em detalhe, e em qual parte do projeto.")
    f.bloco_texto("Justificativa", d.get("justificativa"), "Por que a mudança é necessária e o que acontece se ela não for feita.")

    impactos = {i["aspecto"]: i for i in d.get("impactos", [])}
    t = f.tabela(len(ASPECTOS) + 2, 4, [3.0, 2.2, 7.8, 4.0])
    f.titulo_tabela(t, "Análise de impacto da mudança")
    for j, h in enumerate(["Aspecto", "Impactado?", "Justificativa do impacto", "Documento de referência"]):
        escrever(t.cell(1, j), h)
    f.cabecalho(t.rows[1].cells)
    for i, aspecto in enumerate(ASPECTOS, 2):
        item = impactos.get(aspecto)
        escrever(t.cell(i, 0), aspecto, negrito=True)
        if not preencher:
            escrever(t.cell(i, 1), "☐ Sim\n☐ Não")
            continue
        escrever(t.cell(i, 1), "Sim" if item else "Não", negrito=bool(item))
        escrever(t.cell(i, 2), item["justificativa"] if item else "Sem impacto identificado.", dica=not item)
        escrever(t.cell(i, 3), (item or {}).get("documento") or ("a definir" if item else "-"))
    f.espaco()

    resumo = d.get("resumo_impacto", {})
    t = f.tabela(2, 2, [LARGURA / 2] * 2)
    for j, (rotulo, chave, fmt) in enumerate((("Impacto em prazo", "prazo_dias", lambda v: f"{v:+d} dias corridos"),
                                              ("Impacto em custo", "custo", lambda v: "R$ " + f"{v:+,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")))):
        escrever(t.cell(0, j), rotulo, negrito=True)
        sombrear(t.cell(0, j), CLARA)
        v = resumo.get(chave)
        texto, e_dica = f.valor(fmt(v) if v is not None else None)
        escrever(t.cell(1, j), texto, dica=e_dica, negrito=not e_dica and v is not None, tamanho=12)
    f.espaco()

    riscos = d.get("riscos_impactos", {})
    t = f.tabela(3, 2, [5.0, LARGURA - 5.0])
    f.titulo_tabela(t, "Riscos e impactos")
    for i, (rotulo, chave, dica) in enumerate((("Riscos da mudança", "riscos", "O que pode dar errado ao executar a mudança."),
                                                ("Impactos de não fazer", "impactos", "O que acontece se a mudança for reprovada.")), 1):
        escrever(t.cell(i, 0), rotulo, negrito=True)
        sombrear(t.cell(i, 0), CLARA)
        texto, e_dica = f.valor(riscos.get(chave), dica)
        escrever(t.cell(i, 1), texto, dica=e_dica)
    f.espaco()

    plano = d.get("plano_execucao", []) if preencher else []
    linhas = max(len(plano), 1 if preencher else 5)
    t = f.tabela(linhas + 2, 5, [1.6, 7.0, 3.4, 2.4, 2.6])
    f.titulo_tabela(t, "Plano de execução da mudança")
    for j, h in enumerate(["Seq.", "Atividade", "Responsável", "Data", "Duração prevista"]):
        escrever(t.cell(1, j), h)
    f.cabecalho(t.rows[1].cells)
    for i in range(linhas):
        item = plano[i] if i < len(plano) else {}
        escrever(t.cell(i + 2, 0), str(i + 1))
        for j, chave in enumerate(("atividade", "responsavel", "data", "duracao"), 1):
            escrever(t.cell(i + 2, j), item.get(chave) or (A_DEFINIR if preencher else ""))
    f.espaco()

    f.bloco_texto("Observações finais", d.get("observacoes"), "")

    t = f.tabela(2, 1, [LARGURA])
    f.titulo_tabela(t, "Avaliação final")
    decisao = d.get("decisao")
    marcas = "     ".join(f"{'☒' if decisao == op else '☐'} {op}" for op in DECISOES)
    if preencher and not decisao:
        marcas += "     (pendente de avaliação)"
    escrever(t.cell(1, 0), marcas, negrito=True)
    f.espaco()

    papeis = ["Solicitante", "Gerente do projeto", "Sponsor"]
    t = f.tabela(len(papeis) + 1, 3, [4.5, 8.0, 4.5])
    for j, h in enumerate(["Papel", "Nome e assinatura", "Data"]):
        escrever(t.cell(0, j), h)
    f.cabecalho(t.rows[0].cells)
    for i, papel in enumerate(papeis, 1):
        escrever(t.cell(i, 0), papel, negrito=True)
        t.rows[i].height = Cm(1.0)

    doc.core_properties.author = doc.core_properties.last_modified_by = "pmo-ai-skills"
    doc.core_properties.title = "Requisição de Mudança"
    doc.save(destino)


def validar(d):
    erros = []
    vistos = set()
    for i, item in enumerate(d.get("impactos", []), 1):
        if item.get("aspecto") not in ASPECTOS:
            erros.append(f"impactos[{i}].aspecto deve ser um destes: {', '.join(ASPECTOS)}")
        elif item["aspecto"] in vistos:
            erros.append(f"impactos[{i}].aspecto repetido: {item['aspecto']}")
        vistos.add(item.get("aspecto"))
        if not item.get("justificativa"):
            erros.append(f"impactos[{i}].justificativa é obrigatória para um aspecto impactado")
    resumo = d.get("resumo_impacto", {})
    if resumo.get("prazo_dias") is not None and type(resumo["prazo_dias"]) is not int:
        erros.append("resumo_impacto.prazo_dias deve ser um número inteiro de dias")
    if resumo.get("custo") is not None and type(resumo["custo"]) not in (int, float):
        erros.append("resumo_impacto.custo deve ser um número")
    if d.get("decisao") is not None and d["decisao"] not in DECISOES:
        erros.append(f"decisao deve ser um destes: {', '.join(DECISOES)}")
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
    aspectos = ", ".join(i["aspecto"] for i in dados.get("impactos", [])) or "nenhum"
    print(f"Change request: {sys.argv[2]} | aspectos impactados: {aspectos} | decisão: {dados.get('decisao') or 'pendente'}")


if __name__ == "__main__":
    main()
