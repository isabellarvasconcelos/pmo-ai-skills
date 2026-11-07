# PMO AI Skills

[English](README.md) | **Português**

Skills de IA que transformam etapas da gestão de projetos em entregáveis prontos para revisão.

Este repositório parte de uma prática real de PMO, reescrita em versão genérica. Todos os exemplos usam dados fictícios.

## Skills

| Skill | Fase | Entrega |
|---|---|---|
| [Grill Me: Projetos](skills/grill-me-projects/SKILL.md) | Antes da iniciação | Entrevista de descoberta que termina em um PRD |
| [TAP: Termo de Abertura do Projeto](skills/project-charter/SKILL.md) | Iniciação | Planilha XLSX com o termo e as aquisições |
| [Matriz de Riscos, Issues e Oportunidades](skills/risk-issue-opportunity-matrix/SKILL.md) | Planejamento, monitoramento e controle | Planilha XLSX com o registro dos itens e o grau calculado |
| [Status Report](skills/status-report/SKILL.md) | Monitoramento e controle | Slide PPTX com farol, Curva S, atividades e riscos |
| [Change Request](skills/change-request/SKILL.md) | Monitoramento e controle | Documento DOCX com a mudança, a análise de impacto e a aprovação |
| [Ata de Reunião](skills/meeting-minutes/SKILL.md) | Todas | Documento DOCX com decisões e itens de ação, inclusive a partir de anotações soltas |
| [Termo de Encerramento](skills/project-closure/SKILL.md) | Encerramento | Documento DOCX com aprendizados, previsto x realizado e status do escopo |
| [Lições Aprendidas](skills/lessons-learned/SKILL.md) | Monitoramento e controle, encerramento | Planilha XLSX com pontos positivos e negativos e planos de ação |

As instruções das skills estão em inglês, e os documentos podem ser gerados em inglês ou em português.

## Como testar

Requer Python 3.9 ou mais recente.

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

cd skills/project-charter
python scripts/generate_charter.py examples/input.pt-BR.json meu_tap.xlsx
```

Para gerar em português, inclua `"language": "pt-BR"` na entrada ou use `--lang pt-BR`. Sem isso, o documento sai em inglês. O modelo em branco também pode ser gerado nos dois idiomas:

```bash
python scripts/generate_charter.py --lang pt-BR tap_em_branco.xlsx
```

As skills que geram arquivos têm seu próprio script, além de entradas de exemplo e os resultados gerados em `examples/`, em inglês e em português (`*.pt-BR.*`). A Grill Me é só uma conversa: veja a [sessão de exemplo](skills/grill-me-projects/examples/discovery-session.pt-BR.md).

No TAP, veja a [entrada de exemplo](skills/project-charter/examples/input.pt-BR.json), o [resultado gerado](skills/project-charter/examples/output.pt-BR.xlsx) e o [modelo em branco](skills/project-charter/assets/charter_template.pt-BR.xlsx).

Todas as skills usam o mesmo projeto fictício, um check-in digital de entregas, para que os números batam de um documento para o outro.

## Princípios

- **Não inventar dados:** o que não foi informado fica como "a definir".
- **Manter a pessoa no controle:** a IA prepara o documento, e a revisão e a aprovação continuam sendo humanas.
- **Validar o que dá para validar:** valores e tipos inválidos são recusados, e os totais vêm de fórmulas.

## Créditos

A skill [Grill Me: Projetos](skills/grill-me-projects/SKILL.md) adapta a [grill-me](https://github.com/mattpocock/skills/tree/main/skills/productivity/grill-me) de Matt Pocock (MIT), a partir de uma versão de Tiago Ornelas. Veja os [créditos completos](skills/grill-me-projects/CREDITS.md).

## Licença

[MIT](LICENSE). Você pode usar, adaptar e redistribuir, mantendo o aviso de copyright. Trechos adaptados de terceiros mantêm também a licença original, indicada nos créditos.
