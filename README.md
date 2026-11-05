# PMO AI Skills

Skills de IA que transformam etapas da gestão de projetos em entregáveis prontos para revisão.

Este repositório parte de uma prática real de PMO, reescrita em versão genérica. Todos os exemplos usam dados fictícios.

## Skills

| Skill | Fase | Entrega |
|---|---|---|
| [Grill Me: Projetos](skills/grill-me-projetos/SKILL.md) | Antes da iniciação | Entrevista de descoberta que termina em um PRD |
| [TAP: Termo de Abertura do Projeto](skills/tap-termo-abertura-projeto/SKILL.md) | Iniciação | Planilha XLSX com o termo e as aquisições |
| [Matriz de Riscos, Issues e Oportunidades](skills/matriz-riscos-issues-oportunidades/SKILL.md) | Planejamento, monitoramento e controle | Planilha XLSX com o registro dos itens e o grau calculado |
| [Status Report](skills/status-report/SKILL.md) | Monitoramento e controle | Slide PPTX com farol, Curva S, atividades e riscos |
| [Change Request](skills/change-request/SKILL.md) | Monitoramento e controle | Documento DOCX com a mudança, a análise de impacto e a aprovação |
| [Ata de Reunião](skills/ata-reuniao/SKILL.md) | Todas | Documento DOCX com decisões e itens de ação, inclusive a partir de anotações soltas |
| [Termo de Encerramento](skills/termo-encerramento-projeto/SKILL.md) | Encerramento | Documento DOCX com aprendizados, previsto x realizado e status do escopo |

## Como testar

Requer Python 3.9 ou mais recente.

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

cd skills/tap-termo-abertura-projeto
python scripts/gerar_tap.py examples/entrada.json meu_tap.xlsx
```

As skills que geram arquivos têm seu próprio script, além de uma entrada de exemplo e o resultado gerado em `examples/`. A Grill Me é só uma conversa: veja a [sessão de exemplo](skills/grill-me-projetos/examples/sessao-checkin-digital.md).

No TAP, veja a [entrada de exemplo](skills/tap-termo-abertura-projeto/examples/entrada.json), o [resultado gerado](skills/tap-termo-abertura-projeto/examples/saida.xlsx) e o [modelo em branco](skills/tap-termo-abertura-projeto/assets/modelo_tap.xlsx).

## Princípios

- **Não inventar dados:** o que não foi informado fica como "a definir".
- **Manter a pessoa no controle:** a IA prepara o documento, e a revisão e a aprovação continuam sendo humanas.
- **Validar o que dá para validar:** valores e tipos inválidos são recusados, e os totais vêm de fórmulas.

## Créditos

A skill [Grill Me: Projetos](skills/grill-me-projetos/SKILL.md) adapta a [grill-me](https://github.com/mattpocock/skills/tree/main/skills/productivity/grill-me) de Matt Pocock (MIT), a partir de uma versão de Tiago Ornelas. Veja os [créditos completos](skills/grill-me-projetos/CREDITOS.md).

## Licença

[MIT](LICENSE). Você pode usar, adaptar e redistribuir, mantendo o aviso de copyright. Trechos adaptados de terceiros mantêm também a licença original, indicada nos créditos.
