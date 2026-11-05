# PMO AI Skills

Skills de IA que transformam etapas da gestão de projetos em entregáveis prontos para revisão.

Este repositório parte de uma prática real de PMO, reescrita em versão genérica. Todos os exemplos usam dados fictícios.

## Skills

| Skill | Fase | Entrega |
|---|---|---|
| [TAP: Termo de Abertura do Projeto](skills/tap-termo-abertura-projeto/SKILL.md) | Iniciação | Planilha XLSX com o termo e as aquisições |
| [Matriz de Riscos, Issues e Oportunidades](skills/matriz-riscos-issues-oportunidades/SKILL.md) | Planejamento, monitoramento e controle | Planilha XLSX com o registro dos itens e o grau calculado |
| [Status Report](skills/status-report/SKILL.md) | Monitoramento e controle | Slide PPTX com farol, Curva S, atividades e riscos |

## Como testar

Requer Python 3.9 ou mais recente.

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

cd skills/tap-termo-abertura-projeto
python scripts/gerar_tap.py examples/entrada.json meu_tap.xlsx
```

Cada skill tem seu próprio script, além de uma entrada de exemplo e o resultado gerado em `examples/`.

No TAP, veja a [entrada de exemplo](skills/tap-termo-abertura-projeto/examples/entrada.json), o [resultado gerado](skills/tap-termo-abertura-projeto/examples/saida.xlsx) e o [modelo em branco](skills/tap-termo-abertura-projeto/assets/modelo_tap.xlsx).

## Princípios

- **Não inventar dados:** o que não foi informado fica como "a definir".
- **Manter a pessoa no controle:** a IA prepara o documento, e a revisão e a aprovação continuam sendo humanas.
- **Validar o que dá para validar:** valores e tipos inválidos são recusados, e os totais vêm de fórmulas.
