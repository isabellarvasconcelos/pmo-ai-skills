---
name: matriz-riscos-issues-oportunidades
description: Gera e atualiza a Matriz de Riscos, Issues e Oportunidades de um projeto em XLSX, calculando o grau de cada item por consequência x probabilidade. Use quando o usuário pedir para registrar, atualizar ou revisar riscos, issues ou oportunidades de um projeto.
---

# Matriz de Riscos, Issues e Oportunidades

**Fase do ciclo de vida:** planejamento, monitoramento e controle (revisão contínua).

## Fluxo

1. Reúna os itens informados pelo usuário. Se houver um TAP, leia as premissas e restrições, porque cada premissa pode gerar um risco.
2. Sugira itens que pareçam faltar, mas só os inclua depois que o usuário confirmar.
3. Peça ao usuário a consequência e a probabilidade de cada item, ou proponha notas para ele confirmar.
4. Monte um JSON com a estrutura de `examples/entrada.json`. Omita o que não tiver sido informado.
5. Execute `python scripts/gerar_matriz.py entrada.json saida.xlsx`.
6. Entregue o arquivo, destaque os itens "Significativo" e liste o que ficou como "a definir".

Em uma atualização, mantenha o `id` dos itens existentes. Itens novos recebem o próximo número livre.

## Classificação

O grau é consequência (1 a 5) x probabilidade (1 a 5):

| Resultado | Grau |
|---|---|
| 1 a 4 | Não Significativo |
| 5 a 12 | Importante |
| 15 a 25 | Significativo |

## Regras

- **Nunca invente informação.** Notas, custos e responsáveis sem dado ficam como "a definir".
- Diferencie os tipos: `RISCO` pode acontecer, `ISSUE` já aconteceu e `OPORTUNIDADE` é um efeito positivo possível.
- Use em `medida`: `Mitigar`, `Compartilhar`, `Transferir`, `Eliminar` ou `Justificar` para riscos e issues; `Explorar`, `Compartilhar`, `Melhorar` ou `Aceitar` para oportunidades.
- Use em `status`: `ABERTO`, `CONCLUÍDO` ou `CANCELADO`.
- Não digite o grau à mão. Ele vem de uma fórmula na planilha.
- Mantenha um tom direto e executivo.
