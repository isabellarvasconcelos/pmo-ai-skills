---
name: status-report
description: Gera o Status Report de um projeto em um slide PPTX, com farol, % previsto x realizado, Curva S, próximas atividades, atividades atrasadas e riscos. Use quando o usuário pedir o status report periódico de um projeto.
---

# Status Report

**Fase do ciclo de vida:** monitoramento e controle.

## Fluxo

1. Reúna os dados do período: % previsto e realizado, Curva S, próximas atividades, atividades atrasadas e riscos ativos.
2. Se houver uma matriz de riscos, use os itens abertos de grau "Significativo" e "Importante".
3. Pergunte a criticidade do projeto se ela não tiver sido informada, porque ela define o farol e a periodicidade.
4. Monte um JSON com a estrutura de `examples/entrada.json`. Omita o que não tiver sido informado.
5. Execute `python scripts/gerar_status.py entrada.json saida.pptx`.
6. Entregue o arquivo, informe o farol e o desvio e, se o farol estiver vermelho, diga quem deve ser acionado.

## Farol

O desvio é relativo ao previsto: `(previsto - realizado) / previsto`. Um projeto adiantado tem desvio negativo.

| Criticidade | Periodicidade | Limite de escalonamento | Ação acima do limite |
|---|---|---|---|
| BAIXO | Quinzenal | 15% | Acionar o gerente de projetos |
| MÉDIO | Semanal | 10% | Acionar o sponsor e o gerente |
| ALTO | Semanal, com reunião executiva mensal | 5% | Convocar gerentes e diretoria |

- **Verde:** desvio de até metade do limite.
- **Amarelo:** desvio acima da metade do limite, até o limite.
- **Vermelho:** desvio acima do limite, o que exige escalonamento.
- **Cinza:** faltam o previsto, o realizado ou a criticidade.

## Regras

- **Nunca invente informação.** O que não foi informado aparece como "a definir".
- Não escolha a cor do farol à mão. O script calcula a partir dos dados.
- O slide tem espaço limitado. Se houver mais itens do que cabem, o script avisa quantos ficaram de fora. Priorize os mais críticos e cite os demais na resposta.
- Em riscos, `severidade` corresponde ao nível de consequência da matriz de riscos.
- Mantenha um tom direto e executivo.
