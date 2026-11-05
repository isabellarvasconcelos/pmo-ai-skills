---
name: termo-encerramento-projeto
description: Gera o Termo de Encerramento do Projeto em DOCX, com aprendizados, comparativo previsto x realizado de prazo e custo (considerando as mudanças aprovadas), status do escopo, motivo do encerramento e assinaturas. Use quando o usuário pedir para formalizar o encerramento de um projeto.
---

# Termo de Encerramento do Projeto

**Fase do ciclo de vida:** encerramento.

## Fluxo

1. Reúna a linha de base original no TAP: data de início, data final prevista e orçamento total.
2. Reúna as mudanças aprovadas (change requests), com o impacto de cada uma na data final e no custo. Elas atualizam a linha de base. Se nenhuma for informada, confirme com o usuário que não houve mudanças, porque o termo vai comparar o realizado com a linha de base original.
3. Reúna o realizado: data final real e custo real.
4. Pergunte sobre o escopo (concluído, homologado, pendências) e o motivo do encerramento. Se houver uma matriz de lições aprendidas, use as lições mais relevantes com o mesmo `id` (`LIC001`...).
5. Monte um JSON com a estrutura de `examples/entrada.json`. Omita o que não tiver sido informado.
6. Execute `python scripts/gerar_termo.py entrada.json saida.docx`.
7. Entregue o arquivo e resuma os desvios de prazo e custo, separando o que veio de mudanças aprovadas do que é desvio de execução.

## Comparativo

| Coluna | Cálculo |
|---|---|
| Previsto original | Linha de base do TAP |
| Mudanças aprovadas | Soma dos impactos dos change requests aprovados |
| Previsto atualizado | Original + mudanças aprovadas |
| Realizado | Valor real no encerramento |
| Desvio | Realizado - previsto atualizado, também em % |

O prazo é a duração em dias corridos, da data de início até a data final. Comparar com o previsto atualizado evita tratar como atraso ou estouro o que foi aprovado formalmente.

## Regras

- **Nunca invente informação.** O que não foi informado aparece como "a definir", e o desvio só é calculado quando houver dados suficientes.
- Se o escopo não foi concluído ou homologado, ou se há pendências, a justificativa é obrigatória.
- Em `motivo`, use `CONCLUSÃO`, `SUBSTITUIÇÃO` ou `CANCELAMENTO`. Substituição e cancelamento exigem `motivo_detalhe`.
- A declaração de que todas as entregas foram concluídas só aparece quando o motivo é conclusão e o escopo foi concluído.
- O termo deve ser assinado pelas mesmas pessoas que assinaram o TAP.
- Mantenha um tom direto e executivo.
