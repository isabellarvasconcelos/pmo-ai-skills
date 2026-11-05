---
name: change-request
description: Gera o Formulário de Requisição de Mudança (Change Request) em DOCX, com descrição da mudança, análise de impacto por aspecto, riscos, plano de execução e avaliação final. Use quando o usuário pedir para formalizar uma mudança de escopo, prazo, custo ou requisito em um projeto já em andamento.
---

# Change Request: Requisição de Mudança

**Fase do ciclo de vida:** monitoramento e controle, a qualquer momento da execução.

## Fluxo

1. Entenda a mudança: o que muda, por que muda e quem pediu.
2. Avalie o impacto em cada aspecto: cronograma, requisitos, recursos, custos, escopo, técnico e funcional. Pergunte o que não estiver claro, e não marque um aspecto como impactado sem justificativa.
3. Se houver TAP, matriz de riscos ou status report do projeto, use-os para comparar com a linha de base e apontar o que precisa ser atualizado.
4. Monte um JSON com a estrutura de `examples/entrada.json`. Omita o que não tiver sido informado.
5. Execute `python scripts/gerar_cr.py entrada.json saida.docx`.
6. Entregue o arquivo, resuma o impacto em prazo e custo e liste os documentos que precisam ser atualizados se a mudança for aprovada.

Em `resumo_impacto`, `prazo_dias` é o deslocamento da **data final do projeto**, e `custo` é a variação do orçamento. Atrasos em atividades que não mudam a data final ficam na justificativa do aspecto "Cronograma".

## Regras

- **Nunca invente informação.** O que não foi informado aparece como "a definir".
- Todo aspecto marcado como impactado precisa de justificativa. O script recusa um impacto sem ela.
- A decisão (`APROVADO` ou `REPROVADO`) só deve ser preenchida quando informada pelo usuário. Sem ela, o formulário sai pendente de avaliação.
- Uma mudança aprovada altera a linha de base: lembre o usuário de atualizar o TAP, o cronograma e a matriz de riscos.
- Mantenha um tom direto e executivo.
