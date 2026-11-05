---
name: tap-termo-abertura-projeto
description: Gera o Termo de Abertura do Projeto (TAP) em XLSX, com justificativa, objetivo SMART, AS IS x TO BE, escopo, requisitos, solução técnica, impactos, recursos, orçamento e aquisições. Use quando o usuário pedir para abrir formalmente um projeto ou formalizar seu escopo inicial.
---

# TAP: Termo de Abertura do Projeto

**Fase do ciclo de vida:** iniciação.

## Fluxo

1. Reúna os dados que o usuário já forneceu, sem repetir perguntas sobre o que já está claro.
2. Pergunte apenas o que falta para entender justificativa, objetivo, escopo e orçamento. Se o usuário não souber, siga em frente.
3. Monte um JSON com a estrutura de `examples/entrada.json`. Omita o que não tiver sido informado.
4. Execute `python scripts/gerar_tap.py entrada.json saida.xlsx`.
5. Entregue o arquivo e liste os campos que ficaram como "a definir" e os pontos do escopo que merecem revisão.

## Regras

- **Nunca invente informação.** Campos sem dado ficam como "a definir", o que o script faz automaticamente.
- Escreva o objetivo no formato SMART: verbo no infinitivo, resultado mensurável e prazo. Se faltar a métrica ou o prazo, aponte a lacuna em vez de criá-la.
- Declare exclusões em "não contempla", porque elas evitam disputas de escopo mais tarde.
- Trate cada premissa como um possível risco e sugira levá-la à matriz de riscos.
- Não edite totais à mão. O orçamento usa fórmulas, e o total da aba "Aquisições" alimenta a aba "TAP".
- Use `tipo` apenas com `CAPEX` ou `OPEX`. O script recusa outros valores e valores numéricos inválidos.
- Mantenha um tom direto e executivo.
