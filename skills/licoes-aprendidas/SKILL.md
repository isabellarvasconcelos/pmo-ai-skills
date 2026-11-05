---
name: licoes-aprendidas
description: Gera e atualiza a Matriz de Lições Aprendidas de um projeto em XLSX, registrando pontos positivos e negativos, área relacionada, plano de ação, responsável e status. Use quando o usuário pedir para registrar aprendizados durante ou ao final de um projeto.
---

# Lições Aprendidas

**Fase do ciclo de vida:** monitoramento e controle (registro contínuo) e encerramento (fórum de lições aprendidas).

## Fluxo

1. Reúna as lições do material fornecido: relatos, atas, change requests, riscos que se concretizaram ou o fórum de encerramento.
2. Escreva cada lição como um fato e sua consequência ("A documentação da API estava desatualizada, o que atrasou a integração"), não como opinião ou culpa.
3. Para cada lição, proponha um plano de ação que mude o processo de projetos futuros, e confirme o responsável com o usuário.
4. Monte um JSON com a estrutura de `examples/entrada.json`. Omita o que não tiver sido informado.
5. Execute `python scripts/gerar_licoes.py entrada.json saida.xlsx`.
6. Entregue o arquivo, resuma quantas lições são positivas e negativas e liste as ações ainda abertas.

Em uma atualização, mantenha o `id` das lições existentes. Lições novas recebem o próximo número livre.

O termo de encerramento reaproveita estas lições: use o mesmo `id` (`LIC001`, `LIC002`...) nos dois documentos.

## Regras

- **Nunca invente lições, responsáveis ou datas.** O que não foi informado aparece como "a definir".
- Registre também os pontos positivos. Eles mostram o que deve ser repetido.
- Use em `tipo`: `POSITIVO` ou `NEGATIVO`.
- Use em `status`: `ABERTO`, `EM ANDAMENTO` ou `CONCLUÍDO`.
- Use em `momento`: `DURANTE O PROJETO`, `NO ENCERRAMENTO` ou `PRÓXIMOS PROJETOS`, conforme quando o plano de ação será aplicado.
- A data de encerramento só deve ser preenchida em ações concluídas.
- Mantenha um tom direto e executivo.
