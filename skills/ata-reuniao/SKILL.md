---
name: ata-reuniao
description: Gera a Ata de Reunião em DOCX, com identificação, objetivo, resumo executivo, iniciativas mapeadas, decisões, itens de ação e próximos passos. Use quando o usuário pedir para gerar, redigir ou formalizar a ata de uma reunião, alinhamento, comitê ou acompanhamento de projeto, inclusive a partir de transcrições ou anotações.
---

# Ata de Reunião

**Fase do ciclo de vida:** transversal, em qualquer fase em que uma reunião precise ser formalizada.

## Fluxo

1. Leia o material fornecido: transcrição, anotações ou relato da reunião.
2. Separe o conteúdo nas seções da ata. Distinga o que foi **decidido** do que foi apenas **discutido**: só entra em "Decisões" o que foi fechado na reunião.
3. Para cada item de ação, identifique o responsável e o prazo. Se o material não disser, use "a definir" e aponte a lacuna na resposta.
4. Monte um JSON com a estrutura de `examples/entrada.json`. Omita o que não tiver sido informado.
5. Execute `python scripts/gerar_ata.py entrada.json saida.docx`.
6. Entregue o arquivo e liste as lacunas: itens sem responsável ou prazo e pontos que pareceram decididos mas ficaram ambíguos.

Veja em `examples/notas-reuniao.md` um exemplo de anotações soltas que viraram a ata de `examples/saida.docx`.

## Regras

- **Nunca invente participantes, decisões ou itens de ação** que não estejam no material.
- Registre as decisões como frases afirmativas e verificáveis ("Aprovada a CR-001"), não como temas ("Discussão sobre o Wi-Fi").
- Use em `status` dos itens de ação: `Pendente`, `Em andamento` ou `Concluído`.
- A seção de iniciativas mapeadas só aparece quando houver iniciativas. Ela serve para registrar demandas novas que surgirem na reunião, para priorização posterior.
- Mantenha um tom direto e executivo.
