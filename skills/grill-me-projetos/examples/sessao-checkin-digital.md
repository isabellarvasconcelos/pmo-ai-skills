# Exemplo: descoberta do check-in digital de entregas

Sessão fictícia que mostra como a skill conduz a entrevista. Empresa, pessoas e números são inventados. Este é o mesmo projeto usado nos exemplos de TAP, matriz de riscos e status report deste repositório.

## Pedido inicial

> As filas na portaria estão enormes. Quero digitalizar o check-in dos caminhões.

## Rodada 1

A fronteira inicial tem quatro decisões que não dependem umas das outras.

❓ **P1. Problema principal**: a dor é o tempo de espera na portaria, a falta de visibilidade das entregas ou o retrabalho de digitar as fichas? As três podem existir, mas cada uma leva a um resultado diferente.

➡️ Recomendo tratar o tempo de espera como problema principal, porque ele é visível para quem chega, é mensurável e é o que motivou o pedido.

❓ **P2. Abrangência**: o check-in digital vale para todas as unidades ou só para uma?

➡️ Recomendo começar por uma unidade. Isso permite validar o processo antes de replicar.

❓ **P3. Agendamento**: o agendamento prévio de entregas faz parte desta demanda?

➡️ Recomendo deixar de fora. Agendamento muda o processo das transportadoras e do comercial, o que forma outra frente de trabalho, com resultado próprio.

❓ **P4. Resultado esperado**: quanto tempo leva hoje o atendimento na portaria, e qual seria um resultado bom?

➡️ Não encontrei esse dado nos materiais disponíveis. Se não houver medição, recomendo medir durante uma semana antes de fixar a meta.

**Respostas:**
1. Tempo de espera.
2. Só a unidade principal.
3. Fora de escopo.
4. A portaria registrou média de 20 minutos no último mês. A meta é chegar a 8.

## Rodada 2

As respostas abriram novas decisões.

❓ **P5. Exceções**: o que acontece quando o caminhão chega sem nota fiscal válida?

➡️ Recomendo não concluir o check-in e encaminhar o motorista ao atendimento da portaria, que já trata esse caso hoje.

❓ **P6. Visibilidade da fila**: a equipe da portaria precisa ver a fila em tempo real, ou basta o registro?

➡️ Recomendo incluir a visualização da fila, porque é ela que permite à portaria reagir antes que a espera cresça.

❓ **P7. Pronto para usar**: quando a equipe da portaria estará apta a operar o novo processo?

➡️ Recomendo exigir treinamento com lista de presença como condição de aceite.

**Respostas:** 5, 6 e 7 aceitos como recomendados.

**Vazamento de escopo identificado:** surgiu a pergunta sobre como o almoxarifado deveria reorganizar o recebimento com base na fila. Isso decide como outra área reage ao resultado, então foi registrado em "Fora de escopo", com o almoxarifado como responsável.

## Histórias consolidadas

1. Como motorista, preciso fazer o check-in sem preencher ficha em papel, para ser liberado mais rápido.
2. Como equipe da portaria, preciso ver a fila de veículos em tempo real, para priorizar o atendimento.
3. Como equipe da portaria, preciso que o check-in seja bloqueado quando a nota fiscal não for válida, para não liberar entregas irregulares.
4. Para medir o resultado do projeto, é necessário registrar o horário de chegada e de liberação de cada veículo, para que o tempo médio de atendimento possa ser acompanhado.

**Confirmação:** a pessoa confirmou as quatro histórias, e o PRD foi gerado.

## PRD resultante (resumo)

# Check-in digital de entregas

## Declaração do problema

Na unidade principal, o check-in de caminhões é feito em ficha de papel e digitado depois pela portaria. O atendimento leva em média 20 minutos, segundo o registro da portaria no último mês, o que forma filas e atrasa o recebimento.

## Objetivo e resultado esperado

Reduzir o tempo médio de atendimento na portaria de 20 para 8 minutos, mantido por 30 dias consecutivos.

## Escopo

### Incluído

- Check-in digital na portaria da unidade principal.
- Visualização da fila em tempo real pela portaria.
- Treinamento da equipe da portaria.

### Fora de escopo

- **Agendamento de entregas:** muda o processo das transportadoras e do comercial e deve ser tratado como iniciativa própria.
- **Implantação nas demais unidades:** será avaliada depois da validação nesta unidade.
- **Reorganização do recebimento no almoxarifado:** decisão do almoxarifado, a partir dos dados de fila.

## Regras e critérios de aceite

| ID | Condição | Resultado esperado | Exceções ou observações |
|---|---|---|---|
| RN01 | Dado que o veículo não tem nota fiscal válida, quando o motorista tenta concluir o check-in | O check-in não é concluído e o motorista é encaminhado ao atendimento | Nenhuma |
| CA01 | Dado o período de 30 dias após a implantação | O tempo médio de atendimento é de até 8 minutos | Medido pelos horários de chegada e liberação |

## Pontos para validação posterior

- **Hipótese:** a cobertura de Wi-Fi na portaria é suficiente para o check-in digital.
