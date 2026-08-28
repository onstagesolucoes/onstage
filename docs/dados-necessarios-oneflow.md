# Dados necessários do OneFlow para a revisão de fechamento

Especificação do que a skill `fechamento-fiscal-folha` precisa receber para
conferir **valores** de apuração, e não apenas o estado do processo no
ClickUp.

## Por que este documento existe

A skill hoje consegue validar o *processo*: cruza o que a tarefa do ClickUp
afirma com o que existe de lastro (anexos, protocolos, campos preenchidos) e
detecta erros da automação `OneFlow → ClickUp → Nibo`. Um teste real
(Simples Nacional, Jul/2026) confirmou que isso funciona — encontrou
duplicidade de chave técnica, subtarefas concluídas sob tarefa-mãe não
iniciada e obrigação vencida com alerta de prazo ainda verde.

O que ela **não** consegue fazer é conferir número: faturamento, RBT12,
alíquota efetiva, valor do DAS, Fator R. Nada disso trafega pelo ClickUp.
As seções 1 a 3 da skill (Simples Nacional, Lucro Presumido e Folha) só
passam a funcionar quando esses dados chegarem até ela.

Este documento lista o que pedir. Ele descreve **dados**, não rotas: a
documentação da API do OneFlow (SwaggerHub, `oneflowoficial/integracoes`)
não estava acessível a partir do ambiente onde a skill foi construída, então
nenhum endpoint é nomeado aqui de propósito. Use a lista como checklist ao
ler o Swagger: para cada bloco, identifique a rota que entrega aquilo.

## Prioridade 1 — sem isso não se confere número nenhum

### A. Apuração / PGDAS da competência

competência · receita bruta declarada · RBT12 · receita acumulada no ano ·
segregação por atividade/anexo · anexo aplicado · faixa · alíquota nominal ·
parcela a deduzir · **alíquota efetiva** · valor desmembrado por tributo
(IRPJ, CSLL, COFINS, PIS, CPP, ICMS, ISS) · retenções deduzidas · valor
total · nº do recibo/declaração · flag de retificadora

Alimenta a seção 1.2 da skill. É o núcleo: sem isso a conciliação
`FATURAMENTO → PGDAS → DAS` não existe.

### B. Guias emitidas (DAS / DARF)

tipo · competência · período de apuração · valor principal · multa · juros ·
**total** · vencimento · linha digitável · **status de pagamento** · data de
pagamento

Alimenta as seções 1.3 e 1.6. O status de pagamento é o que fecha o buraco
encontrado no teste: hoje dá para ver que o prazo passou, mas não se a guia
foi paga.

### C. Notas fiscais da competência

número · série · data de emissão · tomador · valor · discriminação/serviço ·
**status (ativa / cancelada / substituída)** · referência da nota
substituída · retenções (ISS, IRRF, PIS/COFINS/CSLL, INSS)

Alimenta a seção 1.1 — conferência de faturamento e detecção de duplicidade.

> Os scripts atuais do repositório (`ler_xml*.py`) leem XML de **NFe**.
> Clientes de serviços (médicos, por exemplo) emitem **NFS-e municipal**,
> que é outro documento com outro layout. Se a API do OneFlow já consolidar
> os dois, resolve um problema que hoje é nosso.

## Prioridade 2 — sem isso a skill confere, mas não questiona

### D. Histórico de faturamento (13+ meses)

faturamento mensal por competência

Alimenta os limiares de alerta de 30% / 50%, as médias de 3 e 12 meses e a
evolução do RBT12. É o que separa "o número está certo" de "o número faz
sentido".

### E. Fator R

numerador (folha 12 meses, incluindo pró-labore e encargos) · denominador
(RBT12) · **percentual** · anexo resultante

Corresponde ao campo `Percentual Fator R` do ClickUp, hoje vazio. Para
carteira concentrada em serviços, decide Anexo III × V — é o cálculo de
maior impacto financeiro da operação.

### F. Folha e pró-labore da competência

pró-labore por sócio · folha total · INSS segurado e patronal · IRRF ·
FGTS · base previdenciária · headcount

Alimenta as seções 1.5 e 3.1.

## Prioridade 3 — completam a revisão

### G. Eventos de folha do mês
admissões · rescisões · férias · 13º salário — seção 3.2.

### H. Status das obrigações acessórias
eSocial (eventos periódicos e status de transmissão) · DCTFWeb (valor,
recibo, status) · FGTS Digital — seção 3.3.

### I. Rastreabilidade da apuração
identificador da apuração no OneFlow · status · data/hora de conclusão ·
usuário responsável.

Resolve dois problemas: preenche o campo `Código OneFlow` (hoje vazio) e
permite amarrar tarefa do ClickUp ↔ apuração de origem.

## Duas perguntas a fazer à API além das rotas

**Existe webhook de conclusão de apuração?** O disparo desejado é "quando a
apuração conclui no OneFlow". Com webhook, o agente reage no momento certo.
Sem ele, vira polling — e a pergunta passa a ser qual rota lista apurações
concluídas por período, para não varrer cliente a cliente.

**O que a API devolve que o PDF também tem?** Onde coincidirem, usar a API:
extrair número de PDF é frágil e caro. O PDF continua valendo como
comprovante para anexar no campo `Anexo da Obrigação` do ClickUp, hoje vazio
em todas as tarefas examinadas.

## Como usar isto

Ao obter acesso ao Swagger do OneFlow, mapeie rota a rota contra os blocos
A–I e registre aqui o que cada rota cobre. Onde não houver rota, a skill
permanece cega naquele ponto — e isso deve ficar explícito no painel de
revisão, para ninguém ler um veredito 🔴 e supor que a apuração foi conferida.
