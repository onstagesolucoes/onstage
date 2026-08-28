# Mapa do ClickUp — Fechamento Mensal

Referência dos IDs e campos usados pela skill `fechamento-fiscal-folha`.
Se a estrutura do ClickUp mudar, atualize este arquivo — a skill lê daqui.

## Localização

| Nível | Nome | ID |
|---|---|---|
| Workspace | Workspace | `90171048316` |
| Space | `4 \| Fiscal` | `90176845682` |
| Folder | `Fechamento Mensal` | `901710796494` |
| **Lista principal** | **`Fechamento Mensal`** | **`901715930997`** |
| Lista | `Revisão de base de clientes` | `901716413864` |
| Lista | `Projeto — Automação do Fechamento` | `901716042314` |
| Lista | `Fechamento Mensal - Teste` | `901716273989` |
| Lista | `Cadastro de Clientes` (space `6 I Operações`) | `901715756451` |

O campo `Cliente` da tarefa é uma relação (`tasks`) que aponta para a lista
`Cadastro de Clientes` (`901715756451`).

## Estrutura das tarefas

**Tarefa-mãe** — uma por cliente/competência, nome no padrão
`<código> - <Nome do Cliente>` (ex.: `77 - Benessere Serviços Médicos (Dra. Lara Corcetti)`).

**Subtarefas padronizadas**, na ordem do fluxo:

1. `Validação de Notas Fiscais`
2. `Validação de Fator R`
3. `PGDAS-D — Declaração`
4. `DAS — Guia de Pagamento`
5. `Pró-Labore`
6. `DARF Previdenciário`
7. `Envio das Obrigações — Nibo`
8. `Envio ao Cliente — WhatsApp`

Pode existir também uma tarefa de controle de idempotência no padrão
`<Cliente> — <Mês/Ano> (controle de idempotência)`.

Status de tarefa observados na lista: `não iniciado`, `complete`. Use
`clickup_get_task` com `expand_statuses: true` para obter a lista completa
antes de qualquer atualização de status.

## Campos personalizados da lista `901715930997`

### Controle de fase

| Campo | ID | Tipo | Valores |
|---|---|---|---|
| `Etapa Consolidada` | `1be24f56-19e7-42c1-9a12-47cea68b15db` | drop_down | Não iniciado · Aguardando OneFlow · Documentos disponíveis · Enviado ao Nibo · Protocolado · Enviado ao cliente · Concluído |
| `Etapa do Fechamento` | `39b11dc1-8e00-43c3-b87f-48c4bf52c45f` | drop_down | A iniciar · Em validação · Aguardando OneFlow · Pronto para envio Nibo · Enviado no Nibo · Aguardando WhatsApp · **Revisar** · Concluído |
| `Progresso do Fechamento` | `d85444df-7c48-4b5a-bd9e-9aa903fe9866` | automatic_progress | — |
| `Concluída` | `8610caf5-fa99-47e6-abe2-57b1a1d06cff` | checkbox | — |

### Integração OneFlow

| Campo | ID | Tipo | Valores |
|---|---|---|---|
| `Status OneFlow` | `badc90f7-8f37-4d42-bc6f-c27ac4e7c5b6` | drop_down | Não Iniciado · Concluído |
| `Finalizado OneFlow` | `4adc0092-d676-4856-898b-989172631917` | drop_down | Sim · Não |
| `Código OneFlow` | `e3d8f2c8-8288-44e2-a4c1-36c430dfceef` | short_text | — |

### Integração Nibo (envio das obrigações)

| Campo | ID | Tipo | Valores |
|---|---|---|---|
| `Status Nibo` | `14ac2127-3528-4250-97ea-7a1e661cd021` | drop_down | Enviado · Pendente Envio · Parcial · **Erro** |
| `Protocolo Nibo` | `76598145-14b3-48af-b4b8-7fbb7121de1d` | short_text | — |
| `Data Envio Nibo` | `6c0d95e0-9557-4aa5-be42-75ea8116bbbf` | date | — |

> Nunca preencher estes três campos. São afirmação de entrega feita.

### Competência e prazos

| Campo | ID | Tipo |
|---|---|---|
| `Competência — Mês/Ano` | `50e2583b-1605-4633-aef6-9d57acd44b7d` | text |
| `Competência` | `8ccc3d68-dd05-4aac-b4c2-bafc46ce9c07` | date |
| `Prazo Vencimento` | `48541efa-9bc0-4834-b788-f69bbc76d836` | date |
| `Data Meta Entrega` | `f7cafcb5-bb5b-4192-9a5f-82c05c5391ac` | date |
| `Enviado dentro da Meta` | `487f058f-e87a-460c-bff1-4721f6c882b8` | formula (só leitura) |

`Alerta Prazo` — `d6482b35-6fac-4ff6-b279-ad1c5e18616e` (drop_down):
Dentro da Meta · Dentro do Prazo · Vence em 3 dias · Vence Amanhã ·
Vence Hoje · Gerou Multas · Não Aplicável · Atenção · Atrasado · Sem Data ·
Concluído.

### Dados fiscais

| Campo | ID | Tipo | Valores |
|---|---|---|---|
| `11. Regime Tributario` | `6f2d5415-8bfb-4aaf-8ee1-8f25af053b62` | drop_down | MEI · Simples Nacional · Lucro Presumido · Lucro Real |
| `Percentual Fator R` | `26322627-d584-4ae2-83f7-4c7ea82b1b6f` | number | — |
| `Anexo da Obrigação` | `8dabcb4f-f294-46c7-99cf-6a394d6fc1c4` | attachment | — |

### Classificação e origem

| Campo | ID | Tipo | Valores |
|---|---|---|---|
| `Perfil de Fechamento` | `3b11c117-92d9-46be-85d9-108b779a5743` | drop_down | P1…P8 · P9 — Simples Nacional sem movimento · Aguardando classificação · Exceção manual |
| `Origem da Geração` | `bec52526-9254-4559-992c-88f49d9bce04` | drop_down | Produção · Teste · Migração · Reprocessamento · Manual |
| `Tipo` | `8dcfdaf1-0fba-4f36-a517-95c7ad273078` | drop_down | Automação · Manual |
| `Chave Técnica do Fechamento` | `99d4d02e-0563-4e7c-8e70-06aa338a97fa` | short_text | chave de idempotência |
| `Classificação CX` | `287395d0-04ca-4c46-b0ec-aee7c57a8787` | drop_down | Relacionamento padrão · Primeiro fechamento · Atenção especial CX · Recuperação de confiança · Cliente estratégico · Promotor / Embaixador |
| `Condição Comercial Especial` | `553b995a-cf48-4110-9fc8-8e915025b9c4` | drop_down | Sim · Não |

### Sócio responsável

| Campo | ID | Tipo |
|---|---|---|
| `Responsável pelo CNPJ` | `583907aa-aa68-4950-b3b0-0d62e8c08a06` | short_text |
| `CPF do Sócio Responsável` | `e6814cd0-3cef-4441-a75b-e15f9823992a` | short_text |
| `Qualificação do Sócio Responsável` | `db56d3d0-afbd-431c-b8d4-eb3b24ecd9c6` | drop_down (Sócio-Administrador · Sócio) |
| `Cliente` (relação) | `9fa20311-3155-42ce-92fb-00c68201d8ca` | tasks → `901715756451` |

## Como localizar a tarefa de um cliente/competência

```
clickup_filter_tasks(
  list_ids: ["901715930997"],
  include_closed: true,
  subtasks: true
)
```

Os resultados são paginados de 100 em 100 — quando `has_more` for `true`,
chame de novo com `page: next_page` até `has_more` ser `false`, senão a
busca fica incompleta e você pode concluir por engano que uma tarefa não
existe.

Para ler campos e anexos de uma tarefa:

```
clickup_get_task(
  task_id: "<id>",
  include: ["custom_fields", "subtasks", "attachments", "checklists"],
  expand_statuses: true
)
```
