# Conexa → NIBO (tela de conferência) → baixa no ClickUp

Automação em dois workflows do n8n:

1. **`conexa-nibo-envio-fatura.json`** — disparado por webhook do Conexa quando uma
   fatura de serviço é gerada. Busca o cliente no Supabase, baixa o PDF da fatura no
   Conexa, envia para a tela de conferência do NIBO Obrigações e cria a tarefa do dia
   no ClickUp.
2. **`nibo-clickup-confirmacao.json`** — roda 1x por dia. Verifica no NIBO quais
   envios já foram confirmados manualmente pelo contador ("confirmar entrega manual
   do protocolo") e marca a tarefa correspondente como concluída no ClickUp (a
   "baixa").

Os dois reaproveitam o mesmo projeto/tabela Supabase (`clientes`) e o mesmo padrão de
autenticação do NIBO já validados em produção em outra automação (OneFlow → NIBO).

## Antes de importar

1. Rode a migração `supabase/migrations/001_conexa_nibo_clickup.sql` no seu projeto
   Supabase (adiciona `conexa_customer_id` e `clickup_list_id` em `clientes`, e cria a
   tabela `envios_nibo_faturas`).
2. Popule, para cada cliente que vai usar essa automação:
   - `clientes.conexa_customer_id` — o `customer.id` retornado no payload do webhook
     do Conexa
   - `clientes.clickup_list_id` — o ID da lista do ClickUp onde a tarefa diária desse
     cliente deve ser criada
   - confirme que `clientes.cnpj` e `clientes.ativo` já estão preenchidos (usados pelo
     workflow 1)

## Credenciais / variáveis de ambiente do n8n

Nenhum segredo fica hardcoded nos arquivos JSON — os dois workflows usam expressões
`{{ $env.NOME_DA_VARIAVEL }}`. Configure no ambiente do n8n (Settings → Environment
Variables, ou variáveis de ambiente do processo/container):

| Variável | Descrição |
|---|---|
| `SUPABASE_URL` | URL REST do projeto Supabase, ex. `https://xxxx.supabase.co` |
| `SUPABASE_SERVICE_KEY` | Service role key do Supabase (mesma já usada no fluxo OneFlow → NIBO) |
| `NIBO_API_KEY` | `X-API-Key` do NIBO (mesma já usada no fluxo OneFlow → NIBO) |
| `NIBO_USER_ID` | `X-User-Id` usado no passo de conferência do NIBO |
| `NIBO_ACCOUNTING_FIRM_ID` | ID do escritório contábil no NIBO (`accountingfirms/{id}`) |
| `CONEXA_API_BASE_URL` | Base da API do Conexa (⚠️ confirmar na doc do Postman) |
| `CONEXA_API_TOKEN` | Token de autenticação da API do Conexa (⚠️ confirmar tipo de auth na doc) |
| `CLICKUP_API_TOKEN` | Personal token da API do ClickUp |
| `CLICKUP_STATUS_CONCLUIDO` | (opcional) nome exato do status "concluído" configurado na sua lista do ClickUp — default `complete` se não definido |

## ⚠️ Pontos marcados TODO-DOC — validar antes de ativar

Os domínios `nibo.readme.io` e `documenter.getpostman.com` estavam bloqueados na
sessão em que esta automação foi montada, então os pontos abaixo foram deixados
explicitamente marcados nos nós do n8n (comentário/nota no node) para você confirmar
contra a documentação real antes de ativar em produção:

1. **Path/evento exato do webhook do Conexa** para "fatura de serviço gerada" (nó
   `Webhook Conexa - Fatura Gerada` do workflow 1) — cadastre esse path no painel do
   Conexa e ajuste o `path` do nó se necessário.
2. **Campo que identifica a fatura no payload do webhook** (nó
   `Extrair Dados do Webhook`) — está usando `chargeId` como palpite (visto em outro
   evento do Conexa), mas pode ter um nome diferente no evento de fatura gerada.
3. **Endpoint de download do PDF da fatura** no Conexa (nó
   `Conexa - Baixar PDF da Fatura`) — o usuário confirmou que existe uma rota própria
   para isso; o path usado é um placeholder.
4. **Endpoint "confirmar entrega manual do protocolo"** do NIBO (nó
   `NIBO - Confirmar Entrega Manual do Protocolo`, workflow 2) — método, path e campo
   de resposta que indica confirmação ainda não foram validados.
5. **Nome do status "concluído"** na sua lista do ClickUp (nó
   `ClickUp - Concluir Tarefa`, workflow 2) — ClickUp exige o nome exato do status
   configurado na lista; ajuste `CLICKUP_STATUS_CONCLUIDO` se não for `complete`.
6. O nó **`Envio OK?`** do workflow 1 verifica sucesso checando ausência de um campo
   `error` nos nós anteriores (comportamento do n8n com `onError: continueRegularOutput`)
   — vale rodar uma execução de teste no n8n para confirmar que esse formato bate com
   a versão do seu n8n antes de confiar nele em produção.

O que **já está confirmado** (validado em produção em outra automação real, reaproveitado
aqui sem alteração de fundo): os 3 passos de envio ao NIBO (criar arquivo → upload no
Azure Blob → criar conferência) e o padrão de acesso ao Supabase via REST
(`apikey`/`Authorization` com a service key, upsert com `Prefer: resolution=merge-duplicates`).

## Importando e testando

1. No n8n: **Workflows → Import from File** e selecione cada um dos dois JSONs.
2. Configure as variáveis de ambiente listadas acima.
3. Ajuste os pontos `⚠️ TODO-DOC` conforme a documentação real do Conexa/NIBO/ClickUp.
4. Rode manualmente o workflow 1 (**Execute Workflow** com um payload de teste, ou
   dispare o webhook real com um cliente de teste) e confira no Supabase se a linha
   foi gravada em `envios_nibo_faturas` com `status='enviado'`.
5. Rode manualmente o workflow 2 e confira se ele encontra o registro pendente, checa
   a confirmação no NIBO e (quando confirmado) marca a tarefa no ClickUp e atualiza o
   status para `'confirmado'` no Supabase.
6. Só então ative os dois workflows (o webhook do workflow 1 e o agendamento diário do
   workflow 2).

## Erros

Não há canal de alerta (e-mail/Slack) configurado — falhas no envio ficam registradas
em `envios_nibo_faturas.status='erro'` / `error_message`, consultáveis diretamente no
Supabase. Pode ser adicionado depois se for necessário.
