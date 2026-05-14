STEP Dashboard Tracking - Portugal v38.6

Base utilizada:
- step_dashboard_tracking_CLIENTE_CORRIGIDO_v38.5_CORRIGE_BOTOES_BR_PT

Alteração aplicada:
- API interna do Portal do Cliente na versão Portugal.
- Botão "Gerar API" no topo, após "Instalar app".
- Botão "Gerar API" também no painel do cliente.
- Token completo com botão de copiar no momento da criação.
- Lista de chaves com preview mascarado.
- Revogar chave ativa.
- Excluir chave revogada.
- Endpoint de dados somente leitura.

Endpoints criados:
- GET /api/client-api-keys
- POST /api/client-api-keys
- DELETE /api/client-api-keys
- GET /api/client-data

Segurança:
- /api/client-data aceita somente GET.
- Chaves criadas possuem somente escopo read:projects.
- Token completo só aparece no momento da criação.
- Supabase armazena apenas hash do token.
- Chave ativa não pode ser excluída diretamente; precisa revogar antes.

SQL necessário:
- Executar supabase/SQL-CRIAR-CLIENT-API-KEYS.sql no Supabase.

Observação:
- Mantidas as regras e telas da v38.5 Portugal, incluindo login BR/PT e correções dos botões.
