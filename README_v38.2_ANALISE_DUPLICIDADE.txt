v38.2 - Análise e correção da validação de duplicidade

Achado:
- O bloqueio acontece no endpoint /api/admin-users, antes de salvar.
- A função antiga retornava só booleano e não mostrava qual usuário causava o conflito.
- Agora a validação separa:
  Admin/PCP = universal
  demais usuários/clientes = username + BR/PT

Também foi adicionado retorno detalhado no erro:
- login conflitante
- nome
- perfil/setor
- ambiente existente
- ambiente solicitado
- clientKey existente

Incluído:
DIAGNOSTICO_DUPLICIDADE_USUARIOS.sql
