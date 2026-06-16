(function () {
  'use strict';

  const STORAGE_KEY = 'step-dashboard-language';
  const DEFAULT_LANG = 'pt-BR';
  const SUPPORTED_LANGS = ['pt-BR', 'en-US', 'es-ES'];
  const LANGUAGE_LABELS = {
    'pt-BR': 'Português',
    'en-US': 'English',
    'es-ES': 'Español',
  };

  const DICT = {
    'en-US': {
      'Idioma': 'Language',
      'Selecionar idioma do painel': 'Select dashboard language',
      'Português': 'Portuguese',
      'Inglês': 'English',
      'Espanhol': 'Spanish',
      'PAINEL OPERACIONAL DE PROJETOS': 'PROJECTS OPERATIONAL DASHBOARD',
      'STEP • Painel Operacional': 'STEP • Operational Dashboard',
      'Preparando seu painel operacional': 'Preparing your operational dashboard',
      'Carregando informações...': 'Loading information...',
      'Estamos validando seu acesso.': 'We are validating your access.',
      'Aguarde, conectando aos dados operacionais.': 'Please wait, connecting to operational data.',
      'Acesso': 'Access',
      'Acesso bloqueado': 'Access locked',
      'Faça login para liberar o painel, os alertas por setor e a administração.': 'Sign in to unlock the dashboard, sector alerts and administration.',
      'offline': 'offline',
      'Login': 'Login',
      'Meus alertas': 'My alerts',
      'Minhas sinalizações': 'My flags',
      'Alertas Projetos': 'Project alerts',
      'Apontamentos': 'Stage updates',
      'Admin': 'Admin',
      'Sair': 'Sign out',
      'Instalar app': 'Install app',
      'Gerar API': 'Generate API',
      'Alterar senha': 'Change password',
      'Brasil': 'Brazil',
      'Portugal': 'Portugal',
      '1. Total de projetos': '1. Total projects',
      '2. Projetos iniciados': '2. Started projects',
      '3. Projetos não iniciados': '3. Not started projects',
      '4. Projetos em On Hold': '4. Projects on hold',
      '5. Projetos em produção': '5. Projects in production',
      '6. Projetos em qualidade': '6. Projects in quality',
      '7. Projetos em pintura': '7. Projects in painting',
      '8. Projetos preparados para envio': '8. Projects prepared for shipment',
      'Clique para detalhar': 'Click for details',
      'Tags --': 'Tags --',
      'Apenas aguardando envio': 'Awaiting shipment only',
      'Peso total programado': 'Total planned weight',
      'Peso total soldado': 'Total welded weight',
      'Total pendente de solda': 'Total pending welding',
      '(M²) Programada': 'Planned area (m²)',
      'Total enviado -- kg': 'Total shipped -- kg',
      '-- m²': '-- m²',
      'Projetos por cliente': 'Projects by client',
      'Visão por cliente': 'Client view',
      'Voltar': 'Back',
      'Projetos': 'Projects',
      'Alertas de prazo': 'Deadline alerts',
      'Base operacional': 'Operational base',
      'Aguardando atualização...': 'Waiting for update...',
      'Atualizar agora': 'Update now',
      'Buscar projeto, cliente, item ou referência': 'Search project, client, item or reference',
      'Filtrar demanda': 'Filter demand',
      'Todas as demandas': 'All demands',
      'Filtrar tipo': 'Filter type',
      'Todos os tipos': 'All types',
      'Filtrar semana': 'Filter week',
      'Todas as semanas': 'All weeks',
      'Filtrar status': 'Filter status',
      'Todos os status': 'All statuses',
      'Limpar': 'Clear',
      'Baixar Excel': 'Download Excel',
      '-- resultado(s)': '-- result(s)',
      'Projeto': 'Project',
      'Tipo': 'Type',
      'Cliente': 'Client',
      'Unidade': 'Unit',
      'Término planejado': 'Planned finish',
      'Itens': 'Items',
      'Peso soldado (kg)': 'Welded weight (kg)',
      'Semana finalizado': 'Finished week',
      'Peso (kg)': 'Weight (kg)',
      'Área operacional (m²)': 'Operational area (m²)',
      'Etapa atual': 'Current stage',
      '% Individual': '% Individual',
      '% Geral': '% Overall',
      'Status': 'Status',
      'Fab. início': 'Fab. start',
      'Caldeiraria': 'Boilermaking',
      'Calderaria': 'Boilermaking',
      'Solda': 'Welding',
      'Inspeção (QC)': 'Inspection (QC)',
      'TH': 'Hydro Test',
      'Data de envio': 'Shipment date',
      'Concluído?': 'Completed?',
      'Carregando projetos...': 'Loading projects...',
      'Detalhamento operacional': 'Operational details',
      'Resumo rápido do projeto selecionado. Dê 2 cliques na linha para abrir o detalhamento completo.': 'Quick summary of the selected project. Double-click the row to open the full details.',
      'Fechar': 'Close',
      'Online': 'Online',
      'Offline': 'Offline',
      'Sistema online • sincronização operacional': 'System online • operational sync',
      'Versão dos dados: --': 'Data version: --',
      'Itens internos e andamento por etapa': 'Internal items and stage progress',
      'Atenção': 'Attention',
      'Prazos em alerta': 'Deadlines in alert',
      'O alerta reaparece a cada 4 horas até que a planilha seja atualizada.': 'The alert reappears every 4 hours until the sheet is updated.',
      'Acesso setorial': 'Sector access',
      'Entre com seu usuário e senha': 'Enter your username and password',
      'Usuário': 'Username',
      'Senha': 'Password',
      'Mostrar': 'Show',
      'Entrar': 'Sign in',
      'Acesse com seu usuário setorial ou admin.': 'Access with your sector user or admin account.',
      'Alerta Operacional': 'Operational Alert',
      'Meus alertas por setor': 'My sector alerts',
      'Os alertas abaixo respeitam o setor do login atual.': 'The alerts below follow the sector of the current login.',
      'Administração': 'Administration',
      'Crie logins por setor, edite usuários e sincronize os dados com o repositório.': 'Create sector logins, edit users and sync data with the repository.',
      'Usuários e alertas setoriais': 'Users and sector alerts',
      'Novo usuário': 'New user',
      'Histórico operacional': 'Operational history',
      'Novo alerta operacional': 'New operational alert',
      'Registro de Alertas Operacionais': 'Operational Alert Log',
      'Pesquisar no histórico de alertas': 'Search alert history',
      'Subir pro GitHub': 'Push to GitHub',
      'GitHub sync: offline': 'GitHub sync: offline',
      'Perfil': 'Role',
      'Setor': 'Sector',
      'Nome': 'Name',
      'Nome exibido no Portal': 'Name shown in Portal',
      'Cliente vinculado': 'Linked client',
      'Logo do cliente': 'Client logo',
      'Importar logo do computador': 'Import logo from computer',
      'Zoom -': 'Zoom -',
      'Zoom +': 'Zoom +',
      'Resetar': 'Reset',
      'Aplicar ajuste': 'Apply adjustment',
      'Fotos por plataforma / vessel': 'Photos by platform / vessel',
      'Adicionar foto dessa plataforma': 'Add photo for this platform',
      'PMs adicionais para visualizar': 'Additional PMs to view',
      'Carregando nomes de PM do Tracking...': 'Loading PM names from Tracking...',
      'Nenhum PM adicional selecionado.': 'No additional PM selected.',
      'Competências da Qualidade': 'Quality competencies',
      'Engenharia': 'Engineering',
      'Suprimento': 'Supply',
      'Pintura': 'Painting',
      'Qualidade': 'Quality',
      'Logística': 'Logistics',
      'Produção': 'Production',
      'PCP': 'PCP',
      'Final Inspection QC': 'Final Inspection QC',
      'Inspeção Dimensional Inicial': 'Initial Dimensional Inspection',
      'Inspeção Dimensional Final': 'Final Dimensional Inspection',
      'END / NDE': 'NDE',
      'Visualizar Painel do Cliente ao abrir uma BSP': 'View Client Dashboard when opening a BSP',
      'Criar usuário': 'Create user',
      'Cancelar edição': 'Cancel editing',
      'Usuários cadastrados': 'Registered users',
      'Nenhum usuário carregado.': 'No user loaded.',
      'Monitor de acessos ao vivo': 'Live access monitor',
      'Setores monitorados': 'Monitored sectors',
      'Usuários online e última tela visualizada no sistema.': 'Online users and last screen viewed in the system.',
      '0 online': '0 online',
      'Nenhum usuário online agora.': 'No user online now.',
      'Respostas dos alertas': 'Alert responses',
      'Nenhuma resposta recebida ainda.': 'No response received yet.',
      'Título': 'Title',
      'Mensagem': 'Message',
      'Prioridade': 'Priority',
      'Normal': 'Normal',
      'Alta': 'High',
      'Urgente': 'Urgent',
      'Baixa': 'Low',
      'Exigir confirmação de leitura': 'Require read confirmation',
      'Enviar alerta': 'Send alert',
      'Resposta do alerta': 'Alert response',
      'Responder alerta': 'Reply to alert',
      'Escreva sua devolutiva para o administrador.': 'Write your feedback to the administrator.',
      'Sua resposta': 'Your response',
      'Enviar resposta': 'Send response',
      'Cancelar': 'Cancel',
      'Sinalização ao PCP': 'Flag to PCP',
      'Nova sinalização': 'New flag',
      'A informação será enviada ao PCP.': 'The information will be sent to PCP.',
      'Descrição': 'Description',
      'Enviar ao PCP': 'Send to PCP',
      'Novo alerta recebido': 'New alert received',
      'Novo alerta': 'New alert',
      'Atenção imediata necessária.': 'Immediate attention required.',
      'Você recebeu uma nova notificação.': 'You received a new notification.',
      'Abrir alerta': 'Open alert',
      'Segurança de acesso': 'Access security',
      'Digite a senha atual e defina a nova senha de acesso.': 'Enter the current password and define the new access password.',
      'Senha atual': 'Current password',
      'Nova senha': 'New password',
      'Confirmar nova senha': 'Confirm new password',
      'Salvar nova senha': 'Save new password',
      'Apontamentos por etapa': 'Stage updates',
      'Apontamentos setoriais': 'Sector stage updates',
      'Cada setor informa somente o avanço da sua própria etapa e o PCP valida.': 'Each sector reports only its own stage progress and PCP validates it.',
      'Validação PCP / Apontamentos': 'PCP validation / Stage updates',
      'Validação PCP': 'PCP validation',
      'Todos': 'All',
      'Nenhum projeto encontrado para a busca informada.': 'No project found for the search provided.',
      'Nenhuma tag detalhada encontrada.': 'No detailed tag found.',
      'Nenhuma BSP encontrada para este cliente.': 'No BSP found for this client.',
      'Nenhuma BSP nesta unidade.': 'No BSP in this unit.',
      'Atualizando...': 'Updating...',
      'Atualizando informações...': 'Updating information...',
      'Falha ao atualizar': 'Failed to update',
      'tente novamente': 'try again',
      'Atualização: --': 'Update: --',
      'BSPs': 'BSPs',
      'abrir visão executiva': 'open executive view',
      'Cronograma não disponível para esta ISO.': 'Schedule not available for this ISO.',
      'Schedule não disponível para esta BSP.': 'Schedule not available for this BSP.',
      'Etapa': 'Stage',
      'Etapas': 'Stages',
      'Etapa atual': 'Current stage',
      'Status atual': 'Current status',
      'Verificando estoque': 'Checking stock',
      'Separação de material': 'Material separation',
      'Corte e Limpeza': 'Cutting and Cleaning',
      'Corte e limpeza': 'Cutting and Cleaning',
      'Pré - Montagem': 'Pre-assembly',
      'Pré-Montagem': 'Pre-assembly',
      'Pré Montagem': 'Pre-assembly',
      'Pre Montagem': 'Pre-assembly',
      'Pré - Full Welding Dimensional/3D Inspection': 'Pre-weld dimensional/3D inspection',
      'Spool Assemble and Tack Weld': 'Spool assembly and tack weld',
      'Spool Assemble and tack weld': 'Spool assembly and tack weld',
      'Full Welding Execution': 'Full welding execution',
      'Full welding execution': 'Full welding execution',
      'Aguardando END': 'Awaiting NDE',
      'Aguardando NDE': 'Awaiting NDE',
      'END': 'NDE',
      'Inspeção': 'Inspection',
      'Inspeção Dimensional': 'Dimensional Inspection',
      'Inspeção Dimensional Final - 3D': 'Final Dimensional Inspection - 3D',
      'Inspeção Dimensional Inicial - 3D': 'Initial Dimensional Inspection - 3D',
      'Final Dimensional Inpection/3D': 'Final Dimensional Inspection/3D',
      'Final Dimensiona/3D Inspection': 'Final Dimensional/3D Inspection',
      'TH / Teste Hidrostático': 'Hydro Test',
      'Teste Hidrostático': 'Hydro Test',
      'Unitização e Inspeção': 'Unitization and Inspection',
      'Unitização e envio': 'Unitization and Shipment',
      'Unitização': 'Unitization',
      'Aguardando início de pintura': 'Awaiting painting start',
      'Início de pintura': 'Painting start',
      'Aguardando início': 'Awaiting start',
      'Aguardando liberação para envio': 'Awaiting shipment release',
      'Verificação final': 'Final verification',
      'Conferência final': 'Final check',
      'Pré - pintura': 'Pre-painting',
      'Pré pintura': 'Pre-painting',
      'Pintura final': 'Final painting',
      'Em andamento': 'In progress',
      'Em espera': 'On hold',
      'Aguardando': 'Awaiting',
      'Liberado': 'Released',
      'Bloqueado': 'Blocked',
      'Prazo médio': 'Average deadline',
      'Início': 'Start',
      'Término': 'Finish',
      'Título': 'Title',
      'Relatório Operacional STEP': 'STEP Operational Report',
      'Meta até hoje': 'Target to date',
      'Meta macro até hoje': 'Macro target to date',
      'carteira': 'portfolio',
      'unidade': 'unit',
      'concluído': 'completed',
      'fabricação': 'fabrication',
      'API do Portal do Cliente': 'Client Portal API',
      'Pesquisar QR Code': 'Search QR Code',
      'Rastreamento por ISO': 'ISO tracking',
      'Pesquise ISOs já criados automaticamente pelo cache e imprima etiquetas minimalistas.': 'Search ISOs automatically created by the cache and print minimal labels.',
      'Pesquisar ISO, BSP, Work Order, Vessel ou Tag': 'Search ISO, BSP, Work Order, Vessel or Tag',
      'Digite o ISO completo ou parte da BSP...': 'Enter the full ISO or part of the BSP...',
      'Pesquisar': 'Search',
      'Imprimir selecionados': 'Print selected',
      'Baixar selecionados': 'Download selected',
      'Digite uma busca ou clique em pesquisar para carregar os QR Codes recentes.': 'Type a search or click search to load recent QR Codes.',
      'Carregando QR Codes...': 'Loading QR Codes...',
      'Carregando QR Codes recentes...': 'Loading recent QR Codes...',
      'Nenhum QR Code encontrado. Se a BSP/ISO for nova, aguarde a atualização do cache ou clique em Atualizar agora.': 'No QR Code found. If the BSP/ISO is new, wait for the cache update or click Update now.',
      'Nenhum QR Code encontrado para essa busca.': 'No QR Code found for this search.',
      'QR Code automático': 'Automatic QR Code',
      'Visualizar': 'View',
      'Imprimir': 'Print',
      'Baixar': 'Download',
      'Baixar SVG': 'Download SVG',
      'Etiquetas QR Code ISO': 'ISO QR Code Labels',
      'Consultando QR Code no painel STEP.': 'Checking QR Code in the STEP dashboard.',
      'QR Code inválido': 'Invalid QR Code',
      'QR Code ativo': 'Active QR Code',
      'QR Code não localizado': 'QR Code not found',
      'Falha ao consultar QR Code.': 'Failed to check QR Code.',
      'Falha ao consultar QR Codes.': 'Failed to check QR Codes.',
      'Token do QR Code não informado.': 'QR Code token not provided.',
      'QR Code não encontrado.': 'QR Code not found.',
      'Falha ao gerar imagem do QR Code.': 'Failed to generate QR Code image.',
      'Falha ao abrir rastreamento do QR Code.': 'Failed to open QR Code tracking.',
      'Copiar token': 'Copy token',
      'Carregando chaves...': 'Loading keys...',
      'Nenhuma API criada ainda.': 'No API created yet.',
      'Imagens da BSP': 'BSP images',
      'Imagem anterior': 'Previous image',
      'Próxima imagem': 'Next image',
      'Carregando imagens…': 'Loading images…',
      'Nenhuma imagem encontrada.': 'No image found.',
      'Falha ao carregar imagens.': 'Failed to load images.',
      'Importar imagem da BSP': 'Import BSP image',
      'Importar imagem': 'Import image',
      'Selecione pelo menos uma imagem.': 'Select at least one image.',
      'imagem(ns) importada(s) com sucesso. Atualizando galeria...': 'image(s) imported successfully. Updating gallery...',
      'Falha ao importar imagem.': 'Failed to import image.',
      'Em atraso': 'Overdue',
      'No prazo': 'On time',
      'Em tratativa': 'Under review',
      'Engenharia em andamento': 'Engineering in progress',
      'Suprimento em andamento': 'Supply in progress',
      'Produção em andamento': 'Production in progress',
      'Qualidade em andamento': 'Quality in progress',
      'Pintura em andamento': 'Painting in progress',
      'Logística em andamento': 'Logistics in progress',
      'Caldeiraria em andamento': 'Boilermaking in progress',
      'Solda em andamento': 'Welding in progress',
      'Inspeção em andamento': 'Inspection in progress',
      'TH em andamento': 'Hydro Test in progress',
      'On Hold': 'On Hold',
      'Aguardando envio': 'Awaiting shipment',
      'Preparado para envio': 'Prepared for shipment',
      'Preparados para envio': 'Prepared for shipment',
      'Preparando para envio': 'Preparing for shipment',
      'Pendente': 'Pending',
      'Concluído': 'Completed',
      'Concluído.': 'Completed.',
      'Finalizado': 'Completed',
      'Finalizado com atraso': 'Completed late',
      'Em produção': 'In production',
      'Em qualidade': 'In quality',
      'Em pintura': 'In painting',
      'Não iniciado': 'Not started',
      'Sem unidade': 'No unit',
      'Sem conexão no momento.': 'No connection at the moment.',
      'Validando acesso...': 'Validating access...',
      'Falha ao autenticar.': 'Authentication failed.',
      'Acesso validado.': 'Access validated.',
      'Acesso validado. Carregando painel...': 'Access validated. Loading dashboard...',
      'Carregando painel...': 'Loading dashboard...',
      'Carregando dados operacionais...': 'Loading operational data...',
      'Dados carregados com sucesso. Abrindo o painel operacional.': 'Data loaded successfully. Opening the operational dashboard.',
      'Falha ao carregar painel.': 'Failed to load dashboard.',
      'Criando usuário...': 'Creating user...',
      'Salvando usuário...': 'Saving user...',
      'Salvar usuário': 'Save user',
      'Sincronizando com o GitHub...': 'Syncing with GitHub...',
      'Sincronizado com sucesso com o GitHub.': 'Successfully synced with GitHub.',
      'Falha ao sincronizar com o GitHub. Verifique GITHUB_TOKEN, GITHUB_REPO e GITHUB_BRANCH no Netlify.': 'Failed to sync with GitHub. Check GITHUB_TOKEN, GITHUB_REPO and GITHUB_BRANCH in Netlify.',
      'Alerta operacional enviado com sucesso.': 'Operational alert sent successfully.',
      'Enviando alerta...': 'Sending alert...',
      'Falha ao criar alerta operacional.': 'Failed to create operational alert.',
      'Digite a resposta antes de enviar.': 'Type the response before sending.',
      'Enviando resposta...': 'Sending response...',
      'Resposta enviada ao admin.': 'Response sent to admin.',
      'Falha ao enviar resposta.': 'Failed to send response.',
      'Preencha título e descrição da sinalização.': 'Fill in the flag title and description.',
      'Enviando sinalização ao PCP...': 'Sending flag to PCP...',
      'Sinalização enviada ao PCP.': 'Flag sent to PCP.',
      'Falha ao criar sinalização.': 'Failed to create flag.',
      'Observação opcional': 'Optional note',
      'Enviar em massa': 'Send in bulk',
      'Concluir OK': 'Complete OK',
      'Concluir revisão': 'Complete review',
      'Pendências de datas do histórico': 'History date pending items',
      'Nenhum apontamento pendente no momento.': 'No pending stage update at the moment.',
      'Nenhum histórico validado encontrado.': 'No validated history found.',
      'Selecionar todos os apontamentos visíveis': 'Select all visible stage updates',
      'Selecionar apontamento': 'Select stage update',
      'Selecionar todas as pendências visíveis': 'Select all visible pending items',
      'Revisão': 'Review',
      'Revisão concluída': 'Review completed',
      'Revisão PCP': 'PCP review',
      'Enviado': 'Sent',
      'Novo apontamento para validação': 'New stage update for validation',
      'Nova revisão para o PCP': 'New review for PCP',
      'Nova sinalização para o PCP': 'New flag for PCP',
      'Prazos em alerta': 'Deadlines in alert',
      'Alerta automático': 'Automatic alert',
      'Alerta operacional': 'Operational alert',
      'Projetos em desenvolvimento': 'Projects under development',
      'Montando Portal do Cliente...': 'Building Client Portal...',
      'Carregando dados do cliente...': 'Loading client data...',
      'Carregando BSPs...': 'Loading BSPs...',
      'Carregando POs...': 'Loading POs...',
      'Falha ao carregar dados administrativos:': 'Failed to load administrative data:',
      'Falha ao carregar projetos.': 'Failed to load projects.',
      'Falha ao carregar alertas operacionais.': 'Failed to load operational alerts.',
      'Falha ao carregar apontamentos setoriais.': 'Failed to load sector stage updates.',
      'Falha ao carregar usuários.': 'Failed to load users.',
      'Falha ao atualizar dados operacionais.': 'Failed to update operational data.',
      'Falha ao atualizar o Tracking.': 'Failed to update Tracking.',
      'Falha ao salvar as datas da ISO.': 'Failed to save ISO dates.',
      'Datas da ISO salvas com sucesso.': 'ISO dates saved successfully.',
      'A data de início não pode ser maior que a data de término.': 'The start date cannot be later than the finish date.',
      'Salvando datas da ISO...': 'Saving ISO dates...',
      'Painel do Cliente': 'Client Dashboard',
      'Painel Individual da Tag / ISO': 'Tag / ISO Individual Dashboard',
      'Visão Executiva da BSP': 'BSP Executive View',
      'Painel Individual STEP': 'STEP Individual Dashboard',
      'Curva S planejado versus realizado': 'Planned versus actual S-curve',
      'Curva S macro planejado versus realizado': 'Macro planned versus actual S-curve',
      'Curva S com rolagem horizontal': 'S-curve with horizontal scrolling',
      'Dados que alimentam a curva': 'Data feeding the curve',
      'Planejado + datas reais do Tracking quando preenchidas': 'Planned + actual Tracking dates when filled',
      'Atual estimado pelo progresso': 'Current estimate by progress',
      'Desvio de prazo': 'Schedule deviation',
      'Demanda atual': 'Current demand',
      'Etapas principais': 'Main stages',
      'Etapas principais da carteira': 'Main portfolio stages',
      'Etapas principais da unidade': 'Main unit stages',
      'Detalhamento macro das BSPs': 'Macro BSP details',
      'Detalhamento da unidade': 'Unit details',
      'BSPs por vessel': 'BSPs by vessel',
      'Clique em uma unidade para visualizar as BSPs vinculadas.': 'Click a unit to view linked BSPs.',
      'Clique uma vez para listar as BSPs; dê 2 cliques para abrir os gráficos da unidade': 'Click once to list BSPs; double-click to open the unit charts',
      'Clique uma vez para selecionar; dê 2 cliques para abrir a visão executiva': 'Click once to select; double-click to open the executive view',
      'Clique para abrir o painel individual da obra em uma nova aba': 'Click to open the work individual dashboard in a new tab',
      'Clique para abrir o painel individual e o detalhamento de Drawing em uma nova aba': 'Click to open the individual dashboard and Drawing details in a new tab',
      'Buscar as informações mais recentes do Smartsheet': 'Fetch the latest information from Smartsheet',
      'Fechar alerta': 'Close alert',
      'Fechar detalhamento': 'Close details',
      'Abrir alertas de prazo': 'Open deadline alerts',
      'Forçar atualização do Tracking': 'Force Tracking update',
      'Filtro de status': 'Status filter',
      'Filtros de visualização de projetos': 'Project view filters',
      'Mostrar senha': 'Show password',
      'Ocultar senha': 'Hide password',
      'Senha inicial': 'Initial password',
      'Ex.: pintura': 'Ex.: painting',
      'Digite a senha': 'Enter password',
      'Ex.: BSP 24-325-02, Prio, 2432502 ou ISO interno': 'Ex.: BSP 24-325-02, Prio, 2432502 or internal ISO',
      'Buscar PM na lista...': 'Search PM in the list...',
      'Descreva a sinalização para o PCP.': 'Describe the flag for PCP.',
      'Digite a resposta para o admin.': 'Type the response for the admin.',
      'Escreva a notificação do setor.': 'Write the sector notification.',
      'Ex.: BSP prioritário': 'Ex.: Priority BSP',
      'Nome do setor ou responsável': 'Sector or responsible name',
      'Nome da plataforma. Ex.: FORTE, FRADE, BRAVO': 'Platform name. Ex.: FORTE, FRADE, BRAVO',
      'Prévia e ajuste da logo do cliente': 'Client logo preview and adjustment',
      'URL da logo ou ./assets/client-logos/prio.png': 'Logo URL or ./assets/client-logos/prio.png',
      'Ex.: Projeto já enviado': 'Ex.: Project already shipped',

      'Portal do Cliente': 'Client Portal',
      'Demandas filtradas por empresa': 'Items filtered by company',
      'Atualizar': 'Update',
      'Abrir visão executiva da carteira': 'Open portfolio executive view',
      'Peso programado': 'Planned weight',
      'Peso soldado': 'Welded weight',
      'M² programada': 'Planned m²',
      'Progresso médio': 'Average progress',
      'Localizar BSP, PO ou Focal Point': 'Find BSP, PO or focal point',
      'Limpar busca': 'Clear search',
      'Vessels / Unidades': 'Vessels / Units',
      'Carteira por unidade': 'Portfolio by unit',
      'Clique uma vez para listar as BSPs; dê 2 cliques para abrir os gráficos e o PDF da unidade.': 'Click once to list the BSPs; double-click to open the unit charts and PDF.',
      'Ver todas': 'View all',
      'Demandas filtradas por empresa': 'Items filtered by company',
      'Peso programado por cliente': 'Planned weight by client',
      'Tags produção': 'Production tags',
      'Tags qualidade': 'Quality tags',
      'Tags pintura': 'Painting tags',
      'Área m²': 'Area m²',
      'Pendente solda': 'Pending welding',
      'Peso soldado por cliente': 'Welded weight by client',
      'Peso total por cliente': 'Total weight by client',
      'Total enviado': 'Total shipped',
      'Clique uma vez para listar as BSPs; dê 2 cliques para abrir os gráficos e o PDF da unit.': 'Click once to list the BSPs; double-click to open the unit charts and PDF.',
      'Digite BSP, PO ou nome do focal point. Ex.: Sergio Ramos': 'Enter BSP, PO or focal point name. Ex.: Sergio Ramos',
      'Cliente': 'Client',
      'Atualizado:': 'Updated:',
      'Atualização:': 'Update:',
      'Open executive view': 'Open executive view',
      'Projetos em desenvolvimento': 'Projects under development',
      'Vinculada ao cliente': 'Linked to the client',
      'Clique uma vez para listar as BSPs; dê 2 cliques para abrir os gráficos e o PDF da unidade': 'Click once to list the BSPs; double-click to open the unit charts and PDF.',
      'Clique uma vez para listar as BSPs; dê 2 cliques para abrir os gráficos da unidade.': 'Click once to list the BSPs; double-click to open the unit charts.',
      'Aberto': 'Open',
      'Cliente sem logo': 'Client without logo',
      'Cliente não identificado': 'Client not identified',
      'Sem dados': 'No data',
      'Sem dados disponíveis': 'No data available',
      'Atualização indisponível': 'Update unavailable',
      'Peso programado por unidade': 'Planned weight by unit',
      'Peso soldado por unidade': 'Welded weight by unit',
      'M² por unidade': 'm² by unit',
      'M²': 'm²',
      'Focal Point': 'Focal Point',
      'Nome do focal point': 'Focal point name',
      'Visualizar': 'View',
      'Detalhar': 'View details',
      'Resumo': 'Summary',
      'Sem BSPs nesta unidade.': 'No BSPs in this unit.',
      'Nenhuma unidade encontrada.': 'No unit found.',
      'Unidades': 'Units',
      'Lista de BSPs': 'BSP list',
      'Informações do cliente': 'Client information',
      'Dados do cliente': 'Client data',
      'Última atualização': 'Last update',
      'Última atualização do cache': 'Last cache update',
      'agora mesmo': 'just now',
      'aguardando rotina de atualização': 'waiting for refresh routine',
      'cache mantido; tentativa manual não concluída': 'cache kept; manual attempt not completed',
      'sem base salva ainda': 'no saved base yet',
      'Cache:': 'Cache:',
      'Cache': 'Cache',
      'Clique uma vez para selecionar; dê 2 cliques para abrir a visão executiva': 'Click once to select; double-click to open executive view',
      '2 cliques: gráficos / PDF da unidade': 'Double-click: unit charts / PDF',
      'Clique uma vez para listar as BSPs; dê 2 cliques para abrir os gráficos da unidade': 'Click once to list BSPs; double-click to open unit charts',
      'Clique uma vez para listar as BSPs; dê 2 cliques para abrir os gráficos e o PDF da unidade': 'Click once to list BSPs; double-click to open unit charts and PDF',
      'Atualização manual': 'Manual update',
      'Limpar filtros': 'Clear filters',
      'Sem resultados para a busca.': 'No results for the search.',
      'Clique uma vez para listar as BSPs; dê 2 cliques para abrir os gráficos e o PDF da unidade.': 'Click once to list the BSPs; double-click to open the unit charts and the unit PDF.',
      'Projeto já enviado': 'Project already shipped'
    },
    'es-ES': {
      'Idioma': 'Idioma',
      'Selecionar idioma do painel': 'Seleccionar idioma del panel',
      'Português': 'Portugués',
      'Inglês': 'Inglés',
      'Espanhol': 'Español',
      'PAINEL OPERACIONAL DE PROJETOS': 'PANEL OPERATIVO DE PROYECTOS',
      'STEP • Painel Operacional': 'STEP • Panel Operativo',
      'Preparando seu painel operacional': 'Preparando su panel operativo',
      'Carregando informações...': 'Cargando información...',
      'Estamos validando seu acesso.': 'Estamos validando su acceso.',
      'Aguarde, conectando aos dados operacionais.': 'Espere, conectando con los datos operativos.',
      'Acesso': 'Acceso',
      'Acesso bloqueado': 'Acceso bloqueado',
      'Faça login para liberar o painel, os alertas por setor e a administração.': 'Inicie sesión para liberar el panel, las alertas por sector y la administración.',
      'offline': 'sin conexión',
      'Login': 'Inicio de sesión',
      'Meus alertas': 'Mis alertas',
      'Minhas sinalizações': 'Mis avisos',
      'Alertas Projetos': 'Alertas de Proyectos',
      'Apontamentos': 'Avances',
      'Admin': 'Admin',
      'Sair': 'Salir',
      'Instalar app': 'Instalar app',
      'Gerar API': 'Generar API',
      'Alterar senha': 'Cambiar contraseña',
      'Brasil': 'Brasil',
      'Portugal': 'Portugal',
      '1. Total de projetos': '1. Total de proyectos',
      '2. Projetos iniciados': '2. Proyectos iniciados',
      '3. Projetos não iniciados': '3. Proyectos no iniciados',
      '4. Projetos em On Hold': '4. Proyectos en espera',
      '5. Projetos em produção': '5. Proyectos en producción',
      '6. Projetos em qualidade': '6. Proyectos en calidad',
      '7. Projetos em pintura': '7. Proyectos en pintura',
      '8. Projetos preparados para envio': '8. Proyectos preparados para envío',
      'Clique para detalhar': 'Haga clic para detallar',
      'Tags --': 'Tags --',
      'Apenas aguardando envio': 'Solo aguardando envío',
      'Peso total programado': 'Peso total programado',
      'Peso total soldado': 'Peso total soldado',
      'Total pendente de solda': 'Total pendiente de soldadura',
      '(M²) Programada': 'Área programada (m²)',
      'Total enviado -- kg': 'Total enviado -- kg',
      'Projetos por cliente': 'Proyectos por cliente',
      'Visão por cliente': 'Vista por cliente',
      'Voltar': 'Volver',
      'Projetos': 'Proyectos',
      'Alertas de prazo': 'Alertas de plazo',
      'Base operacional': 'Base operativa',
      'Aguardando atualização...': 'Esperando actualización...',
      'Atualizar agora': 'Actualizar ahora',
      'Buscar projeto, cliente, item ou referência': 'Buscar proyecto, cliente, ítem o referencia',
      'Filtrar demanda': 'Filtrar demanda',
      'Todas as demandas': 'Todas las demandas',
      'Filtrar tipo': 'Filtrar tipo',
      'Todos os tipos': 'Todos los tipos',
      'Filtrar semana': 'Filtrar semana',
      'Todas as semanas': 'Todas las semanas',
      'Filtrar status': 'Filtrar estado',
      'Todos os status': 'Todos los estados',
      'Limpar': 'Limpiar',
      'Baixar Excel': 'Descargar Excel',
      '-- resultado(s)': '-- resultado(s)',
      'Projeto': 'Proyecto',
      'Tipo': 'Tipo',
      'Cliente': 'Cliente',
      'Unidade': 'Unidad',
      'Término planejado': 'Fin planificado',
      'Itens': 'Ítems',
      'Peso soldado (kg)': 'Peso soldado (kg)',
      'Semana finalizado': 'Semana finalizada',
      'Peso (kg)': 'Peso (kg)',
      'Área operacional (m²)': 'Área operativa (m²)',
      'Etapa atual': 'Etapa actual',
      '% Individual': '% Individual',
      '% Geral': '% General',
      'Status': 'Estado',
      'Fab. início': 'Inicio fab.',
      'Caldeiraria': 'Calderería',
      'Calderaria': 'Calderería',
      'Solda': 'Soldadura',
      'Inspeção (QC)': 'Inspección (QC)',
      'TH': 'Prueba hidrostática',
      'Data de envio': 'Fecha de envío',
      'Concluído?': '¿Concluido?',
      'Carregando projetos...': 'Cargando proyectos...',
      'Detalhamento operacional': 'Detalle operativo',
      'Resumo rápido do projeto selecionado. Dê 2 cliques na linha para abrir o detalhamento completo.': 'Resumen rápido del proyecto seleccionado. Haga doble clic en la fila para abrir el detalle completo.',
      'Fechar': 'Cerrar',
      'Online': 'En línea',
      'Offline': 'Sin conexión',
      'Sistema online • sincronização operacional': 'Sistema en línea • sincronización operativa',
      'Versão dos dados: --': 'Versión de datos: --',
      'Itens internos e andamento por etapa': 'Ítems internos y avance por etapa',
      'Atenção': 'Atención',
      'Prazos em alerta': 'Plazos en alerta',
      'O alerta reaparece a cada 4 horas até que a planilha seja atualizada.': 'La alerta reaparece cada 4 horas hasta que la planilla sea actualizada.',
      'Acesso setorial': 'Acceso sectorial',
      'Entre com seu usuário e senha': 'Ingrese su usuario y contraseña',
      'Usuário': 'Usuario',
      'Senha': 'Contraseña',
      'Mostrar': 'Mostrar',
      'Entrar': 'Entrar',
      'Acesse com seu usuário setorial ou admin.': 'Acceda con su usuario sectorial o admin.',
      'Alerta Operacional': 'Alerta Operativa',
      'Meus alertas por setor': 'Mis alertas por sector',
      'Os alertas abaixo respeitam o setor do login atual.': 'Las alertas abajo respetan el sector del login actual.',
      'Administração': 'Administración',
      'Crie logins por setor, edite usuários e sincronize os dados com o repositório.': 'Cree accesos por sector, edite usuarios y sincronice los datos con el repositorio.',
      'Usuários e alertas setoriais': 'Usuarios y alertas sectoriales',
      'Novo usuário': 'Nuevo usuario',
      'Histórico operacional': 'Histórico operativo',
      'Novo alerta operacional': 'Nueva alerta operativa',
      'Registro de Alertas Operacionais': 'Registro de Alertas Operativas',
      'Pesquisar no histórico de alertas': 'Buscar en el histórico de alertas',
      'Subir pro GitHub': 'Subir a GitHub',
      'GitHub sync: offline': 'GitHub sync: sin conexión',
      'Perfil': 'Perfil',
      'Setor': 'Sector',
      'Nome': 'Nombre',
      'Nome exibido no Portal': 'Nombre mostrado en el Portal',
      'Cliente vinculado': 'Cliente vinculado',
      'Logo do cliente': 'Logo del cliente',
      'Importar logo do computador': 'Importar logo desde el equipo',
      'Resetar': 'Restablecer',
      'Aplicar ajuste': 'Aplicar ajuste',
      'Fotos por plataforma / vessel': 'Fotos por plataforma / vessel',
      'Adicionar foto dessa plataforma': 'Agregar foto de esta plataforma',
      'PMs adicionais para visualizar': 'PMs adicionales para visualizar',
      'Carregando nomes de PM do Tracking...': 'Cargando nombres de PM del Tracking...',
      'Nenhum PM adicional selecionado.': 'Ningún PM adicional seleccionado.',
      'Competências da Qualidade': 'Competencias de Calidad',
      'Engenharia': 'Ingeniería',
      'Suprimento': 'Suministro',
      'Pintura': 'Pintura',
      'Qualidade': 'Calidad',
      'Logística': 'Logística',
      'Produção': 'Producción',
      'PCP': 'PCP',
      'Final Inspection QC': 'Inspección final QC',
      'Inspeção Dimensional Inicial': 'Inspección dimensional inicial',
      'Inspeção Dimensional Final': 'Inspección dimensional final',
      'END / NDE': 'END / NDE',
      'Visualizar Painel do Cliente ao abrir uma BSP': 'Ver Panel del Cliente al abrir una BSP',
      'Criar usuário': 'Crear usuario',
      'Cancelar edição': 'Cancelar edición',
      'Usuários cadastrados': 'Usuarios registrados',
      'Nenhum usuário carregado.': 'Ningún usuario cargado.',
      'Monitor de acessos ao vivo': 'Monitor de accesos en vivo',
      'Setores monitorados': 'Sectores monitoreados',
      'Usuários online e última tela visualizada no sistema.': 'Usuarios en línea y última pantalla visualizada en el sistema.',
      '0 online': '0 en línea',
      'Nenhum usuário online agora.': 'Ningún usuario en línea ahora.',
      'Respostas dos alertas': 'Respuestas de las alertas',
      'Nenhuma resposta recebida ainda.': 'Ninguna respuesta recibida todavía.',
      'Título': 'Título',
      'Mensagem': 'Mensaje',
      'Prioridade': 'Prioridad',
      'Normal': 'Normal',
      'Alta': 'Alta',
      'Urgente': 'Urgente',
      'Baixa': 'Baja',
      'Exigir confirmação de leitura': 'Exigir confirmación de lectura',
      'Enviar alerta': 'Enviar alerta',
      'Resposta do alerta': 'Respuesta de la alerta',
      'Responder alerta': 'Responder alerta',
      'Escreva sua devolutiva para o administrador.': 'Escriba su devolución para el administrador.',
      'Sua resposta': 'Su respuesta',
      'Enviar resposta': 'Enviar respuesta',
      'Cancelar': 'Cancelar',
      'Sinalização ao PCP': 'Aviso al PCP',
      'Nova sinalização': 'Nuevo aviso',
      'A informação será enviada ao PCP.': 'La información será enviada al PCP.',
      'Descrição': 'Descripción',
      'Enviar ao PCP': 'Enviar al PCP',
      'Novo alerta recebido': 'Nueva alerta recibida',
      'Novo alerta': 'Nueva alerta',
      'Atenção imediata necessária.': 'Atención inmediata necesaria.',
      'Você recebeu uma nova notificação.': 'Usted recibió una nueva notificación.',
      'Abrir alerta': 'Abrir alerta',
      'Segurança de acesso': 'Seguridad de acceso',
      'Digite a senha atual e defina a nova senha de acesso.': 'Ingrese la contraseña actual y defina la nueva contraseña de acceso.',
      'Senha atual': 'Contraseña actual',
      'Nova senha': 'Nueva contraseña',
      'Confirmar nova senha': 'Confirmar nueva contraseña',
      'Salvar nova senha': 'Guardar nueva contraseña',
      'Apontamentos por etapa': 'Avances por etapa',
      'Apontamentos setoriais': 'Avances sectoriales',
      'Cada setor informa somente o avanço da sua própria etapa e o PCP valida.': 'Cada sector informa solamente el avance de su propia etapa y el PCP valida.',
      'Validação PCP / Apontamentos': 'Validación PCP / Avances',
      'Validação PCP': 'Validación PCP',
      'Todos': 'Todos',
      'Nenhum projeto encontrado para a busca informada.': 'No se encontró ningún proyecto para la búsqueda informada.',
      'Nenhuma tag detalhada encontrada.': 'No se encontró ninguna tag detallada.',
      'Nenhuma BSP encontrada para este cliente.': 'No se encontró ninguna BSP para este cliente.',
      'Nenhuma BSP nesta unidade.': 'Ninguna BSP en esta unidad.',
      'Atualizando...': 'Actualizando...',
      'Atualizando informações...': 'Actualizando información...',
      'Falha ao atualizar': 'Error al actualizar',
      'tente novamente': 'intente nuevamente',
      'Atualização: --': 'Actualización: --',
      'BSPs': 'BSPs',
      'abrir visão executiva': 'abrir vista ejecutiva',
      'Cronograma não disponível para esta ISO.': 'Cronograma no disponible para esta ISO.',
      'Schedule não disponível para esta BSP.': 'Schedule no disponible para esta BSP.',
      'Etapa': 'Etapa',
      'Prazo médio': 'Plazo medio',
      'Início': 'Inicio',
      'Término': 'Fin',
      'Relatório Operacional STEP': 'Informe Operativo STEP',
      'Meta até hoje': 'Meta hasta hoy',
      'Meta macro até hoje': 'Meta macro hasta hoy',
      'carteira': 'cartera',
      'unidade': 'unidad',
      'concluído': 'concluido',
      'fabricação': 'fabricación',
      'API do Portal do Cliente': 'API del Portal del Cliente',
      'Copiar token': 'Copiar token',
      'Carregando chaves...': 'Cargando claves...',
      'Nenhuma API criada ainda.': 'Ninguna API creada todavía.',
      'Imagens da BSP': 'Imágenes de la BSP',
      'Imagem anterior': 'Imagen anterior',
      'Próxima imagem': 'Próxima imagen',
      'Carregando imagens…': 'Cargando imágenes…',
      'Nenhuma imagem encontrada.': 'Ninguna imagen encontrada.',
      'Falha ao carregar imagens.': 'Error al cargar imágenes.',
      'Importar imagem da BSP': 'Importar imagen de la BSP',
      'Importar imagem': 'Importar imagen',
      'Selecione pelo menos uma imagem.': 'Seleccione al menos una imagen.',
      'imagem(ns) importada(s) com sucesso. Atualizando galeria...': 'imagen(es) importada(s) con éxito. Actualizando galería...',
      'Falha ao importar imagem.': 'Error al importar imagen.',
      'Em atraso': 'Atrasado',
      'No prazo': 'A tiempo',
      'Em tratativa': 'En tratativa',
      'Engenharia em andamento': 'Ingeniería en curso',
      'Suprimento em andamento': 'Suministro en curso',
      'Produção em andamento': 'Producción en curso',
      'Qualidade em andamento': 'Calidad en curso',
      'Pintura em andamento': 'Pintura en curso',
      'Logística em andamento': 'Logística en curso',
      'Caldeiraria em andamento': 'Calderería en curso',
      'Solda em andamento': 'Soldadura en curso',
      'Inspeção em andamento': 'Inspección en curso',
      'TH em andamento': 'Prueba hidráulica en curso',
      'On Hold': 'En espera',
      'Aguardando envio': 'Aguardando envío',
      'Preparado para envio': 'Preparado para envío',
      'Preparados para envio': 'Preparados para envío',
      'Preparando para envio': 'Preparando para envío',
      'Pendente': 'Pendiente',
      'Concluído': 'Concluido',
      'Concluído.': 'Concluido.',
      'Finalizado': 'Finalizado',
      'Finalizado com atraso': 'Finalizado con atraso',
      'Em produção': 'En producción',
      'Em qualidade': 'En calidad',
      'Em pintura': 'En pintura',
      'Não iniciado': 'No iniciado',
      'Sem unidade': 'Sin unidad',
      'Sem conexão no momento.': 'Sin conexión en este momento.',
      'Validando acesso...': 'Validando acceso...',
      'Falha ao autenticar.': 'Error al autenticar.',
      'Acesso validado.': 'Acceso validado.',
      'Acesso validado. Carregando painel...': 'Acceso validado. Cargando panel...',
      'Carregando painel...': 'Cargando panel...',
      'Carregando dados operacionais...': 'Cargando datos operativos...',
      'Dados carregados com sucesso. Abrindo o painel operacional.': 'Datos cargados con éxito. Abriendo el panel operativo.',
      'Falha ao carregar painel.': 'Error al cargar el panel.',
      'Criando usuário...': 'Creando usuario...',
      'Salvando usuário...': 'Guardando usuario...',
      'Salvar usuário': 'Guardar usuario',
      'Sincronizando com o GitHub...': 'Sincronizando con GitHub...',
      'Sincronizado com sucesso com o GitHub.': 'Sincronizado con éxito con GitHub.',
      'Falha ao sincronizar com o GitHub. Verifique GITHUB_TOKEN, GITHUB_REPO e GITHUB_BRANCH no Netlify.': 'Error al sincronizar con GitHub. Verifique GITHUB_TOKEN, GITHUB_REPO y GITHUB_BRANCH en Netlify.',
      'Alerta operacional enviado com sucesso.': 'Alerta operativa enviada con éxito.',
      'Enviando alerta...': 'Enviando alerta...',
      'Falha ao criar alerta operacional.': 'Error al crear alerta operativa.',
      'Digite a resposta antes de enviar.': 'Escriba la respuesta antes de enviar.',
      'Enviando resposta...': 'Enviando respuesta...',
      'Resposta enviada ao admin.': 'Respuesta enviada al admin.',
      'Falha ao enviar resposta.': 'Error al enviar respuesta.',
      'Preencha título e descrição da sinalização.': 'Complete el título y la descripción del aviso.',
      'Enviando sinalização ao PCP...': 'Enviando aviso al PCP...',
      'Sinalização enviada ao PCP.': 'Aviso enviado al PCP.',
      'Falha ao criar sinalização.': 'Error al crear aviso.',
      'Observação opcional': 'Observación opcional',
      'Enviar em massa': 'Enviar en masa',
      'Concluir OK': 'Concluir OK',
      'Concluir revisão': 'Concluir revisión',
      'Pendências de datas do histórico': 'Pendencias de fechas del histórico',
      'Nenhum apontamento pendente no momento.': 'Ningún avance pendiente en este momento.',
      'Nenhum histórico validado encontrado.': 'Ningún histórico validado encontrado.',
      'Selecionar todos os apontamentos visíveis': 'Seleccionar todos los avances visibles',
      'Selecionar apontamento': 'Seleccionar avance',
      'Selecionar todas as pendências visíveis': 'Seleccionar todas las pendencias visibles',
      'Revisão': 'Revisión',
      'Revisão concluída': 'Revisión concluida',
      'Revisão PCP': 'Revisión PCP',
      'Enviado': 'Enviado',
      'Novo apontamento para validação': 'Nuevo avance para validación',
      'Nova revisão para o PCP': 'Nueva revisión para el PCP',
      'Nova sinalização para o PCP': 'Nuevo aviso para el PCP',
      'Alerta automático': 'Alerta automática',
      'Alerta operacional': 'Alerta operativa',
      'Projetos em desenvolvimento': 'Proyectos en desarrollo',
      'Montando Portal do Cliente...': 'Montando Portal del Cliente...',
      'Carregando dados do cliente...': 'Cargando datos del cliente...',
      'Carregando BSPs...': 'Cargando BSPs...',
      'Carregando POs...': 'Cargando POs...',
      'Falha ao carregar dados administrativos:': 'Error al cargar datos administrativos:',
      'Falha ao carregar projetos.': 'Error al cargar proyectos.',
      'Falha ao carregar alertas operacionais.': 'Error al cargar alertas operativas.',
      'Falha ao carregar apontamentos setoriais.': 'Error al cargar avances sectoriales.',
      'Falha ao carregar usuários.': 'Error al cargar usuarios.',
      'Falha ao atualizar dados operacionais.': 'Error al actualizar datos operativos.',
      'Falha ao atualizar o Tracking.': 'Error al actualizar el Tracking.',
      'Falha ao salvar as datas da ISO.': 'Error al guardar las fechas de la ISO.',
      'Datas da ISO salvas com sucesso.': 'Fechas de la ISO guardadas con éxito.',
      'A data de início não pode ser maior que a data de término.': 'La fecha de inicio no puede ser mayor que la fecha de fin.',
      'Salvando datas da ISO...': 'Guardando fechas de la ISO...',
      'Painel do Cliente': 'Panel del Cliente',
      'Painel Individual da Tag / ISO': 'Panel Individual de Tag / ISO',
      'Visão Executiva da BSP': 'Vista Ejecutiva de la BSP',
      'Painel Individual STEP': 'Panel Individual STEP',
      'Curva S planejado versus realizado': 'Curva S planificada versus real',
      'Curva S macro planejado versus realizado': 'Curva S macro planificada versus real',
      'Curva S com rolagem horizontal': 'Curva S con desplazamiento horizontal',
      'Dados que alimentam a curva': 'Datos que alimentan la curva',
      'Planejado + datas reais do Tracking quando preenchidas': 'Planificado + fechas reales del Tracking cuando estén completadas',
      'Atual estimado pelo progresso': 'Actual estimado por avance',
      'Desvio de prazo': 'Desvío de plazo',
      'Demanda atual': 'Demanda actual',
      'Etapas principais': 'Etapas principales',
      'Etapas principais da carteira': 'Etapas principales de la cartera',
      'Etapas principais da unidade': 'Etapas principales de la unidad',
      'Detalhamento macro das BSPs': 'Detalle macro de las BSPs',
      'Detalhamento da unidade': 'Detalle de la unidad',
      'BSPs por vessel': 'BSPs por vessel',
      'Clique em uma unidade para visualizar as BSPs vinculadas.': 'Haga clic en una unidad para visualizar las BSPs vinculadas.',
      'Clique uma vez para listar as BSPs; dê 2 cliques para abrir os gráficos da unidade': 'Haga clic una vez para listar las BSPs; doble clic para abrir los gráficos de la unidad',
      'Clique uma vez para selecionar; dê 2 cliques para abrir a visão executiva': 'Haga clic una vez para seleccionar; doble clic para abrir la vista ejecutiva',
      'Clique para abrir o painel individual da obra em uma nova aba': 'Haga clic para abrir el panel individual de la obra en una nueva pestaña',
      'Clique para abrir o painel individual e o detalhamento de Drawing em uma nova aba': 'Haga clic para abrir el panel individual y el detalle de Drawing en una nueva pestaña',
      'Buscar as informações mais recentes do Smartsheet': 'Buscar la información más reciente de Smartsheet',
      'Fechar alerta': 'Cerrar alerta',
      'Fechar detalhamento': 'Cerrar detalle',
      'Abrir alertas de prazo': 'Abrir alertas de plazo',
      'Forçar atualização do Tracking': 'Forzar actualización del Tracking',
      'Filtro de status': 'Filtro de estado',
      'Filtros de visualização de projetos': 'Filtros de visualización de proyectos',
      'Mostrar senha': 'Mostrar contraseña',
      'Ocultar senha': 'Ocultar contraseña',
      'Senha inicial': 'Contraseña inicial',
      'Ex.: pintura': 'Ej.: pintura',
      'Digite a senha': 'Ingrese la contraseña',
      'Ex.: BSP 24-325-02, Prio, 2432502 ou ISO interno': 'Ej.: BSP 24-325-02, Prio, 2432502 o ISO interno',
      'Buscar PM na lista...': 'Buscar PM en la lista...',
      'Descreva a sinalização para o PCP.': 'Describa el aviso para el PCP.',
      'Digite a resposta para o admin.': 'Escriba la respuesta para el admin.',
      'Escreva a notificação do setor.': 'Escriba la notificación del sector.',
      'Ex.: BSP prioritário': 'Ej.: BSP prioritario',
      'Nome do setor ou responsável': 'Nombre del sector o responsable',
      'Nome da plataforma. Ex.: FORTE, FRADE, BRAVO': 'Nombre de la plataforma. Ej.: FORTE, FRADE, BRAVO',
      'Prévia e ajuste da logo do cliente': 'Vista previa y ajuste del logo del cliente',
      'URL da logo ou ./assets/client-logos/prio.png': 'URL del logo o ./assets/client-logos/prio.png',
      'Ex.: Projeto já enviado': 'Ej.: Proyecto ya enviado',

      'Portal do Cliente': 'Portal del Cliente',
      'Demandas filtradas por empresa': 'Demandas filtradas por empresa',
      'Atualizar': 'Actualizar',
      'Abrir visão executiva da carteira': 'Abrir vista ejecutiva de la cartera',
      'Peso programado': 'Peso programado',
      'Peso soldado': 'Peso soldado',
      'M² programada': 'M² programada',
      'Progresso médio': 'Progreso medio',
      'Localizar BSP, PO ou Focal Point': 'Localizar BSP, PO o focal point',
      'Limpar busca': 'Limpiar búsqueda',
      'Vessels / Unidades': 'Vessels / Unidades',
      'Carteira por unidade': 'Cartera por unidad',
      'Clique uma vez para listar as BSPs; dê 2 cliques para abrir os gráficos e o PDF da unidade.': 'Haga clic una vez para listar las BSPs; doble clic para abrir los gráficos y el PDF de la unidad.',
      'Ver todas': 'Ver todas',
      'Peso programado por cliente': 'Peso programado por cliente',
      'Tags produção': 'Tags producción',
      'Tags qualidade': 'Tags calidad',
      'Tags pintura': 'Tags pintura',
      'Área m²': 'Área m²',
      'Pendente solda': 'Pendiente soldadura',
      'Peso soldado por cliente': 'Peso soldado por cliente',
      'Peso total por cliente': 'Peso total por cliente',
      'Total enviado': 'Total enviado',
      'Digite BSP, PO ou nome do focal point. Ex.: Sergio Ramos': 'Escriba BSP, PO o nombre del focal point. Ej.: Sergio Ramos',
      'Atualizado:': 'Actualizado:',
      'Atualização:': 'Actualización:',
      'Projetos em desenvolvimento': 'Proyectos en desarrollo',
      'Clique uma vez para listar as BSPs; dê 2 cliques para abrir os gráficos e o PDF da unidade': 'Haga clic una vez para listar las BSPs; doble clic para abrir los gráficos y el PDF de la unidad.',
      'Clique uma vez para listar as BSPs; dê 2 cliques para abrir os gráficos da unidade.': 'Haga clic una vez para listar las BSPs; doble clic para abrir los gráficos de la unidad.',
      'Aberto': 'Abierto',
      'Cliente sem logo': 'Cliente sin logo',
      'Cliente não identificado': 'Cliente no identificado',
      'Sem dados': 'Sin datos',
      'Sem dados disponíveis': 'Sin datos disponibles',
      'Atualização indisponível': 'Actualización no disponible',
      'Peso programado por unidade': 'Peso programado por unidad',
      'Peso soldado por unidade': 'Peso soldado por unidad',
      'M² por unidade': 'M² por unidad',
      'M²': 'M²',
      'Visualizar': 'Visualizar',
      'Detalhar': 'Ver detalle',
      'Resumo': 'Resumen',
      'Sem BSPs nesta unidade.': 'No hay BSPs en esta unidad.',
      'Nenhuma unidade encontrada.': 'No se encontró ninguna unidad.',
      'Unidades': 'Unidades',
      'Lista de BSPs': 'Lista de BSPs',
      'Informações do cliente': 'Información del cliente',
      'Dados do cliente': 'Datos del cliente',
      'Última atualização': 'Última actualización',
      'Atualização manual': 'Actualización manual',
      'Limpar filtros': 'Limpiar filtros',
      'Sem resultados para a busca.': 'No hay resultados para la búsqueda.',

      'PESO': 'PESO',
      'SOLDADO': 'SOLDADO',
      'PESO SOLDADO': 'PESO SOLDADO',
      'Peso restante': 'Peso restante',
      'Tags totais': 'Tags totales',
      'Tags restantes': 'Tags restantes',
      'Curva S | Planejado x Realizado': 'Curva S | Planificado x Realizado',
      'Curva S | Carteira do Cliente': 'Curva S | Cartera del Cliente',
      'Curva S | Unidade': 'Curva S | Unidad',
      'Planejado x realizado consolidado das BSPs': 'Planificado x realizado consolidado de las BSPs',
      'Planejado': 'Planificado',
      'Realizado': 'Realizado',
      'Desvio': 'Desvío',
      'Report do Cliente': 'Informe del Cliente',
      'Dados do Tracking + Work in Progress exibidos dentro da visão principal': 'Datos de Tracking + Work in Progress mostrados dentro de la vista principal',
      'Baixar Excel do Cronograma': 'Descargar Excel del Cronograma',
      'Detalhamento da obra': 'Detalle de la obra',
      'Detalhamento das Tags / ISOs': 'Detalle de Tags / ISOs',
      'Processos da BSP e evolução por etapa': 'Procesos de la BSP y evolución por etapa',
      'Processos da BSP por etapa': 'Procesos de la BSP por etapa',
      'Baseada na data inicial e final da BSP': 'Basada en la fecha inicial y final de la BSP',
      'Menor progresso primeiro; finalizadas no final': 'Menor progreso primero; finalizadas al final',
      'Tag/ISO': 'Tag/ISO',
      'TAG/ISO': 'TAG/ISO',
      'Descrição': 'Descripción',
      'DESCRIÇÃO': 'DESCRIPCIÓN',
      'Observação': 'Observación',
      'OBSERVAÇÃO': 'OBSERVACIÓN',
      'Etapa': 'Etapa',
      'Peso': 'Peso',
      'AG. Emissão de detalhamento': 'Esperando emisión de detalle',
      'AG. emissão de detalhamento': 'Esperando emisión de detalle',
      'Emissão de detalhamento': 'Emisión de detalle',
      'Verificando estoque': 'Verificando stock',
      'Separação de material': 'Separación de material',
      'Aguardando END': 'Esperando END',
      'Aguardando NDE': 'Esperando NDE',
      'Unitização e Inspeção': 'Unitización e Inspección',
      'Unitização': 'Unitización',
      'Inspeção': 'Inspección',
      'Aguardando início de pintura': 'Esperando inicio de pintura',
      'Início de pintura': 'Inicio de pintura',
      'Aguardando início': 'Esperando inicio',
      'Package / Delivery sem progresso registrado.': 'Package / Delivery sin progreso registrado.',
      'Package / Delivery': 'Package / Delivery',
      'Preparado para envio': 'Preparado para envío',
      'Preenchido no Tracking': 'Rellenado en Tracking',
      'Dados do Tracking': 'Datos de Tracking',
      'Work in Progress': 'Work in Progress',
      'Visão principal': 'Vista principal',
      'Cronograma': 'Cronograma',
      'Cronograma da BSP': 'Cronograma de la BSP',
      'Término planejado': 'Fin planificado',
      'Início planejado': 'Inicio planificado',
      'Progresso planejado hoje': 'Progreso planificado hoy',
      'Marco / Observação': 'Hito / Observación',
      'Detalhamento': 'Detalle',
      'Detalhamento da BSP': 'Detalle de la BSP',
      'Detalhamento da unidade': 'Detalle de la unidad',
      'Detalhamento macro das BSPs': 'Detalle macro de las BSPs',
      'Carteira do Cliente': 'Cartera del Cliente',
      'Carteira por unidade': 'Cartera por unidad',
      'Dados que alimentam a curva': 'Datos que alimentan la curva',
      'Término replanejado': 'Fin replanificado',
      'Replanejado': 'Replanificado',
      'Fabricação abaixo do ritmo planejado para a data atual.': 'Fabricación por debajo del ritmo planificado para la fecha actual.',
      'Fabricação': 'Fabricación',
      'Pintura': 'Pintura',
      'Soldado': 'Soldado',
      'Programado': 'Programado',
      'Restante': 'Restante',
      'Total': 'Total',
      'STATUS': 'ESTADO',
      'STAGE': 'ETAPA',
      'DESCRIPTION': 'DESCRIPCIÓN',
      'Report': 'Informe',
      'Projeto já enviado': 'Proyecto ya enviado'
    }
  };


  const DICT_LOOKUP = Object.fromEntries(Object.entries(DICT).map(([lang, entries]) => {
    const map = new Map();
    Object.entries(entries).forEach(([key, value]) => {
      map.set(normalizeKey(key), value);
    });
    return [lang, map];
  }));

  function normalizeKey(value) {
    return String(value || '').replace(/\s+/g, ' ').trim().toLowerCase();
  }

  const textNodeSources = new WeakMap();
  const elementAttrSources = new WeakMap();
  let currentLang = readLanguage();
  let isApplying = false;
  let pendingApply = false;
  const pendingTranslationRoots = new Set();
  let observer = null;

  function readLanguage() {
    try {
      const stored = localStorage.getItem(STORAGE_KEY);
      if (SUPPORTED_LANGS.includes(stored)) return stored;
    } catch (_) {}
    return DEFAULT_LANG;
  }

  function saveLanguage(lang) {
    currentLang = SUPPORTED_LANGS.includes(lang) ? lang : DEFAULT_LANG;
    try { localStorage.setItem(STORAGE_KEY, currentLang); } catch (_) {}
    document.documentElement.lang = currentLang;
  }

  function normalizeText(value) {
    return String(value || '').replace(/\s+/g, ' ').trim();
  }

  function preserveOuterWhitespace(original, translated) {
    const start = String(original).match(/^\s*/)?.[0] || '';
    const end = String(original).match(/\s*$/)?.[0] || '';
    return `${start}${translated}${end}`;
  }

  function looksLikeOperationalCode(text) {
    const value = normalizeText(text);
    if (!value) return true;
    if (/^(BSP|ISO|PO|PM|TH|NDE|END|QC|STEP|URL|API|CPF|CNPJ)\b/i.test(value) && value.length <= 28) return true;
    if (/^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$/i.test(value)) return true;
    if (/^https?:\/\//i.test(value) || value.startsWith('./') || value.startsWith('/api/')) return true;
    if (/^[A-Z]{2,}[\w\-./ ]*\d[\w\-./ ]*$/i.test(value) && value.length <= 35) return true;
    if (/^[\d\s.,:%/\-–—()]+$/.test(value)) return true;
    return false;
  }

  function interpolateDynamic(text, lang) {
    if (lang === DEFAULT_LANG) return text;
    const normalized = normalizeText(text);
    const en = lang === 'en-US';

    let mDuration = normalized.match(/^agora mesmo$/i);
    if (mDuration) return en ? 'just now' : 'ahora mismo';

    mDuration = normalized.match(/^há\s+(\d+)\s+min$/i);
    if (mDuration) return en ? `${mDuration[1]} min ago` : `hace ${mDuration[1]} min`;

    mDuration = normalized.match(/^há\s+(\d+)h(?:\s*(\d+)min)?$/i);
    if (mDuration) return en ? `${mDuration[1]}h${mDuration[2] ? ` ${mDuration[2]}min` : ''} ago` : `hace ${mDuration[1]}h${mDuration[2] ? ` ${mDuration[2]}min` : ''}`;

    mDuration = normalized.match(/^há\s+(\d+)d(?:\s*(\d+)h)?$/i);
    if (mDuration) return en ? `${mDuration[1]}d${mDuration[2] ? ` ${mDuration[2]}h` : ''} ago` : `hace ${mDuration[1]}d${mDuration[2] ? ` ${mDuration[2]}h` : ''}`;

    let mCache = normalized.match(/^Última atualização do cache:\s*(.+?)(?:\s*•\s*(agora mesmo|há\s+\d+\s+min|há\s+\d+h(?:\s*\d+min)?|há\s+\d+d(?:\s*\d+h)?))?(?:\s*•\s*aguardando rotina de atualização)?$/i);
    if (mCache) {
      const age = mCache[2] ? translateText(mCache[2], lang) : '';
      const stale = /aguardando rotina de atualização/i.test(normalized);
      const suffix = [age, stale ? (en ? 'waiting for refresh routine' : 'esperando rutina de actualización') : ''].filter(Boolean).join(' • ');
      return `${en ? 'Last cache update' : 'Última actualización de caché'}: ${mCache[1]}${suffix ? ' • ' + suffix : ''}`;
    }

    let mQr = normalized.match(/^Pesquisando “(.+?)”\.\.\.$/i);
    if (mQr) return en ? `Searching “${mQr[1]}”...` : `Buscando “${mQr[1]}”...`;

    mQr = normalized.match(/^(\d+) QR Code\(s\) encontrado\(s\)\. Selecione para imprimir ou baixar\.$/i);
    if (mQr) return en ? `${mQr[1]} QR Code(s) found. Select to print or download.` : `${mQr[1]} QR Code(s) encontrado(s). Seleccione para imprimir o descargar.`;

    let m = normalized.match(/^Realizado\s+(.+?)\s+contra\s+planejado\s+(.+?):\s*desvio\s+de\s+(.+?)\.?$/i);
    if (m) return en ? `Actual ${m[1]} versus planned ${m[2]}: deviation of ${m[3]}.` : `Realizado ${m[1]} contra planificado ${m[2]}: desvío de ${m[3]}.`;

    m = normalized.match(/^Carteira realizada em\s+(.+?)\s+contra\s+planejado\s+(.+?):\s*desvio macro de\s+(.+?)\.?$/i);
    if (m) return en ? `Portfolio actual ${m[1]} versus planned ${m[2]}: macro deviation of ${m[3]}.` : `Cartera realizada en ${m[1]} contra planificado ${m[2]}: desvío macro de ${m[3]}.`;

    m = normalized.match(/^(.+?)\s+sem progresso registrado\.?$/i);
    if (m) return en ? `${translateText(m[1], lang)} has no recorded progress.` : `${translateText(m[1], lang)} sin progreso registrado.`;

    m = normalized.match(/^(\d+)\s+tag\(s\)\s+•\s+(.+?)\s+programado$/i);
    if (m) return en ? `${m[1]} tag(s) • ${m[2]} planned` : `${m[1]} tag(s) • ${m[2]} programado`;

    m = normalized.match(/^(.+?)\s+soldado\s+•\s+(.+?)$/i);
    if (m) return en ? `${m[1]} welded • ${m[2]}` : `${m[1]} soldado • ${m[2]}`;

    m = normalized.match(/^Término planejado:\s*(.+)$/i);
    if (m) return en ? `Planned finish: ${m[1]}` : `Fin planificado: ${m[1]}`;

    m = normalized.match(/^Término replanejado:\s*(.+)$/i);
    if (m) return en ? `Replanned finish: ${m[1]}` : `Fin replanificado: ${m[1]}`;

    m = normalized.match(/^(\d+) resultado\(s\)$/i);
    if (m) return en ? `${m[1]} result(s)` : `${m[1]} resultado(s)`;

    m = normalized.match(/^(.+?) projeto\(s\) na tela$/i);
    if (m) return en ? `${m[1]} project(s) on screen` : `${m[1]} proyecto(s) en pantalla`;

    m = normalized.match(/^(.+?) imagem\(ns\) disponível\(is\)$/i);
    if (m) return en ? `${m[1]} image(s) available` : `${m[1]} imagen(es) disponible(s)`;

    m = normalized.match(/^(.+?) tag\(s\) • (.+?) kg programado$/i);
    if (m) return en ? `${m[1]} tag(s) • ${m[2]} kg planned` : `${m[1]} tag(s) • ${m[2]} kg planificado`;

    m = normalized.match(/^(.+?) kg soldado • (.+?)$/i);
    if (m) return en ? `${m[1]} kg welded • ${m[2]}` : `${m[1]} kg soldado • ${m[2]}`;

    m = normalized.match(/^(\d+) online • (\d+) usuário\(s\)$/i);
    if (m) return en ? `${m[1]} online • ${m[2]} user(s)` : `${m[1]} en línea • ${m[2]} usuario(s)`;

    m = normalized.match(/^(\d+) online$/i);
    if (m) return en ? `${m[1]} online` : `${m[1]} en línea`;

    m = normalized.match(/^Atualização:\s*(.+)$/i);
    if (m) return en ? `Update: ${m[1]}` : `Actualización: ${m[1]}`;

    m = normalized.match(/^Falha ao atualizar:\s*(.+)$/i);
    if (m) return en ? `Failed to update: ${m[1]}` : `Error al actualizar: ${m[1]}`;

    m = normalized.match(/^(.+) BSP\(s\) vinculada\(s\) ao cliente$/i);
    if (m) return en ? `${m[1]} BSP(s) linked to the client` : `${m[1]} BSP(s) vinculada(s) al cliente`;

    m = normalized.match(/^(.+) • (.+) BSP\(s\)$/i);
    if (m) return `${m[1]} • ${m[2]} BSP(s)`;

    m = normalized.match(/^Responder:\s*(.+)$/i);
    if (m) return en ? `Reply: ${m[1]}` : `Responder: ${m[1]}`;

    m = normalized.match(/^Nova sinalização • (.+)$/i);
    if (m) return en ? `New flag • ${m[1]}` : `Nuevo aviso • ${m[1]}`;

    m = normalized.match(/^Preparando (\d+)\/(\d+):\s*(.+)$/i);
    if (m) return en ? `Preparing ${m[1]}/${m[2]}: ${m[3]}` : `Preparando ${m[1]}/${m[2]}: ${m[3]}`;

    m = normalized.match(/^Enviando (\d+)\/(\d+):\s*(.+)$/i);
    if (m) return en ? `Uploading ${m[1]}/${m[2]}: ${m[3]}` : `Enviando ${m[1]}/${m[2]}: ${m[3]}`;

    m = normalized.match(/^(\d+) imagem\(ns\) importada\(s\) com sucesso\. Atualizando galeria\.\.\.$/i);
    if (m) return en ? `${m[1]} image(s) imported successfully. Updating gallery...` : `${m[1]} imagen(es) importada(s) con éxito. Actualizando galería...`;

    m = normalized.match(/^Editando\s+(.+)\.$/i);
    if (m) return en ? `Editing ${m[1]}.` : `Editando ${m[1]}.`;

    m = normalized.match(/^Resolvida por:\s*(.+)$/i);
    if (m) return en ? `Resolved by: ${m[1]}` : `Resuelta por: ${m[1]}`;

    m = normalized.match(/^Setor:\s*(.+)$/i);
    if (m) return en ? `Sector: ${translateText(m[1], lang)}` : `Sector: ${translateText(m[1], lang)}`;

    m = normalized.match(/^Cliente:\s*(.+)$/i);
    if (m) return en ? `Client: ${m[1]}` : `Cliente: ${m[1]}`;

    m = normalized.match(/^Total:\s*(.+)$/i);
    if (m) return `Total: ${m[1]}`;

    return null;
  }

  function escapeRegExp(value) {
    return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function matchCaseReplacement(match, replacement) {
    if (!match || !replacement) return replacement;
    if (match === match.toUpperCase() && /[A-ZÀ-Ý]/.test(match)) return String(replacement).toUpperCase();
    return replacement;
  }

  function translateFragments(text, lang) {
    const dictionary = DICT[lang] || {};
    let value = text;
    const fragments = Object.keys(dictionary)
      .filter((key) => key.length >= 4 && !/[{}<>]/.test(key))
      .sort((a, b) => b.length - a.length);
    for (const key of fragments) {
      const replacement = dictionary[key];
      const directRegex = new RegExp(escapeRegExp(key), 'gi');
      value = value.replace(directRegex, (match) => matchCaseReplacement(match, replacement));
    }
    return value;
  }

  function translateText(originalText, lang = currentLang) {
    const original = String(originalText ?? '');
    const normalized = normalizeText(original);
    if (!normalized || lang === DEFAULT_LANG) return original;
    if (looksLikeOperationalCode(normalized)) return original;
    const dictionary = DICT[lang] || {};
    if (Object.prototype.hasOwnProperty.call(dictionary, normalized)) {
      return preserveOuterWhitespace(original, dictionary[normalized]);
    }
    const normalizedLookup = DICT_LOOKUP[lang]?.get(normalizeKey(normalized));
    if (normalizedLookup) {
      return preserveOuterWhitespace(original, normalizedLookup);
    }
    const dynamic = interpolateDynamic(normalized, lang);
    if (dynamic) return preserveOuterWhitespace(original, dynamic);
    if (normalized.length <= 260) {
      const translated = translateFragments(normalized, lang);
      if (translated !== normalized) return preserveOuterWhitespace(original, translated);
    }
    return original;
  }

  function shouldSkipElement(el) {
    if (!el || el.nodeType !== 1) return false;
    const tag = el.tagName;
    if (['SCRIPT', 'STYLE', 'NOSCRIPT', 'CODE', 'PRE', 'TEXTAREA'].includes(tag)) return true;
    if (el.closest('[data-i18n-skip], .client-api-code, .token-chip, .code-block')) return true;
    if (el.isContentEditable) return true;
    return false;
  }

  function translateTextNode(node) {
    const parent = node.parentElement;
    if (!parent || shouldSkipElement(parent)) return;
    const raw = String(node.nodeValue || '');
    if (!normalizeText(raw)) return;

    const knownSource = textNodeSources.get(node);
    const source = knownSource || raw;
    if (!knownSource) textNodeSources.set(node, source);
    const nextValue = currentLang === DEFAULT_LANG ? source : translateText(source, currentLang);
    if (node.nodeValue !== nextValue) node.nodeValue = nextValue;
  }

  function getAttrSourceMap(el) {
    let map = elementAttrSources.get(el);
    if (!map) {
      map = {};
      elementAttrSources.set(el, map);
    }
    return map;
  }

  function translateAttributes(el) {
    if (!el || shouldSkipElement(el)) return;
    const attrs = ['placeholder', 'title', 'aria-label'];
    const sourceMap = getAttrSourceMap(el);
    attrs.forEach((attr) => {
      if (!el.hasAttribute(attr)) return;
      const current = el.getAttribute(attr) || '';
      if (!normalizeText(current)) return;
      if (!sourceMap[attr]) sourceMap[attr] = current;
      const source = sourceMap[attr];
      const nextValue = currentLang === DEFAULT_LANG ? source : translateText(source, currentLang);
      if (current !== nextValue) el.setAttribute(attr, nextValue);
    });
  }

  function walkAndTranslate(root) {
    if (!root) return;
    if (root.nodeType === Node.TEXT_NODE) {
      translateTextNode(root);
      return;
    }
    if (root.nodeType !== Node.ELEMENT_NODE && root.nodeType !== Node.DOCUMENT_NODE && root.nodeType !== Node.DOCUMENT_FRAGMENT_NODE) return;
    const elementRoot = root.nodeType === Node.ELEMENT_NODE ? root : null;
    if (elementRoot && shouldSkipElement(elementRoot)) return;
    if (elementRoot) translateAttributes(elementRoot);

    const walker = document.createTreeWalker(root, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT, {
      acceptNode(node) {
        if (node.nodeType === Node.ELEMENT_NODE && shouldSkipElement(node)) return NodeFilter.FILTER_REJECT;
        return NodeFilter.FILTER_ACCEPT;
      }
    });
    let node = walker.currentNode;
    while (node) {
      if (node.nodeType === Node.TEXT_NODE) translateTextNode(node);
      else if (node.nodeType === Node.ELEMENT_NODE) translateAttributes(node);
      node = walker.nextNode();
    }
  }

  function scheduleApply(root = document.body) {
    if (root) pendingTranslationRoots.add(root);
    if (pendingApply || isApplying) return;
    pendingApply = true;
    requestAnimationFrame(() => {
      pendingApply = false;
      const roots = Array.from(pendingTranslationRoots);
      pendingTranslationRoots.clear();
      if (!roots.length) return;
      applyTranslationRoots(roots);
    });
  }

  function applyTranslations(root) {
    if (!root || isApplying) return;
    isApplying = true;
    try {
      document.documentElement.lang = currentLang;
      document.title = currentLang === DEFAULT_LANG ? 'STEP • Painel Operacional' : translateText('STEP • Painel Operacional', currentLang);
      walkAndTranslate(root);
      syncSwitcher();
    } finally {
      isApplying = false;
    }
  }

  function applyTranslationRoots(roots = []) {
    if (isApplying) return;
    const filtered = Array.from(new Set(roots.filter(Boolean))).filter((root) => {
      if (root === document.body || root === document.documentElement || root.nodeType === Node.DOCUMENT_NODE) return true;
      return !roots.some((other) => other && other !== root && other.contains && other.contains(root));
    });
    if (!filtered.length) return;
    isApplying = true;
    try {
      document.documentElement.lang = currentLang;
      document.title = currentLang === DEFAULT_LANG ? 'STEP • Painel Operacional' : translateText('STEP • Painel Operacional', currentLang);
      filtered.forEach((root) => walkAndTranslate(root));
      syncSwitcher();
    } finally {
      isApplying = false;
    }
  }

  function createSwitcher() {
    if (document.getElementById('language-switcher')) return;
    const wrapper = document.createElement('label');
    wrapper.className = 'language-switcher';
    wrapper.setAttribute('for', 'language-switcher');
    wrapper.innerHTML = `
      <span>Idioma</span>
      <select id="language-switcher" aria-label="Selecionar idioma do painel">
        ${SUPPORTED_LANGS.map((lang) => `<option value="${lang}">${LANGUAGE_LABELS[lang]}</option>`).join('')}
      </select>
    `;
    const target = document.querySelector('.topbar-actions.session-actions') || document.querySelector('.topbar-actions') || document.querySelector('.topbar') || document.body;
    if (target.classList && target.classList.contains('topbar-actions')) target.insertBefore(wrapper, target.firstChild);
    else target.appendChild(wrapper);
    const select = wrapper.querySelector('select');
    select.value = currentLang;
    select.addEventListener('change', (event) => {
      saveLanguage(event.target.value);
      applyTranslations(document.body);
      window.dispatchEvent(new CustomEvent('step:language-change', { detail: { language: currentLang } }));
    });
  }

  function syncSwitcher() {
    const select = document.getElementById('language-switcher');
    if (select && select.value !== currentLang) select.value = currentLang;
  }

  function initObserver() {
    if (observer) observer.disconnect();
    observer = new MutationObserver((mutations) => {
      if (isApplying || currentLang === DEFAULT_LANG) return;
      for (const mutation of mutations) {
        if (mutation.type === 'characterData') {
          textNodeSources.delete(mutation.target);
          scheduleApply(mutation.target);
          continue;
        }
        if (mutation.type === 'childList' && mutation.addedNodes.length) {
          mutation.addedNodes.forEach((node) => scheduleApply(node));
          continue;
        }
        if (mutation.type === 'attributes' && ['placeholder', 'title', 'aria-label'].includes(mutation.attributeName)) {
          const map = elementAttrSources.get(mutation.target);
          if (map) delete map[mutation.attributeName];
          scheduleApply(mutation.target);
        }
      }
    });
    observer.observe(document.body, {
      childList: true,
      subtree: true,
      characterData: true,
      attributes: true,
      attributeFilter: ['placeholder', 'title', 'aria-label']
    });
  }

  function init() {
    saveLanguage(currentLang);
    createSwitcher();
    applyTranslations(document.body);
    initObserver();
  }

  window.STEP_I18N = {
    get language() { return currentLang; },
    setLanguage(lang) {
      saveLanguage(lang);
      applyTranslations(document.body);
    },
    t(text, lang = currentLang) {
      return translateText(text, lang);
    },
    refresh() {
      applyTranslations(document.body);
    },
    supported: SUPPORTED_LANGS.slice(),
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init, { once: true });
  } else {
    init();
  }
})();
