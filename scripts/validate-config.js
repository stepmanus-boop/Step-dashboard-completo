const required = ['SUPABASE_URL', 'SUPABASE_SERVICE_ROLE_KEY'];
const smartsheetKeys = ['SMARTSHEET_TOKEN', 'SMARTSHEET_ACCESS_TOKEN', 'SMARTSHEET_API_TOKEN', 'SMARTSHEET_BEARER_TOKEN', 'SMARTSHEET_PAT', 'SMARTSHEET_PERSONAL_ACCESS_TOKEN'];
const missing = required.filter((name) => !process.env[name]);
const hasSmartsheetEnv = smartsheetKeys.some((name) => !!process.env[name]);
if (!hasSmartsheetEnv) {
  console.warn('Aviso: token Smartsheet não veio por variável de ambiente. O projeto usará a configuração interna existente no ZIP. Recomenda-se migrar para SMARTSHEET_TOKEN no Netlify.');
}
if (missing.length) {
  console.error(`Variáveis obrigatórias ausentes: ${missing.join(', ')}`);
  process.exit(1);
}
console.log('Configuração mínima OK.');
