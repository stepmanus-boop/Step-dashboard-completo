const required = ["SMARTSHEET_TOKEN"];
const missing = required.filter((key) => !process.env[key]);

if (missing.length) {
  console.error("Faltam variáveis obrigatórias:", missing.join(", "));
  process.exit(1);
}

console.log("Configuração mínima encontrada.");
