v38.0 - Correção definitiva isUniversalAccess

Correção aplicada:
- O cadastro não depende mais de variável de escopo isUniversalAccess.
- A validação calcula diretamente isUniversalAccessInput(role, sector, alertSectors).
- operationRegion/siteKey são enviados diretamente como BR/PT.
- Admin/PCP continuam universais pela regra.
