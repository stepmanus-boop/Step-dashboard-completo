STEP Dashboard v37.75

Correções desta versão:
- Impressão Zebra ZD230-203dpi ZPL: imprimir selecionados agora gera uma única arte SVG consolidada por faixa, com 3 quadros fixos. Isso evita o problema do Chrome/Zebra imprimir somente o primeiro QR quando há 3 selecionados.
- Baixar selecionados agora baixa a mesma arte Zebra consolidada por grupo de 3, em vez de baixar apenas arquivos individuais soltos.
- Mantido padrão da etiqueta: QR Code + nome completo do ISO.
- Etapa Atual: se a BSP estiver sinalizada como ON HOLD no Smartsheet, a Etapa Atual passa a exibir On Hold. Somente BSPs com sinalização ON HOLD são afetadas.
- Service Worker atualizado para v37.75.

Configuração recomendada na impressão do Chrome:
- Destino: ZDesigner ZD230-203dpi ZPL / Zebra Adm
- Tamanho do papel: Custom da etiqueta
- Margens: Nenhuma, se disponível
- Escala: 100
