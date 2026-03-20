# 🚂 Sistema de Controle de Manobra - RUMO Logística

Uma aplicação web desenvolvida para monitorar, analisar e exportar dados de performance operacional das manobras ferroviárias (Encoste e Retirada de vagões) em múltiplos terminais.

---

## 🚀 Funcionalidades Principais

* **Registro Operacional Detalhado:** Lançamento rápido de dados de *Previsto vs. Realizado* (Horários e Vagões) por terminal, definindo o tipo de operação (Carga, Descarga ou Descarga com reaproveitamento) e o produto.
* **Gestão de Turnos:** Seleção clara do turno da equipe (A, B ou C), que acompanha os dados em todo o fluxo até o relatório final.
* **Mapeamento de Causas de Atraso:** Campo de diagnóstico opcional para registrar os motivos de não cumprimento das metas (ex: Clima, Manutenção, Falha na Locomotiva).
* **Dashboard Analítico Inteligente:** Visualização imediata dos resultados em gráficos interativos:
  * Comparativo de aderência geral (Encoste vs. Retirada).
  * Aderência de horários (Verde/Vermelho).
  * Volume de movimentação por turno.
  * *Ranking* das principais causas de atraso (Barras horizontais).
* **Exportação Avançada (Excel):** Geração de relatórios `.xlsx` formatados automaticamente. O sistema calcula as porcentagens, mescla as células para facilitar a leitura e destaca em vermelho as aderências abaixo de 100%.
* **Comparador de Turnos (Módulo BI):** Ferramenta onde o usuário faz o upload de múltiplos relatórios Excel gerados pelo sistema. O aplicativo lê os dados, calcula a média geral e gera um gráfico destacando (em dourado) a equipe vencedora.
* **Versão Portátil / Executável (.exe):** O sistema inteiro pode ser empacotado em um único arquivo executável para Windows. O usuário final não precisa instalar programação, Node.js ou configurar servidores; basta dar dois cliques e usar.

---

## 🛠️ Tecnologias Utilizadas

* **Backend:** Node.js, Express
* **Processamento de Planilhas:** ExcelJS, Multer
* **Frontend:** HTML5, CSS3, Bootstrap 5
* **Gráficos:** Chart.js + chartjs-plugin-datalabels
* **Empacotamento:** `pkg` (Node.js)