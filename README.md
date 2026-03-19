# 🚂 Controle de Manobra - RUMO Logística

Uma aplicação web full-stack desenvolvida em **Node.js** para monitorar, calcular e exportar o desempenho operacional de manobras ferroviárias (Encoste e Retirada de vagões) em múltiplos terminais. O sistema automatiza a medição de aderência das equipes às metas de horários e volume de vagões, gerando relatórios em Excel e permitindo a comparação de resultados entre turnos.

---

## 🚀 Funcionalidades Principais e Arquitetura

O sistema é dividido em uma interface interativa (HTML/JS) e um backend em Express que processa cálculos, gera arquivos e lê planilhas.

### 1. Regras de Negócio e Cálculo de Aderência
O principal do sistema. As funções de cálculo rodam no backend para garantir a integridade dos dados, avaliando a performance real contra a prevista.
* **Horário:** Tolerância rígida de +/- 1 hora (dentro da tolerância = 100%, fora = 0%).
* **Vagões:** Se o realizado for maior ou igual ao previsto = 100%. Se for menor, calcula-se a proporção.

```javascript
// script.js (Backend) - Lógica de Avaliação
function calcHora(prev, real) {
    if (!prev || !real) return { val: 0, text: '-' };
    const p = new Date(`1970-01-01T${prev}:00`);
    const r = new Date(`1970-01-01T${real}:00`);
    const diff = (r - p) / 36e5; // Calcula a diferença em horas

    if (diff >= -1 && diff <= 1) return { val: 100, text: '100%' };
    return { val: 0, text: '0%' };
}

function calcVags(prev, real) {
    const p = parseInt(prev) || 0;
    const r = parseInt(real) || 0;
    if (p === 0) return { val: 100, text: '100%' };
    if (r >= p) return { val: 100, text: '100%' };
    
    const pct = (r / p) * 100;
    return { val: pct, text: pct.toFixed(0) + '%' };
}
```

### 2. Captura Dinâmica de Dados (Frontend)
A interface (`index.html`) gera cards para cada terminal dinamicamente e captura os dados submetidos, agrupando o "Tipo de Operação" e o "Produto" em uma única variável antes de enviar ao servidor.

```javascript
// index.html - Processamento do Formulário
document.getElementById('formManobra').addEventListener('submit', (e) => {
    e.preventDefault();
    const formData = new FormData(e.target);
    const dadosParaEnviar = [];
    const turno = document.getElementById('turnoSelect').value;

    for (let i = 0; i < totalTerminais; i++) {
        const nome = formData.get(`nome_${i}`);
        const tipoOp = formData.get(`tipo_operacao_${i}`); // Carga, Descarga, e Descarga com reaproveitamento
        const conteudo = formData.get(`conteudo_${i}`);
        const cargaCompleta = conteudo ? `${tipoOp}: ${conteudo}` : tipoOp;
        
        // Estrutura os dados de Encoste e Retirada
        const encHp = formData.get(`enc_hp_${i}`);
        if(encHp) {
            dadosParaEnviar.push({
                terminal: nome, atividade: 'Encoste',
                hPrev: encHp, hReal: formData.get(`enc_hr_${i}`),
                vPrev: formData.get(`enc_vp_${i}`), vReal: formData.get(`enc_vr_${i}`),
                carga: cargaCompleta
            });
        }
        // (Lógica espelhada para Retirada omitida para brevidade)
    }
    
    localStorage.setItem('turnoSelecionado', turno);
    localStorage.setItem('dadosManobra', JSON.stringify(dadosParaEnviar));
    window.location.href = 'graficos.html';
});
```

### 3. Geração de Relatórios Excel
Utilizando a biblioteca `ExcelJS`, o backend recebe os dados, aplica os cálculos de aderência, mescla células, insere cores de destaque e gera uma planilha `.xlsx` pronta para download. O turno é destacado logo na primeira linha.

```javascript
// script.js (Backend) - Exportação de Planilha
app.post('/gerar', async (req, res) => {
    const dados = req.body;
    const turno = req.query.turno || '-';
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet(`Turno ${turno}`);

    // Criação do cabeçalho mestre identificando o turno
    sheet.spliceRows(1, 0, []);
    sheet.getCell('A1').value = `RELATÓRIO OPERACIONAL - TURNO: ${turno}`;
    sheet.mergeCells('A1:I1');
    sheet.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00345E' } };
    
    // Configuração de colunas e iteração de terminais (omitido para brevidade)
    // ...
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=Grade_Manobra_Turno${turno}.xlsx`);
    await workbook.xlsx.write(res);
    res.end();
});
```

### 4. Business Intelligence: Comparação de Turnos
O usuário pode fazer upload de múltiplas planilhas previamente geradas. O sistema processa os arquivos em memória (via `multer`), localiza a coluna de "Aderência Geral" e extrai a média para determinar a melhor equipe.

```javascript
// script.js (Backend) - Leitura e Comparação de Excel
const upload = multer({ storage: multer.memoryStorage() });

app.post('/api/comparar', upload.array('planilhas', 5), async (req, res) => {
    const resultados = [];

    for (const file of req.files) {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(file.buffer);
        const sheet = workbook.worksheets[0];

        let somaAderencia = 0, qtdValores = 0;
        let turnoNome = "Desconhecido";

        // Identifica o turno lendo a célula A1
        const header = sheet.getCell('A1').value;
        if (header && header.includes('TURNO:')) turnoNome = header.split('TURNO:')[1].trim();

        // Lê a coluna H (índice 8) buscando porcentagens
        sheet.eachRow((row) => {
            const val = row.getCell(8).value;
            if (typeof val === 'string' && val.includes('%')) {
                somaAderencia += parseFloat(val.replace('%', ''));
                qtdValores++;
            }
        });

        resultados.push({
            turno: turnoNome,
            media: parseFloat((qtdValores > 0 ? somaAderencia / qtdValores : 0).toFixed(1))
        });
    }

    resultados.sort((a, b) => a.turno.localeCompare(b.turno));
    res.json(resultados);
});
```

*Desenvolvido para otimização e controle operacional logístico.*