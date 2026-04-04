const express = require('express');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');
const multer = require('multer');
const path = require('path');

const { exec } = require('child_process');

const app = express();
const port = 3000;

const upload = multer({ storage: multer.memoryStorage() });

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(express.static('public'));

app.use(express.static(path.join(__dirname, 'public')));

const terminais = [ "Yara", "Bianchini", "Cotribá", "Ceifagro", "Agrofel", "Pradozem", "Três Tentos", "Recebimento Trem", "Formação Trem" ];

app.get('/api/terminais', (req, res) => res.json(terminais));

// --- FUNÇÕES DE CÁLCULO ---
function calcHora(prev, real) {
    if (!prev || !real) return { val: 0, text: '-' };
    const p = new Date(`1970-01-01T${prev}:00`);
    const r = new Date(`1970-01-01T${real}:00`);
    const diff = (r - p) / 36e5;

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

function identificarTurno(horaStr) {
    if (!horaStr) return null;
    const hora = parseInt(horaStr.split(':')[0]);
    
    if (hora >= 7 && hora < 15) return 'T1';
    if (hora >= 15 && hora < 23) return 'T2';
    return 'T3';
}

// GRÁFICOS
app.post('/api/calcular', (req, res) => {
    const dados = req.body;
    const resultadosPorTerminal = [];
    const turnos = { 'T1': 0, 'T2': 0, 'T3': 0 };
    const contagemMotivos = {};

    const nomesTerminais = [...new Set(dados.map(d => d.terminal))];

    nomesTerminais.forEach(nome => {
        const ativsTerminal = dados.filter(d => d.terminal === nome);
        
        let somaEnc = 0, countEnc = 0;
        let somaRet = 0, countRet = 0;
        let somaHora = 0, countHora = 0;

        ativsTerminal.forEach(ativ => {
            if (ativ.hReal) { const t = identificarTurno(ativ.hReal); if(t) turnos[t]++; }
            const motivoAtual = ativ.motivo;
            if (motivoAtual && motivoAtual !== "") {
                contagemMotivos[motivoAtual] = (contagemMotivos[motivoAtual] || 0) + 1;
            }

            const resH = calcHora(ativ.hPrev, ativ.hReal);
            const resV = calcVags(ativ.vPrev, ativ.vReal);
            const media = (resH.val + resV.val) / 2;

            if (ativ.tipo_base === 'Encoste') { somaEnc += media; countEnc++; }
            else if (ativ.tipo_base === 'Retirada') { somaRet += media; countRet++; }

            if (ativ.hReal) { somaHora += resH.val; countHora++; }
        });

        resultadosPorTerminal.push({
            terminal: nome,
            mediaEnc: countEnc > 0 ? parseFloat((somaEnc / countEnc).toFixed(1)) : 0,
            mediaRet: countRet > 0 ? parseFloat((somaRet / countRet).toFixed(1)) : 0,
            mediaHora: countHora > 0 ? parseFloat((somaHora / countHora).toFixed(1)) : 0
        });
    });

    res.json({ terminais: resultadosPorTerminal, movimentacaoTurnos: turnos, motivos: contagemMotivos });
});

// PLANILHA
app.post('/gerar', async (req, res) => {
    try {
        const dados = req.body;
        const turno = req.query.turno || '-';
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet(`Turno ${turno}`);

        sheet.columns = [
            { header: 'Terminal', key: 'A', width: 20 },
            { header: 'Atividade', key: 'B', width: 15 },
            { header: 'Tipo', key: 'C', width: 10 },
            { header: 'Horário', key: 'D', width: 12 },
            { header: 'Vagões', key: 'E', width: 10 },
            { header: 'Aderência Horário', key: 'F', width: 15 },
            { header: 'Aderência Vagões', key: 'G', width: 15 },
            { header: 'Aderência Geral', key: 'H', width: 12 },
            { header: 'Carga', key: 'I', width: 30 },
            { header: 'Motivo de Atraso', key: 'J', width: 25 }
        ];

        sheet.spliceRows(1, 0, []);
        sheet.getCell('A1').value = `RELATÓRIO OPERACIONAL - TURNO: ${turno}`;
        sheet.mergeCells('A1:J1');
        sheet.getCell('A1').font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 14 };
        sheet.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00345E' } };
        sheet.getCell('A1').alignment = { vertical: 'middle', horizontal: 'center' };
        sheet.getRow(1).height = 30;

        const headerRow = sheet.getRow(2);
        headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
        headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
        headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
        headerRow.height = 25;

        const nomesTerminais = [...new Set(dados.map(d => d.terminal))];
        let linhaAtual = 3;

        nomesTerminais.forEach(nome => {
            const ativsTerminal = dados.filter(d => d.terminal === nome);
            if (ativsTerminal.length === 0) return;

            const linhaInicialDoTerminal = linhaAtual;

            ativsTerminal.forEach(ativ => {
                const resH = calcHora(ativ.hPrev, ativ.hReal);
                const resV = calcVags(ativ.vPrev, ativ.vReal);
                const media = (resH.val + resV.val) / 2;
                const cargaTexto = ativ.carga || '-';
                const motivoTexto = ativ.motivo || '';

                // Valores
                sheet.getRow(linhaAtual).values = [ nome, ativ.atividade, 'Prev', ativ.hPrev || '-', ativ.vPrev || '-', resH.text, resV.text, media.toFixed(0) + '%', cargaTexto, motivoTexto ];
                sheet.getRow(linhaAtual + 1).values = [ null, null, 'Real', ativ.hReal || '-', ativ.vReal || '-', null, null, null, null, null ];

                // Cores de reprovação
                if (media < 100) sheet.getCell(`H${linhaAtual}`).font = { color: { argb: 'FFFF0000' }, bold: true };

                const boxBorder = { style: 'thin', color: { argb: 'FF000000' } };
                for(let r = 0; r < 2; r++) {
                    const rowObj = sheet.getRow(linhaAtual + r);
                    rowObj.alignment = { vertical: 'middle', horizontal: 'center' };
                    if (r === 1) rowObj.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
                    for(let c = 1; c <= 10; c++) {
                        sheet.getCell(linhaAtual + r, c).border = { top: boxBorder, left: boxBorder, bottom: boxBorder, right: boxBorder };
                    }
                }

                sheet.mergeCells(`B${linhaAtual}:B${linhaAtual + 1}`);
                sheet.mergeCells(`F${linhaAtual}:F${linhaAtual + 1}`);
                sheet.mergeCells(`G${linhaAtual}:G${linhaAtual + 1}`);
                sheet.mergeCells(`H${linhaAtual}:H${linhaAtual + 1}`);

                linhaAtual += 2;
            });

            sheet.mergeCells(`A${linhaInicialDoTerminal}:A${linhaAtual - 1}`);
            sheet.mergeCells(`I${linhaInicialDoTerminal}:I${linhaAtual - 1}`);
            sheet.mergeCells(`J${linhaInicialDoTerminal}:J${linhaAtual - 1}`);
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=Grade_Manobra_Turno_${turno}.xlsx`);
        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        console.error(error);
        res.status(500).send("Erro");
    }
});

//COMPARAR PLANILHAS
app.post('/api/comparar', upload.array('planilhas', 5), async (req, res) => {
    try {
        const resultadosTurnos = [];
        const consolidadoMotivos = {};

        for (const file of req.files) {
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(file.buffer);
            const sheet = workbook.worksheets[0];

            let somaAderencia = 0;
            let qtdValores = 0;
            let turnoNome = "Desconhecido";

            const header = sheet.getCell('A1').value;
            if (header && typeof header === 'string' && header.includes('TURNO:')) {
                turnoNome = header.split('TURNO:')[1].trim();
            }

            sheet.eachRow((row, rowNumber) => {
                if (rowNumber <= 2) return;

                const valAderencia = row.getCell(8).value;
                if (typeof valAderencia === 'string' && valAderencia.includes('%')) {
                    const numero = parseFloat(valAderencia.replace('%', ''));
                    if (!isNaN(numero)) {
                        somaAderencia += numero;
                        qtdValores++;
                    }
                }

                const valMotivo = row.getCell(10).value;
                if (valMotivo && typeof valMotivo === 'string' && valMotivo !== "" && valMotivo !== "-") {
                    consolidadoMotivos[valMotivo] = (consolidadoMotivos[valMotivo] || 0) + 1;
                }
            });

            const mediaFinal = qtdValores > 0 ? (somaAderencia / qtdValores) : 0;
            resultadosTurnos.push({
                turno: turnoNome,
                media: parseFloat(mediaFinal.toFixed(1))
            });
        }

        resultadosTurnos.sort((a, b) => a.turno.localeCompare(b.turno));

        res.json({
            turnos: resultadosTurnos,
            motivosConsolidados: consolidadoMotivos
        });
    } catch (error) {
        console.error("Erro ao ler planilhas:", error);
        res.status(500).json({ erro: "Falha ao processar os arquivos." });
    }
});

app.listen(port, () => {
    console.log(`Servidor rodando! Acesse: http://localhost:${port}`);
    exec(`start http://localhost:${port}`);
});