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

// API GRÁFICOS
app.post('/api/calcular', (req, res) => {
    const dados = req.body;
    const resultadosPorTerminal = [];
    
    const turnos = { 'T1': 0, 'T2': 0, 'T3': 0 };

    const nomesTerminais = [...new Set(dados.map(d => d.terminal))];

    nomesTerminais.forEach(nome => {
        const enc = dados.find(d => d.terminal === nome && d.atividade === 'Encoste') || {};
        const ret = dados.find(d => d.terminal === nome && d.atividade === 'Retirada') || {};

        if (enc.hReal) { const t = identificarTurno(enc.hReal); if(t) turnos[t]++; }
        if (ret.hReal) { const t = identificarTurno(ret.hReal); if(t) turnos[t]++; }

        const resEncH = calcHora(enc.hPrev, enc.hReal);
        const resEncV = calcVags(enc.vPrev, enc.vReal);
        const mediaEnc = (resEncH.val + resEncV.val) / 2;

        const resRetH = calcHora(ret.hPrev, ret.hReal);
        const resRetV = calcVags(ret.vPrev, ret.vReal);
        const mediaRet = (resRetH.val + resRetV.val) / 2;

        let mediaHora = 0;
        let countH = 0;
        if (enc.hReal) { mediaHora += resEncH.val; countH++; }
        if (ret.hReal) { mediaHora += resRetH.val; countH++; }
        if (countH > 0) mediaHora = mediaHora / countH;

        resultadosPorTerminal.push({
            terminal: nome,
            mediaEnc: parseFloat(mediaEnc.toFixed(1)),
            mediaRet: parseFloat(mediaRet.toFixed(1)),
            mediaHora: parseFloat(mediaHora.toFixed(1))
        });
    });

    res.json({
        terminais: resultadosPorTerminal,
        movimentacaoTurnos: turnos
    });
});

// --- PLANILHA ---
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
            { header: 'Carga', key: 'I', width: 30 }
        ];

        sheet.spliceRows(1, 0, []);
        sheet.getCell('A1').value = `RELATÓRIO OPERACIONAL - TURNO: ${turno}`;
        sheet.mergeCells('A1:I1');
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
            const enc = dados.find(d => d.terminal === nome && d.atividade === 'Encoste') || {};
            const ret = dados.find(d => d.terminal === nome && d.atividade === 'Retirada') || {};

            const resEncH = calcHora(enc.hPrev, enc.hReal);
            const resEncV = calcVags(enc.vPrev, enc.vReal);
            const mediaEnc = (resEncH.val + resEncV.val) / 2;

            const resRetH = calcHora(ret.hPrev, ret.hReal);
            const resRetV = calcVags(ret.vPrev, ret.vReal);
            const mediaRet = (resRetH.val + resRetV.val) / 2;

            const cargaTexto = enc.carga || ret.carga || '-';

            // 1. INSERIR VALORES
            sheet.getRow(linhaAtual).values = [ nome, 'Encoste', 'Prev', enc.hPrev || '-', enc.vPrev || '-', resEncH.text, resEncV.text, mediaEnc.toFixed(0) + '%', cargaTexto ];
            sheet.getRow(linhaAtual + 1).values = [ null, null, 'Real', enc.hReal || '-', enc.vReal || '-', null, null, null, null ];
            sheet.getRow(linhaAtual + 2).values = [ null, 'Retirada', 'Prev', ret.hPrev || '-', ret.vPrev || '-', resRetH.text, resRetV.text, mediaRet.toFixed(0) + '%', null ];
            sheet.getRow(linhaAtual + 3).values = [ null, null, 'Real', ret.hReal || '-', ret.vReal || '-', null, null, null, null ];

            // 2. APLICAR CORES DE FONTE (Vermelho se nota < 100)
            if (mediaEnc < 100) sheet.getCell(`H${linhaAtual}`).font = { color: { argb: 'FFFF0000' }, bold: true };
            if (mediaRet < 100) sheet.getCell(`H${linhaAtual + 2}`).font = { color: { argb: 'FFFF0000' }, bold: true };

            const boxBorder = { style: 'thin', color: { argb: 'FF000000' } };
            
            for(let r = 0; r < 4; r++) {
                const rowObj = sheet.getRow(linhaAtual + r);
                rowObj.alignment = { vertical: 'middle', horizontal: 'center' };
                if (r === 1 || r === 3) {
                    rowObj.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
                }

                for(let c = 1; c <= 9; c++) {
                    sheet.getCell(linhaAtual + r, c).border = { top: boxBorder, left: boxBorder, bottom: boxBorder, right: boxBorder };
                }
            }

            sheet.mergeCells(`A${linhaAtual}:A${linhaAtual + 3}`);
            sheet.mergeCells(`B${linhaAtual}:B${linhaAtual + 1}`);
            sheet.mergeCells(`B${linhaAtual + 2}:B${linhaAtual + 3}`);
            sheet.mergeCells(`I${linhaAtual}:I${linhaAtual + 3}`);

            sheet.mergeCells(`F${linhaAtual}:F${linhaAtual + 1}`);
            sheet.mergeCells(`G${linhaAtual}:G${linhaAtual + 1}`);
            sheet.mergeCells(`H${linhaAtual}:H${linhaAtual + 1}`);

            sheet.mergeCells(`F${linhaAtual + 2}:F${linhaAtual + 3}`);
            sheet.mergeCells(`G${linhaAtual + 2}:G${linhaAtual + 3}`);
            sheet.mergeCells(`H${linhaAtual + 2}:H${linhaAtual + 3}`);

            linhaAtual += 4;
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
        const resultados = [];

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

            sheet.eachRow((row) => {
                const val = row.getCell(8).value;
                if (typeof val === 'string' && val.includes('%')) {
                    const numero = parseFloat(val.replace('%', ''));
                    if (!isNaN(numero)) {
                        somaAderencia += numero;
                        qtdValores++;
                    }
                }
            });

            const mediaFinal = qtdValores > 0 ? (somaAderencia / qtdValores) : 0;
            resultados.push({
                turno: turnoNome,
                media: parseFloat(mediaFinal.toFixed(1))
            });
        }

        resultados.sort((a, b) => a.turno.localeCompare(b.turno));

        res.json(resultados);
    } catch (error) {
        console.error("Erro ao ler planilhas:", error);
        res.status(500).json({ erro: "Falha ao processar os arquivos." });
    }
});

app.listen(port, () => {
    console.log(`Servidor rodando! Acesse: http://localhost:${port}`);
    exec(`start http://localhost:${port}`);
});