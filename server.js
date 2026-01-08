const express = require('express');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');

const app = express();
const port = 3000;

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(express.static('public')); 

const terminais = [ "Yara", "Bianchini", "Cotribá", "Ceifagro", "Agrofel", "Pradozem", "Três Tentos", "Recebimento trem", "Formação Trem" ];

app.get('/api/terminais', (req, res) => res.json(terminais));

// --- FUNÇÕES DE CÁLCULO ---
function calcHora(prev, real) {
    if (!prev || !real) return { val: 0, text: '-' };
    const p = new Date(`1970-01-01T${prev}:00`);
    const r = new Date(`1970-01-01T${real}:00`);
    const diff = (r - p) / 36e5;
    if (diff >= -4 && diff <= 1) return { val: 100, text: '100%' };
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

// --- ROTA API PARA O GRÁFICO ---
app.post('/api/calcular', (req, res) => {
    const dados = req.body;
    const resultados = [];
    const nomesTerminais = [...new Set(dados.map(d => d.terminal))];

    nomesTerminais.forEach(nome => {
        const enc = dados.find(d => d.terminal === nome && d.atividade === 'Encoste') || {};
        const ret = dados.find(d => d.terminal === nome && d.atividade === 'Retirada') || {};

        const resEncH = calcHora(enc.hPrev, enc.hReal);
        const resEncV = calcVags(enc.vPrev, enc.vReal);
        const mediaEnc = (resEncH.val + resEncV.val) / 2;

        const resRetH = calcHora(ret.hPrev, ret.hReal);
        const resRetV = calcVags(ret.vPrev, ret.vReal);
        const mediaRet = (resRetH.val + resRetV.val) / 2;

        let mediaGeral = 0;
        if(enc.terminal && ret.terminal) mediaGeral = (mediaEnc + mediaRet) / 2;
        else if(enc.terminal) mediaGeral = mediaEnc;
        else if(ret.terminal) mediaGeral = mediaRet;

        resultados.push({
            terminal: nome,
            mediaEnc: parseFloat(mediaEnc.toFixed(1)),
            mediaRet: parseFloat(mediaRet.toFixed(1)),
            mediaGeral: parseFloat(mediaGeral.toFixed(1))
        });
    });

    res.json(resultados);
});

// --- ROTA DE GERAÇÃO DA PLANILHA ---
app.post('/gerar', async (req, res) => {
    try {
        const dados = req.body;
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Grade de Manobra');

        sheet.columns = [
            { header: 'Terminal', key: 'A', width: 20 },
            { header: 'Atividade', key: 'B', width: 15 },
            { header: 'Tipo', key: 'C', width: 10 },
            { header: 'Horário', key: 'D', width: 12 },
            { header: 'Vagões', key: 'E', width: 10 },
            { header: 'Aderência Hr', key: 'F', width: 15 },
            { header: 'Aderência Vgs', key: 'G', width: 15 },
            { header: '% Geral', key: 'H', width: 12 },
            { header: 'Carga', key: 'I', width: 30 }
        ];

        const headerRow = sheet.getRow(1);
        headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
        headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
        headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
        headerRow.height = 25;

        const nomesTerminais = [...new Set(dados.map(d => d.terminal))];
        let linhaAtual = 2; 

        nomesTerminais.forEach(nome => {
            const enc = dados.find(d => d.terminal === nome && d.atividade === 'Encoste') || {};
            const ret = dados.find(d => d.terminal === nome && d.atividade === 'Retirada') || {};

            const resEncH = calcHora(enc.hPrev, enc.hReal);
            const resEncV = calcVags(enc.vPrev, enc.vReal);
            const mediaEnc = (resEncH.val + resEncV.val) / 2;

            const resRetH = calcHora(ret.hPrev, ret.hReal);
            const resRetV = calcVags(ret.vPrev, ret.vReal);
            const mediaRet = (resRetH.val + resRetV.val) / 2;

            // Define qual carga usar (Encoste ou Retirada, o que tiver preenchido)
            const cargaTexto = enc.carga || ret.carga || '-';

            // Linha 1
            sheet.getRow(linhaAtual).values = [
                nome, 'Encoste', 'Prev', enc.hPrev || '-', enc.vPrev || '-', '-', '-', '-', cargaTexto
            ];
            // Linha 2
            sheet.getRow(linhaAtual + 1).values = [
                null, null, 'Real', enc.hReal || '-', enc.vReal || '-', 
                resEncH.text, resEncV.text, mediaEnc.toFixed(0) + '%', null
            ];
            // Linha 3
            sheet.getRow(linhaAtual + 2).values = [
                null, 'Retirada', 'Prev', ret.hPrev || '-', ret.vPrev || '-', '-', '-', '-', null
            ];
            // Linha 4
            sheet.getRow(linhaAtual + 3).values = [
                null, null, 'Real', ret.hReal || '-', ret.vReal || '-', 
                resRetH.text, resRetV.text, mediaRet.toFixed(0) + '%', null
            ];

            // --- MESCLAGENS ---
            sheet.mergeCells(`A${linhaAtual}:A${linhaAtual + 3}`); // Terminal (4 linhas)
            
            sheet.mergeCells(`B${linhaAtual}:B${linhaAtual + 1}`); // Encoste (2 linhas)
            sheet.mergeCells(`B${linhaAtual + 2}:B${linhaAtual + 3}`); // Retirada (2 linhas)

            // AQUI ESTÁ A MUDANÇA DA CARGA (4 linhas agora):
            sheet.mergeCells(`I${linhaAtual}:I${linhaAtual + 3}`); 

            // Estilos
            for (let i = 0; i < 4; i++) {
                sheet.getRow(linhaAtual + i).alignment = { vertical: 'middle', horizontal: 'center' };
            }

            sheet.getRow(linhaAtual + 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
            sheet.getRow(linhaAtual + 3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };

            const boxBorder = { style: 'thin', color: { argb: 'FF000000' } };
            for(let r = linhaAtual; r <= linhaAtual+3; r++) {
                for(let c = 1; c <= 9; c++) {
                    sheet.getCell(r, c).border = { top: boxBorder, left: boxBorder, bottom: boxBorder, right: boxBorder };
                }
            }

            if (mediaEnc < 100) sheet.getCell(`H${linhaAtual + 1}`).font = { color: { argb: 'FFFF0000' }, bold: true };
            if (mediaRet < 100) sheet.getCell(`H${linhaAtual + 3}`).font = { color: { argb: 'FFFF0000' }, bold: true };

            linhaAtual += 4;
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Grade_Manobra_Final.xlsx');
        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        console.error(error);
        res.status(500).send("Erro");
    }
});

app.listen(port, () => console.log(`Servidor rodando na porta ${port}`));