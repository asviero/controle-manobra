const express = require('express');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');

const app = express();
const port = 3000;

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(express.static('public')); 

// Seus Terminais (conforme modelo)
const terminais = [ "Yara", "Bianchini", "Cotribá", "Ceifagro", "Agrofel", "Pradozem", "Três Tentos", "Recebimento trem", "Formação Trem" ];

app.get('/api/terminais', (req, res) => res.json(terminais));

// --- FUNÇÕES DE CÁLCULO (Mantendo suas premissas) ---
function calcHora(prev, real) {
    if (!prev || !real) return { val: 0, text: '-' };
    const p = new Date(`1970-01-01T${prev}:00`);
    const r = new Date(`1970-01-01T${real}:00`);
    const diff = (r - p) / 36e5;
    
    // Regra: até 4h antes (negativo) ou até 1h depois (positivo)
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

// --- ROTA DE GERAÇÃO DA PLANILHA ---
app.post('/gerar', async (req, res) => {
    try {
        const dados = req.body;
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Grade de Manobra');

        // 1. Configurar Colunas (Troquei Justificativa por Carga)
        sheet.columns = [
            { header: 'Terminal', key: 'A', width: 20 },
            { header: 'Atividade', key: 'B', width: 15 },
            { header: 'Tipo', key: 'C', width: 10 },
            { header: 'Horário', key: 'D', width: 12 },
            { header: 'Vagões', key: 'E', width: 10 },
            { header: 'Aderência Hr', key: 'F', width: 15 },
            { header: 'Aderência Vgs', key: 'G', width: 15 },
            { header: '% Geral', key: 'H', width: 12 },
            { header: 'Carga', key: 'I', width: 30 } // Alterado aqui
        ];

        // 2. Estilizar Cabeçalho
        const headerRow = sheet.getRow(1);
        headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
        headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
        headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
        headerRow.height = 25;

        // Identificar nomes únicos de terminais enviados
        const nomesTerminais = [...new Set(dados.map(d => d.terminal))];

        let linhaAtual = 2; 

        nomesTerminais.forEach(nome => {
            // Pegar dados deste terminal
            const enc = dados.find(d => d.terminal === nome && d.atividade === 'Encoste') || {};
            const ret = dados.find(d => d.terminal === nome && d.atividade === 'Retirada') || {};

            // Calcular
            const resEncH = calcHora(enc.hPrev, enc.hReal);
            const resEncV = calcVags(enc.vPrev, enc.vReal);
            const mediaEnc = (resEncH.val + resEncV.val) / 2;

            const resRetH = calcHora(ret.hPrev, ret.hReal);
            const resRetV = calcVags(ret.vPrev, ret.vReal);
            const mediaRet = (resRetH.val + resRetV.val) / 2;

            // --- DESENHAR AS 4 LINHAS ---
            
            // Linha 1: Encoste Prev
            sheet.getRow(linhaAtual).values = [
                nome, 'Encoste', 'Prev', enc.hPrev || '-', enc.vPrev || '-', '-', '-', '-', enc.carga || '-'
            ];

            // Linha 2: Encoste Real
            sheet.getRow(linhaAtual + 1).values = [
                null, null, 'Real', enc.hReal || '-', enc.vReal || '-', 
                resEncH.text, resEncV.text, mediaEnc.toFixed(0) + '%', null
            ];

            // Linha 3: Retirada Prev
            sheet.getRow(linhaAtual + 2).values = [
                null, 'Retirada', 'Prev', ret.hPrev || '-', ret.vPrev || '-', '-', '-', '-', ret.carga || '-'
            ];

            // Linha 4: Retirada Real
            sheet.getRow(linhaAtual + 3).values = [
                null, null, 'Real', ret.hReal || '-', ret.vReal || '-', 
                resRetH.text, resRetV.text, mediaRet.toFixed(0) + '%', null
            ];

            // --- FORMATAÇÃO VISUAL ---

            // Mesclar células (Terminal e Carga para ficar bonito)
            sheet.mergeCells(`A${linhaAtual}:A${linhaAtual + 3}`); // Terminal
            sheet.mergeCells(`B${linhaAtual}:B${linhaAtual + 1}`); // Encoste Label
            sheet.mergeCells(`B${linhaAtual + 2}:B${linhaAtual + 3}`); // Retirada Label
            
            // Mesclar a Carga para ocupar as duas linhas do Encoste e duas da Retirada
            sheet.mergeCells(`I${linhaAtual}:I${linhaAtual + 1}`); 
            sheet.mergeCells(`I${linhaAtual + 2}:I${linhaAtual + 3}`); 

            // Alinhamento
            for (let i = 0; i < 4; i++) {
                sheet.getRow(linhaAtual + i).alignment = { vertical: 'middle', horizontal: 'center' };
            }

            // Cores alternadas
            sheet.getRow(linhaAtual + 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
            sheet.getRow(linhaAtual + 3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };

            // Bordas
            const boxBorder = { style: 'thin', color: { argb: 'FF000000' } };
            for(let r = linhaAtual; r <= linhaAtual+3; r++) {
                for(let c = 1; c <= 9; c++) {
                    sheet.getCell(r, c).border = {
                        top: boxBorder, left: boxBorder, bottom: boxBorder, right: boxBorder
                    };
                }
            }

            // Destaque Vermelho se aderência baixa
            if (mediaEnc < 100) sheet.getCell(`H${linhaAtual + 1}`).font = { color: { argb: 'FFFF0000' }, bold: true };
            if (mediaRet < 100) sheet.getCell(`H${linhaAtual + 3}`).font = { color: { argb: 'FFFF0000' }, bold: true };

            linhaAtual += 4;
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Grade_Manobra_Final.xlsx');

        await workbook.xlsx.write(res);
        res.end();

    } catch (error) {
        console.error("Erro no Excel:", error);
        res.status(500).send("Erro ao gerar planilha.");
    }
});

app.listen(port, () => console.log(`Servidor rodando na porta ${port}`));