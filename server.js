const express = require('express');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');

const { exec } = require('child_process');

const app = express();
const port = 3000;

const upload = multer({ storage: multer.memoryStorage() });

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json({ limit: '5mb' }));
app.use(express.static('public'));
app.use(express.static(path.join(__dirname, 'public')));

const TERMINAIS = ['Yara', 'Bianchini', 'Cotribá', 'Ceifrago', 'Agrofel', 'Pradozem', 'Três Tentos', 'Recebimento Trem', 'Formação Trem'];
const MOTIVOS = [
    'Absenteísmo Rumo', 'Atraso de Trem', 'Avaria Balança Terminal', 'Avaria de Locomotiva',
    'Avaria Mecânica Terminal', 'Avaria Trator Terminal', 'Chuva', 'Falta de Energia',
    'Falta de Equipe Terminal', 'Manutenção Terminal', 'Manutenção Via Permanente',
    'Prioridade Embarque Rodoviário', 'Prioridade Operacional Rumo',
    'Prioridade para Descarga de Caminhões', 'Sem Atraso'
];

app.get('/api/terminais', (req, res) => res.json(TERMINAIS));

// ─────────────────────────────────────────────────────────────────────────
// PERSISTÊNCIA DO HISTÓRICO (arquivos JSON, um por registro)
// ─────────────────────────────────────────────────────────────────────────
const baseDir = process.pkg ? path.dirname(process.execPath) : __dirname;
const historicoDir = path.join(baseDir, 'historico');
if (!fs.existsSync(historicoDir)) fs.mkdirSync(historicoDir, { recursive: true });

function novoId() {
    return Date.now().toString(36) + '-' + crypto.randomBytes(3).toString('hex');
}

function obterTurnoDaHora(hp, hr) {
    const horaStr = hr || hp;
    if (!horaStr) return 'Geral';
    const h = parseInt(horaStr.split(':')[0], 10);
    if (h >= 7 && h < 15) return 'A';
    if (h >= 15 && h < 23) return 'B';
    return 'C';
}

function salvarRegistro({ turno, data, terminais }) {
    const turnosDetectados = ['A', 'B', 'C'];
    let primeiroRegistro = null;
    const dataLimpa = String(data || '').replace(/^Data:\s*/i, '').trim();

    turnosDetectados.forEach(tLabel => {
        const terminaisTurno = {};
        let temOperacao = false;

        TERMINAIS.forEach(term => {
            if (!terminais[term] || !terminais[term].manobras) return;
            const manobrasFiltradas = terminais[term].manobras.filter(m => obterTurnoDaHora(m.hp, m.hr) === tLabel);
            
            if (manobrasFiltradas.length > 0) {
                temOperacao = true;
                terminaisTurno[term] = {
                    tipo: terminais[term].tipo || '',
                    produto: terminais[term].produto || '',
                    manobras: manobrasFiltradas
                };
            }
        });

        if (temOperacao) {
            // Procura se já existe um registro salvo para esta mesma data e turno
            const registrosExistentes = listarRegistros();
            let registroEncontrado = registrosExistentes.find(r => r.turno === tLabel && r.data_turno === dataLimpa);

            let id;
            if (registroEncontrado) {
                // Reaproveita o ID existente para atualizar em vez de duplicar
                id = registroEncontrado.id;
            } else {
                // Cria um novo ID se não existir
                id = novoId();
            }

            const registro = {
                id,
                turno: tLabel,
                data_turno: dataLimpa,
                dados_json: terminaisTurno,
                criadoEm: new Date().toISOString()
            };

            fs.writeFileSync(path.join(historicoDir, `${id}.json`), JSON.stringify(registro), 'utf8');
            if (!primeiroRegistro) primeiroRegistro = registro;
        }
    });

    return primeiroRegistro || { id: novoId(), turno: 'A', data_turno: '', dados_json: terminais };
}

function listarRegistros() {
    return fs.readdirSync(historicoDir)
        .filter(f => f.endsWith('.json'))
        .map(f => {
            try { return JSON.parse(fs.readFileSync(path.join(historicoDir, f), 'utf8')); }
            catch { return null; }
        })
        .filter(Boolean);
}

function lerRegistro(id) {
    const p = path.join(historicoDir, `${id}.json`);
    if (!fs.existsSync(p)) return null;
    return JSON.parse(fs.readFileSync(p, 'utf8'));
}

// Converte 'dd/mm/yyyy' em Date; retorna null se inválido
function parseDataBR(str) {
    if (!str) return null;
    const [d, m, y] = String(str).split('/').map(Number);
    if (!d || !m || !y) return null;
    return new Date(y, m - 1, d);
}

// ─────────────────────────────────────────────────────────────────────────
// CÁLCULO DE ADERÊNCIA
// ─────────────────────────────────────────────────────────────────────────
function toMin(s) {
    if (!s) return null;
    const [h, m] = String(s).split(':').map(Number);
    if (Number.isNaN(h) || Number.isNaN(m)) return null;
    return h * 60 + m;
}
function calcAH(hp, hr) {
    const p = toMin(hp), r = toMin(hr);
    if (p === null || r === null) return null;
    return Math.abs(p - r) <= 60 ? 100 : 0;
}
function calcAV(vp, vr) {
    const p = parseInt(vp) || 0;
    if (!p) return null;
    return Math.min(100, Math.round(((parseInt(vr) || 0) / p) * 100));
}

function calcularAgregadoRegistro(dados) {
    let somaAg = 0, countAg = 0, totalVgR = 0;
    Object.values(dados || {}).forEach(d => {
        (d.manobras || []).forEach(m => {
            const ah = calcAH(m.hp, m.hr);
            const av = calcAV(m.vp, m.vr);
            const ag = (ah !== null && av !== null) ? (ah + av) / 2 : (ah ?? av ?? null);
            if (ag !== null) { somaAg += ag; countAg++; }
            totalVgR += parseInt(m.vr) || 0;
        });
    });
    return {
        adhMedia: countAg > 0 ? somaAg / countAg : null,
        totalVagoes: totalVgR
    };
}

// ─────────────────────────────────────────────────────────────────────────
// PLANILHA — Ordem: Geral, depois Turnos A, B e C
// ─────────────────────────────────────────────────────────────────────────
function montarWorkbook({ terminais, turno, data, isVazia = false }) {
    const workbook = new ExcelJS.Workbook();
    // A aba "Geral" fica em primeiro, seguida pelas abas de cada turno
    const abas = ['Geral', 'A', 'B', 'C'];

    abas.forEach(tLabel => {
        const sheet = workbook.addWorksheet(tLabel === 'Geral' ? 'Geral' : `Turno ${tLabel}`);
        
        sheet.columns = [
            { key: 'A', width: 20 }, { key: 'B', width: 24 }, { key: 'C', width: 18 },
            { key: 'D', width: 14 }, { key: 'E', width: 13 }, { key: 'F', width: 12 },
            { key: 'G', width: 13 }, { key: 'H', width: 12 }, { key: 'I', width: 15 },
            { key: 'J', width: 15 }, { key: 'K', width: 14 }, { key: 'L', width: 28 }
        ];

        sheet.getCell('A1').value = tLabel === 'Geral' 
            ? 'RELATÓRIO OPERACIONAL - VISÃO GERAL (DIÁRIO)' 
            : `RELATÓRIO OPERACIONAL - TURNO: ${tLabel}`;
            
        sheet.mergeCells('A1:L1');
        sheet.getCell('A1').font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 14 };
        sheet.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00345E' } };
        sheet.getCell('A1').alignment = { vertical: 'middle', horizontal: 'center' };
        sheet.getRow(1).height = 28;

        const dataLimpa = String(data || '').replace(/^Data:\s*/i, '').trim();
        sheet.getCell('A2').value = `Data: ${dataLimpa}`;
        sheet.mergeCells('A2:L2');
        sheet.getCell('A2').font = { bold: true, size: 11, color: { argb: 'FF00345E' } };
        sheet.getCell('A2').alignment = { vertical: 'middle', horizontal: 'left' };
        sheet.getRow(2).height = 20;

        const headers = ['Terminal', 'Tipo Operação', 'Produto', 'Manobra', 'Horário Previsto', 'Vagões Previstos',
            'Horário Realizado', 'Vagões Realizados', 'Aderência Horário', 'Aderência Vagões', 'Aderência Geral', 'Motivo de Atraso'];
        const headerRow = sheet.getRow(3);
        headers.forEach((h, i) => headerRow.getCell(i + 1).value = h);
        headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
        headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
        headerRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        headerRow.height = 30;

        let linha = 4;
        const boxBorder = { style: 'thin', color: { argb: 'FF000000' } };
        let temDadoNaAba = false;

        TERMINAIS.forEach(nome => {
            const d = terminais[nome];
            if (!d) return;
            
            // Se for a aba 'Geral' ou planilha vazia, traz tudo. Se for A, B ou C, filtra pela hora.
            const manobras = (tLabel === 'Geral' || isVazia) 
                ? (d.manobras || []) 
                : (d.manobras || []).filter(m => obterTurnoDaHora(m.hp, m.hr) === tLabel);

            if (!manobras.length) return;
            temDadoNaAba = true;

            const linhaInicial = linha;

            manobras.forEach(m => {
                const ah = calcAH(m.hp, m.hr);
                const av = calcAV(m.vp, m.vr);
                const ag = (ah !== null && av !== null) ? (ah + av) / 2 : (ah ?? av ?? null);
                
                let labelTipo = 'Encoste';
                if(m.tipo === 'retirada') labelTipo = 'Retirada';
                else if(m.tipo === 'chegada') labelTipo = 'Chegada';
                else if(m.tipo === 'expedicao') labelTipo = 'Expedição';

                const row = sheet.getRow(linha);
                row.values = {
                    A: nome, B: d.tipo || '', C: d.produto || '',
                    D: `${labelTipo} ${m.index}`,
                    E: m.hp || '', F: m.vp || 0,
                    G: m.hr || '', H: m.vr || 0,
                    I: ah !== null ? ah + '%' : '-',
                    J: av !== null ? av + '%' : '-',
                    K: ag !== null ? ag.toFixed(1) + '%' : '-',
                    L: m.motivo || 'Sem Atraso'
                };
                row.alignment = { vertical: 'middle', horizontal: 'center' };
                if (ag !== null && ag < 100) row.getCell('K').font = { color: { argb: 'FFFF0000' }, bold: true };
                for (let c = 1; c <= 12; c++) sheet.getCell(linha, c).border = { top: boxBorder, left: boxBorder, bottom: boxBorder, right: boxBorder };
                linha++;
            });

            if (linha - 1 > linhaInicial) {
                ['A', 'B', 'C'].forEach(col => sheet.mergeCells(`${col}${linhaInicial}:${col}${linha - 1}`));
            }
        });

        if (!temDadoNaAba) {
            sheet.getCell('A4').value = "Nenhuma manobra registrada.";
            sheet.mergeCells('A4:L4');
            sheet.getCell('A4').alignment = { horizontal: 'center', vertical: 'middle' };
            sheet.getCell('A4').font = { italic: true, color: { argb: 'FF888888' } };
        }
    });

    return workbook;
}

function parseWorkbook(workbook) {
    let dataGeral = '';
    const terminaisGlobais = {};

    workbook.worksheets.forEach(sheet => {
        const dataCell = sheet.getCell('A2').value;
        if (dataCell && typeof dataCell === 'string' && dataCell.includes('Data:')) {
            dataGeral = dataCell.split('Data:')[1].trim();
        }

        sheet.eachRow((row, rowNumber) => {
            if (rowNumber <= 3) return; 
            const terminal = row.getCell(1).value;
            if (!terminal || typeof terminal !== 'string' || terminal.includes('Nenhuma manobra')) return;

            const manobraLabel = row.getCell(4).value;
            if (!manobraLabel) return;
            
            const match = /^(Encoste|Retirada|Chegada|Expedi[çc][ãa]o)\s+(\d+)$/i.exec(String(manobraLabel).trim());
            if (!match) return;

            if (!terminaisGlobais[terminal]) {
                terminaisGlobais[terminal] = {
                    tipo: row.getCell(2).value || '',
                    produto: row.getCell(3).value || '',
                    manobras: []
                };
            }
            const val = c => { const v = row.getCell(c).value; return (v === null || v === undefined) ? '' : v; };
            
            let tipoInterno = match[1].toLowerCase().replace('ç','c').replace('ã','a');
            
            terminaisGlobais[terminal].manobras.push({
                tipo: tipoInterno,
                index: parseInt(match[2], 10),
                hp: val(5) || '',
                vp: parseInt(val(6)) || 0,
                hr: val(7) || '',
                vr: parseInt(val(8)) || 0,
                motivo: val(12) || 'Sem Atraso'
            });
        });
    });

    return { turno: 'Diário', data: dataGeral, terminais: terminaisGlobais };
}

// ─────────────────────────────────────────────────────────────────────────
// ROTAS
// ─────────────────────────────────────────────────────────────────────────
app.post('/gerar', async (req, res) => {
    try {
        const { terminais, turno, data } = req.body;
        const workbook = montarWorkbook({ terminais, turno, data });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=Relatorio_Operacional.xlsx`);
        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        console.error(error);
        res.status(500).send('Erro ao gerar planilha.');
    }
});

app.get('/api/planilha-vazia', async (req, res) => {
    try {
        const terminaisVazios = {};
        TERMINAIS.forEach(t => {
            let manobras = [];
            if (t === 'Recebimento Trem') {
                manobras = [{ tipo: 'chegada', index: 1, hp: '', vp: '', hr: '', vr: '', motivo: 'Sem Atraso' }];
            } else if (t === 'Formação Trem') {
                manobras = [{ tipo: 'expedicao', index: 1, hp: '', vp: '', hr: '', vr: '', motivo: 'Sem Atraso' }];
            } else {
                manobras = [
                    { tipo: 'encoste', index: 1, hp: '', vp: '', hr: '', vr: '', motivo: 'Sem Atraso' },
                    { tipo: 'retirada', index: 1, hp: '', vp: '', hr: '', vr: '', motivo: 'Sem Atraso' }
                ];
            }
            terminaisVazios[t] = { tipo: '', produto: '', manobras };
        });
        const workbook = montarWorkbook({ terminais: terminaisVazios, turno: 'Diário', data: '', isVazia: true });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Planilha_Campo_Vazia.xlsx');
        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        console.error(error);
        res.status(500).send('Erro ao gerar planilha vazia.');
    }
});

app.post('/api/importar', upload.single('arquivo'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ erro: 'Nenhum arquivo enviado.' });
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const { turno, terminais } = parseWorkbook(workbook);
        if (!Object.keys(terminais).length) throw new Error('Planilha sem dados reconhecíveis.');
        res.json({ turno, terminais });
    } catch (error) {
        console.error(error);
        res.status(400).json({ erro: 'Não foi possível importar. Verifique se é um relatório gerado por este sistema.' });
    }
});

app.post('/api/salvar-historico-json', (req, res) => {
    try {
        const { terminais, turno, data } = req.body;
        const registro = salvarRegistro({ turno, data, terminais });
        res.json({ ok: true, id: registro ? registro.id : null });
    } catch (error) {
        console.error(error);
        res.status(500).json({ erro: 'Erro ao salvar histórico.' });
    }
});

app.post('/api/historico/salvar', upload.array('arquivos', 10), async (req, res) => {
    try {
        let salvos = 0;
        for (const file of req.files || []) {
            try {
                const workbook = new ExcelJS.Workbook();
                await workbook.xlsx.load(file.buffer);
                const { turno, data, terminais } = parseWorkbook(workbook);
                if (Object.keys(terminais).length) {
                    salvarRegistro({ turno, data, terminais });
                    salvos++;
                }
            } catch (e) { }
        }
        res.json({ salvos });
    } catch (error) {
        console.error(error);
        res.status(500).json({ erro: 'Erro ao salvar planilhas no histórico.' });
    }
});

app.get('/api/historico', (req, res) => {
    try {
        const { dataIni, dataFim, turno } = req.query;
        const dIni = parseDataBR(dataIni);
        const dFim = parseDataBR(dataFim);

        let registros = listarRegistros().filter(r => {
            if (turno && r.turno !== turno) return false;
            const d = parseDataBR(r.data_turno);
            if (dIni && d && d < dIni) return false;
            if (dFim && d && d > dFim) return false;
            return true;
        });

        const resumidos = registros.map(r => {
            const agg = calcularAgregadoRegistro(r.dados_json);
            return { id: r.id, turno: r.turno, data_turno: r.data_turno, adh_media: agg.adhMedia };
        }).sort((a, b) => (parseDataBR(a.data_turno) || 0) - (parseDataBR(b.data_turno) || 0));

        const comAdh = resumidos.filter(r => r.adh_media !== null);
        const adhMedia = comAdh.length ? comAdh.reduce((s, r) => s + r.adh_media, 0) / comAdh.length : null;
        const totalVagoes = registros.reduce((s, r) => s + calcularAgregadoRegistro(r.dados_json).totalVagoes, 0);

        res.json({ registros: resumidos, adhMedia, totalVagoes });
    } catch (error) {
        console.error(error);
        res.status(500).json({ erro: 'Erro ao buscar histórico.' });
    }
});

app.get('/api/historico/:id', (req, res) => {
    const registro = lerRegistro(req.params.id);
    if (!registro) return res.status(404).json({ erro: 'Registro não encontrado.' });
    res.json(registro);
});

app.post('/api/comparar', upload.array('planilhas', 5), async (req, res) => {
    try {
        const resultadosTurnos = [];
        const consolidadoMotivos = {};

        for (const file of req.files) {
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(file.buffer);
            
            // Pega a primeira aba (Geral) para consolidar
            const sheet = workbook.worksheets[0];

            let somaAderencia = 0;
            let qtdValores = 0;
            let turnoNome = 'Geral';

            sheet.eachRow((row, rowNumber) => {
                if (rowNumber <= 3) return;

                const valAderencia = row.getCell(11).value;
                if (typeof valAderencia === 'string' && valAderencia.includes('%')) {
                    const numero = parseFloat(valAderencia.replace('%', ''));
                    if (!isNaN(numero)) {
                        somaAderencia += numero;
                        qtdValores++;
                    }
                }

                const valMotivo = row.getCell(12).value;
                if (valMotivo && typeof valMotivo === 'string' && valMotivo !== '' && valMotivo !== '-' && valMotivo !== 'Sem Atraso') {
                    consolidadoMotivos[valMotivo] = (consolidadoMotivos[valMotivo] || 0) + 1;
                }
            });

            const mediaFinal = qtdValores > 0 ? (somaAderencia / qtdValores) : 0;
            resultadosTurnos.push({
                turno: turnoNome,
                media: parseFloat(mediaFinal.toFixed(1))
            });
        }

        res.json({
            turnos: resultadosTurnos,
            motivosConsolidados: consolidadoMotivos
        });
    } catch (error) {
        console.error('Erro ao ler planilhas:', error);
        res.status(500).json({ erro: 'Falha ao processar os arquivos.' });
    }
});

// Altera a inicialização para usar a porta 0 (porta livre automática) ou uma porta padrão se preferir
const server = app.listen(0, () => {
    const portaAtual = server.address().port;
    console.log(`Servidor rodando! Acesse: http://localhost:${portaAtual}`);
    try { 
        exec(`start http://localhost:${portaAtual}`); 
    } catch (e) { 
        /* ambiente sem 'start' (ex: Linux/macOS) */ 
    }
});