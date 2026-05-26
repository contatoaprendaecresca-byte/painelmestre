/**
 * ClaudeConnect.gs
 * Integração com Claude AI — Análise automática do MESTRE
 *
 * ✅ Não altera nenhum arquivo existente (motor, insights, menu, etc.)
 * ✅ Adiciona entrada no menu via onOpen() existente — veja instruções no final
 * ✅ Lê os dados do MESTRE, chama a API do Claude e salva o resultado
 *
 * SETUP (uma vez só):
 *  1. Abra a aba PROPRIEDADES da planilha
 *  2. Adicione uma linha: | CLAUDE_API_KEY | sua-chave-aqui |
 *  3. Clique no menu → Mestre → Sync PROPRIEDADES → Script Properties
 *  4. Pronto. Rode "🤖 Claude: Analisar últimos 30 dias" para testar.
 *
 * COMO OBTER SUA CHAVE:
 *  Acesse: https://console.anthropic.com/settings/keys
 *  Crie uma chave (API Key) e cole na aba PROPRIEDADES
 */
 
// ─── CONFIGURAÇÕES ───────────────────────────────────────────────────────────
 
const CLAUDE_CFG = {
  MODEL:           'claude-sonnet-4-20250514',
  MAX_TOKENS:      1500,
  OUTPUT_SHEET:    'CLAUDE_INSIGHTS',
  SOURCE_SHEET:    'MESTRE',
  PROP_KEY:        'CLAUDE_API_KEY',
  DEFAULT_DAYS:    30,        // dias lidos por padrão
  MAX_ROWS_TO_API: 60,        // máximo de linhas enviadas pro Claude (evita payload enorme)
};
 
// ─── FUNÇÕES PÚBLICAS (aparecem no menu) ─────────────────────────────────────
 
/** Analisa os últimos 30 dias automaticamente */
function claudeAnalisar30d() {
  const hoje = new Date();
  const inicio = claude_addDays_(hoje, -(CLAUDE_CFG.DEFAULT_DAYS - 1));
  claudeGerarAnalise_(inicio, hoje, 'automático (30d)');
}
 
/** Analisa os últimos 7 dias automaticamente */
function claudeAnalisar7d() {
  const hoje = new Date();
  const inicio = claude_addDays_(hoje, -6);
  claudeGerarAnalise_(inicio, hoje, 'automático (7d)');
}
 
/** Permite escolher o período manualmente */
function claudeAnalisarPeriodoManual() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    '📆 Período para análise',
    'Digite no formato: DD/MM/AAAA a DD/MM/AAAA\nExemplo: 01/02/2026 a 28/02/2026',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
 
  const partes = (resp.getResponseText() || '').trim().split(/\s+a\s+/i);
  if (partes.length !== 2) {
    ui.alert('❌ Formato inválido. Use: DD/MM/AAAA a DD/MM/AAAA');
    return;
  }
 
  const ini = claude_parseBRDate_(partes[0].trim());
  const fim = claude_parseBRDate_(partes[1].trim());
  if (!ini || !fim) {
    ui.alert('❌ Datas inválidas. Confira dia/mês/ano.');
    return;
  }
 
  claudeGerarAnalise_(ini, fim, 'manual');
}
 
/** Instala gatilho semanal: toda segunda 07:15 (após o motor e o insight rodarem) */
function claudeInstalarGatilhoSemanal() {
  claudeRemoverGatilhoSemanal();
  ScriptApp.newTrigger('claudeAnalisar30d')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .nearMinute(15)
    .create();
  SpreadsheetApp.getUi().alert('⏰ Gatilho semanal instalado: segunda-feira 07:15\n(roda depois do Motor 06:00 e do Insight 06:45)');
}
 
/** Remove o gatilho semanal do Claude */
function claudeRemoverGatilhoSemanal() {
  const handlers = ['claudeAnalisar30d', 'claudeAnalisar7d'];
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (handlers.includes(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  if (removed > 0) {
    SpreadsheetApp.getUi().alert(`🧹 ${removed} gatilho(s) do Claude removido(s).`);
  }
}
 
// ─── CORE ────────────────────────────────────────────────────────────────────
 
function claudeGerarAnalise_(dataInicio, dataFim, modo) {
  // 1. Chave da API
  const apiKey = PropertiesService.getScriptProperties().getProperty(CLAUDE_CFG.PROP_KEY);
  if (!apiKey) {
    SpreadsheetApp.getUi().alert(
      '❌ Chave do Claude não encontrada.\n\n' +
      'Adicione na aba PROPRIEDADES:\n' +
      'Coluna A: CLAUDE_API_KEY\n' +
      'Coluna B: sua-chave\n\n' +
      'Depois clique em: Mestre → Sync PROPRIEDADES → Script Properties'
    );
    return;
  }
 
  // 2. Lê dados do MESTRE
  const sh = SpreadsheetApp.getActive().getSheetByName(CLAUDE_CFG.SOURCE_SHEET);
  if (!sh) throw new Error('Aba MESTRE não encontrada.');
 
  const dadosFiltrados = claude_lerMestre_(sh, dataInicio, dataFim);
  if (dadosFiltrados.length === 0) {
    SpreadsheetApp.getUi().alert('❌ Nenhum dado encontrado no período selecionado.');
    return;
  }
 
  // 3. Monta o prompt
  const prompt = claude_montarPrompt_(dadosFiltrados, dataInicio, dataFim);
 
  // 4. Chama a API do Claude
  const analise = claude_chamarAPI_(apiKey, prompt);
 
  // 5. Salva na aba CLAUDE_INSIGHTS
  claude_salvarSaida_(analise, dadosFiltrados, dataInicio, dataFim, modo);
 
  // 6. Mostra o resultado em um dialog
  claude_mostrarDialogo_(analise);
}
 
// ─── LEITURA DO MESTRE ───────────────────────────────────────────────────────
 
function claude_lerMestre_(sh, dataInicio, dataFim) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
 
  const values = sh.getRange(1, 1, lastRow, sh.getLastColumn()).getValues();
  const headers = values[0].map(h => String(h || '').trim());
  const headerMap = {};
  headers.forEach((h, i) => { if (h) headerMap[h] = i; });
 
  // Índices das colunas que interessam
  const idx = {
    data:        claude_colIdx_(headerMap, ['Data']),
    conta:       claude_colIdx_(headerMap, ['Conta ID']),
    fb:          claude_colIdx_(headerMap, ['Facebook (R$)']),
    total:       claude_colIdx_(headerMap, ['Total (R$)']),
    sessoes:     claude_colIdx_(headerMap, ['Sessões GA4']),
    transacoes:  claude_colIdx_(headerMap, ['Transações GA4']),
    recCapt:     claude_colIdx_(headerMap, ['Receita Captada GA4 (R$)']),
    recPaga:     claude_colIdx_(headerMap, ['Receita Paga Bagy (R$)', 'Receita Paga LojaVirtual (R$)']),
    pedidos:     claude_colIdx_(headerMap, ['Pedidos Pagos Bagy', 'Pedidos Pagos LojaVirtual']),
    naoPagos:    claude_colIdx_(headerMap, ['Não Pagos Bagy', 'Não Pagos LojaVirtual (R$)']),
    ticket:      claude_colIdx_(headerMap, ['Ticket Médio (R$)']),
    bounce:      claude_colIdx_(headerMap, ['Taxa De Rejeição GA4']),
    roas:        claude_colIdx_(headerMap, ['ROAS/Diario [CAP]']),
  };
 
  const iniTime = claude_startOfDay_(dataInicio).getTime();
  const fimTime = claude_startOfDay_(dataFim).getTime();
 
  const linhas = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const d = claude_toDate_(row[idx.data]);
    if (!d) continue;
    const dt = claude_startOfDay_(d).getTime();
    if (dt < iniTime || dt > fimTime) continue;
 
    linhas.push({
      data:       claude_fmtDDMM_(d),
      conta:      idx.conta >= 0 ? String(row[idx.conta] || '') : '',
      invest:     claude_toNum_(idx.total >= 0 ? row[idx.total] : (idx.fb >= 0 ? row[idx.fb] : 0)),
      sessoes:    claude_toNum_(idx.sessoes >= 0 ? row[idx.sessoes] : 0),
      transacoes: claude_toNum_(idx.transacoes >= 0 ? row[idx.transacoes] : 0),
      recCaptada: claude_toNum_(idx.recCapt >= 0 ? row[idx.recCapt] : 0),
      recPaga:    claude_toNum_(idx.recPaga >= 0 ? row[idx.recPaga] : 0),
      pedidos:    claude_toNum_(idx.pedidos >= 0 ? row[idx.pedidos] : 0),
      naoPagos:   claude_toNum_(idx.naoPagos >= 0 ? row[idx.naoPagos] : 0),
      ticket:     claude_toNum_(idx.ticket >= 0 ? row[idx.ticket] : 0),
      bounce:     claude_toNum_(idx.bounce >= 0 ? row[idx.bounce] : 0),
      roas:       claude_toNum_(idx.roas >= 0 ? row[idx.roas] : 0),
    });
  }
 
  // Limita o número de linhas enviadas pra não exceder o payload
  if (linhas.length > CLAUDE_CFG.MAX_ROWS_TO_API) {
    return linhas.slice(-CLAUDE_CFG.MAX_ROWS_TO_API);
  }
  return linhas;
}
 
// ─── PROMPT ──────────────────────────────────────────────────────────────────
 
function claude_montarPrompt_(dados, ini, fim) {
  // Totais do período
  const totalInvest  = dados.reduce((s, r) => s + r.invest, 0);
  const totalSessoes = dados.reduce((s, r) => s + r.sessoes, 0);
  const totalPedidos = dados.reduce((s, r) => s + r.pedidos, 0);
  const totalRecPaga = dados.reduce((s, r) => s + r.recPaga, 0);
  const totalNaoPag  = dados.reduce((s, r) => s + r.naoPagos, 0);
  const roasGeral    = totalInvest > 0 ? (totalRecPaga / totalInvest).toFixed(2) : 'n/a';
 
  // Tabela simplificada (data | invest | sessoes | pedidos | recPaga | roas)
  const linhasTabela = dados.map(r =>
    `${r.data} | R$${r.invest.toFixed(0)} | ${r.sessoes} sessões | ${r.pedidos} pedidos | R$${r.recPaga.toFixed(0)} pago | ROAS ${r.roas.toFixed(1)}x`
  ).join('\n');
 
  const conta = dados[0]?.conta || 'cliente';
 
  return `Você é um especialista em e-commerce e marketing digital brasileiro. Analise os dados abaixo de uma loja de moda feminina (${conta}) e gere um diagnóstico direto e útil.
 
PERÍODO: ${claude_fmtDDMM_(ini)} a ${claude_fmtDDMM_(fim)}
 
RESUMO DO PERÍODO:
- Investimento total: R$${totalInvest.toFixed(0)}
- Sessões totais: ${totalSessoes}
- Pedidos pagos: ${totalPedidos}
- Receita paga: R$${totalRecPaga.toFixed(0)}
- Pedidos não pagos (abandono): R$${totalNaoPag.toFixed(0)}
- ROAS geral: ${roasGeral}x
 
DADOS DIÁRIOS (data | invest | sessões | pedidos pagos | receita paga | roas):
${linhasTabela}
 
RESPONDA EM PORTUGUÊS BRASILEIRO, de forma clara e direta. Estruture assim:
 
1. DIAGNÓSTICO GERAL (2-3 frases resumindo o período)
2. PONTOS POSITIVOS (máximo 3 bullets)
3. PONTOS DE ATENÇÃO (máximo 3 bullets — focado em onde a loja está perdendo dinheiro)
4. AÇÃO PRIORITÁRIA DA SEMANA (1 ação específica e prática, com base nos dados)
 
Seja objetivo. Não use linguagem técnica desnecessária. O texto será lido pelo dono da loja.`;
}
 
// ─── API DO CLAUDE ────────────────────────────────────────────────────────────
 
function claude_chamarAPI_(apiKey, prompt) {
  const payload = {
    model:      CLAUDE_CFG.MODEL,
    max_tokens: CLAUDE_CFG.MAX_TOKENS,
    messages:   [{ role: 'user', content: prompt }]
  };
 
  const options = {
    method:             'post',
    contentType:        'application/json',
    headers: {
      'x-api-key':         apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  };
 
  const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
  const code = resp.getResponseCode();
 
  if (code !== 200) {
    const body = resp.getContentText();
    Logger.log('Claude API error %s: %s', code, body);
    throw new Error(`Erro na API do Claude (${code}). Verifique sua chave em PROPRIEDADES.`);
  }
 
  const json = JSON.parse(resp.getContentText());
  return (json.content && json.content[0] && json.content[0].text) || '(sem resposta)';
}
 
// ─── SAÍDA ────────────────────────────────────────────────────────────────────
 
function claude_salvarSaida_(analise, dados, ini, fim, modo) {
  const ss   = SpreadsheetApp.getActive();
  const nome = CLAUDE_CFG.OUTPUT_SHEET;
  let sh     = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    sh.getRange('A1').setValue('Análises geradas pelo Claude AI').setFontWeight('bold');
    sh.setColumnWidth(1, 300);
    sh.setColumnWidth(2, 700);
    sh.setFrozenRows(1);
  }
 
  // Cabeçalho da seção
  const nextRow = sh.getLastRow() + 2;
  sh.getRange(nextRow, 1, 1, 4)
    .setValues([[`📆 ${claude_fmtDDMM_(ini)} a ${claude_fmtDDMM_(fim)}`, `Conta: ${dados[0]?.conta || '?'}`, `Modo: ${modo}`, new Date()]])
    .setFontWeight('bold')
    .setBackground('#f0f4ff');
 
  // Texto da análise
  const textRow = nextRow + 1;
  sh.getRange(textRow, 1, 1, 4).merge();
  sh.getRange(textRow, 1)
    .setValue(analise)
    .setWrap(true);
  sh.setRowHeight(textRow, 300);
 
  SpreadsheetApp.flush();
}
 
function claude_mostrarDialogo_(analise) {
  const safe = claude_escapeHtml_(analise);
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 14px;">
      <div style="font-size: 14px; margin-bottom: 10px; font-weight: bold;">🤖 Análise do Claude</div>
      <textarea style="width:100%; height:440px; padding:10px; border:1px solid #ddd;
        border-radius:8px; font-size:13px; line-height:1.5;">${safe}</textarea>
      <div style="margin-top: 10px; font-size: 12px; color: #555;">
        Dica: Ctrl+A → Ctrl+C para copiar tudo. Também salvo na aba <b>${CLAUDE_CFG.OUTPUT_SHEET}</b>.
      </div>
    </div>
  `).setWidth(620).setHeight(560);
  SpreadsheetApp.getUi().showModelessDialog(html, '🤖 Claude AI — Análise');
}
 
// ─── HELPERS ──────────────────────────────────────────────────────────────────
 
function claude_colIdx_(headerMap, possiveisNomes) {
  for (const nome of possiveisNomes) {
    if (Object.prototype.hasOwnProperty.call(headerMap, nome)) return headerMap[nome];
  }
  return -1;
}
 
function claude_toNum_(v) {
  if (v === null || v === '' || v === undefined) return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  let s = String(v).trim().replace(/[R$\s\u00A0]/g, '').replace(/\./g, '').replace(',', '.');
  const n = Number(s);
  return isFinite(n) ? n : 0;
}
 
function claude_toDate_(v) {
  if (v instanceof Date) return v;
  if (typeof v === 'string') {
    const m = v.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  }
  return null;
}
 
function claude_startOfDay_(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}
 
function claude_addDays_(d, n) {
  const x = new Date(d);
  x.setDate(x.getDate() + n);
  return x;
}
 
function claude_fmtDDMM_(d) {
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  return `${dd}/${mm}`;
}
 
function claude_parseBRDate_(s) {
  const m = String(s || '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  return (d.getDate() === Number(m[1])) ? d : null;
}
 
function claude_escapeHtml_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
 
/*
 ╔══════════════════════════════════════════════════════════════╗
 ║  ADICIONAR AO MENU (Menu.gs / Main.txt) — INSTRUÇÕES        ║
 ║                                                              ║
 ║  No arquivo Main.txt, dentro da função onOpen(e),           ║
 ║  ANTES de root.addToUi(), adicione:                         ║
 ║                                                              ║
 ║  const mClaude = ui.createMenu('🤖 Claude AI');             ║
 ║  mClaude                                                     ║
 ║    .addItem('📊 Analisar últimos 7 dias',  'claudeAnalisar7d')  ║
 ║    .addItem('📊 Analisar últimos 30 dias', 'claudeAnalisar30d') ║
 ║    .addItem('📆 Período manual…', 'claudeAnalisarPeriodoManual') ║
 ║    .addSeparator()                                           ║
 ║    .addItem('⏰ Instalar gatilho semanal (seg 07:15)', 'claudeInstalarGatilhoSemanal') ║
 ║    .addItem('🧹 Remover gatilho', 'claudeRemoverGatilhoSemanal'); ║
 ║  root.addSubMenu(mClaude);                                   ║
 ╚══════════════════════════════════════════════════════════════╝
*/
 
