// Certidões são da empresa contratada: enviadas uma vez, valem para todos os contratos dela.
const { chromium } = require('playwright');
const fs = require('fs');
const SP = process.argv[2];
const LIB = fs.readFileSync(SP + '/xlsxpkg/package/dist/xlsx.full.min.js', 'utf8');
const URL = 'file://' + SP + '/gestao-contratos.html';
const erros = [];
const ok = (c, m) => { console.log((c ? 'OK   ' : 'FALHA ') + m); if (!c) erros.push(m); };
(async () => {
  const b = await chromium.launch();
  const ctx = await b.newContext({ viewport: { width: 1440, height: 900 } });
  await ctx.route(/jsdelivr/, r => r.fulfill({ body: LIB, contentType: 'application/javascript' }));
  await ctx.route(/fonts\./, r => r.fulfill({ body: '' }));
  const pg = await ctx.newPage(); pg.on('pageerror', e => erros.push('JS: ' + e.message));
  await pg.goto(URL); await pg.click('#f-login button[type=submit]'); await pg.waitForSelector('.kpis');
  // painel: o FGTS vencido da empresa de seguranca aparece uma vez so
  const painel = await pg.textContent('#main');
  const fgts = (painel.match(/Certificado de regularidade do FGTS/g) || []).length;
  ok(painel.includes('vale para 2 contratos'), 'painel: certidão da empresa diz que vale para 2 contratos');
  ok(fgts === 2, 'painel: um alerta por fornecedor (segurança + limpeza = 2 alertas de FGTS): ' + fgts);
  // contrato novo da mesma empresa, sem certidoes anexadas nele
  const id = await pg.evaluate(() => { const a = [...document.querySelectorAll('a')]; return null; });
  await pg.goto(URL + '#contratos'); await pg.fill('#f-q', 'Vigilância da loja Centro'); await pg.waitForTimeout(150); await pg.click('#tb tr.link'); await pg.waitForSelector('#chk .chk-i');
  const chk = await pg.textContent('#chk');
  ok(/Certidão negativa federal[^]*?enviado no CTR-2026-00005[^]*?vale para os 2 contratos/.test(chk), 'checklist: certidão vem do outro contrato da mesma empresa');
  ok(/Autorização da Polícia Federal[^]*?Em dia/.test(chk), 'autorização da PF do fornecedor conta como em dia');
  ok(/Certificado de regularidade do FGTS[^]*?Vencido/.test(chk), 'FGTS vencido do fornecedor aparece vencido aqui também');
  await pg.selectOption('#u-cat', 'crf-fgts');
  ok((await pg.textContent('#u-expl')).includes('vale para os 2 contratos'), 'envio: explica que vale para os 2 contratos');
  // envia o FGTS novo neste contrato: o outro contrato fica em dia tambem
  const d = new Date(); d.setDate(d.getDate() + 120); const val = d.toISOString().slice(0, 10);
  await pg.fill('#u-val', val);
  await pg.setInputFiles('#u-file', { name: 'fgts.pdf', mimeType: 'application/pdf', buffer: Buffer.alloc(3000, 3) });
  await pg.waitForFunction(() => /Arquivo salvo|Salvo no SharePoint/.test(document.querySelector('#fila').textContent), null, { timeout: 15000 });
  await pg.waitForTimeout(600);
  ok(/Certificado de regularidade do FGTS[^]*?Em dia/.test(await pg.textContent('#chk')), 'FGTS novo: em dia neste contrato');
  await pg.goto(URL + '#contrato/5'); await pg.waitForSelector('#chk .chk-i');
  ok(/Certificado de regularidade do FGTS[^]*?enviado no[^]*?Em dia/.test(await pg.textContent('#chk')), 'e também no outro contrato da mesma empresa');
  await pg.screenshot({ path: SP + '/forn.png', fullPage: true });
  await b.close();
  console.log('\nERROS/FALHAS:', erros.length ? erros : 'nenhum');
})();
