// Decisão no prazo de aviso: renovar, renegociar ou encerrar, com quem cuida e até quando. Nada se apaga.
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
  const kp = await pg.textContent('.kpis');
  ok(/sem decisão registrada/.test(kp), 'painel: cartão do prazo de aviso conta os sem decisão');
  const at = await pg.textContent('.alerts');
  ok(at.includes('Decisão: Renegociar'), 'painel: decisão já registrada aparece no alerta');
  ok(at.includes('Sem decisão registrada'), 'painel: os outros aparecem como sem decisão');
  // lista: filtro sem decisao
  await pg.goto(URL + '#contratos?st=semdecisao'); await pg.waitForSelector('#tb tr');
  const n1 = await pg.$$eval('#tb tr.link', r => r.length);
  ok(n1 > 0, 'lista: filtro "Sem decisão" mostra contratos: ' + n1);
  // abre o primeiro e registra decisao
  await pg.click('#tb tr.link'); await pg.waitForSelector('#dec');
  ok((await pg.textContent('#dec-pill')).includes('Sem decisão'), 'contrato: selo "Sem decisão"');
  await pg.click('#b-decisao'); await pg.waitForSelector('#dc-dec');
  await pg.selectOption('#dc-dec', 'encerrar');
  ok((await pg.textContent('#dc-nota')).includes('distrato'), 'modal: explica o que fazer ao encerrar');
  await pg.selectOption('#dc-resp', { index: 1 });
  await pg.fill('#dc-obs', 'Fornecedor atrasou 3 vezes');
  await pg.click('#m-ok'); await pg.waitForFunction(() => /Encerrar/.test((document.querySelector('#dec-pill') || {}).textContent || ''));
  ok((await pg.textContent('#dec')).includes('Fornecedor atrasou 3 vezes'), 'contrato: decisão registrada aparece no cartão');
  await pg.waitForFunction(() => /decisão registrada/.test(document.querySelector('#hist').textContent));
  ok(true, 'histórico: decisão registrada');
  // muda a decisao: a anterior fica guardada
  await pg.click('#b-decisao'); await pg.waitForSelector('#dc-dec'); await pg.selectOption('#dc-dec', 'renovar'); await pg.click('#m-ok');
  await pg.waitForFunction(() => /Renovar/.test(document.querySelector('#dec-pill').textContent));
  ok(/Decisões anteriores \(1\)/.test(await pg.textContent('#dec')), 'decisão anterior continua guardada');
  await pg.goto(URL + '#contratos?st=semdecisao'); await pg.waitForSelector('#tb tr');
  const n2 = await pg.$$eval('#tb tr.link', r => r.length);
  ok(n2 === n1 - 1, `lista: saiu do filtro "Sem decisão" (${n1} → ${n2})`);
  await pg.screenshot({ path: SP + '/decisao.png' });
  await b.close();
  console.log('\nERROS/FALHAS:', erros.length ? erros : 'nenhum');
})();
