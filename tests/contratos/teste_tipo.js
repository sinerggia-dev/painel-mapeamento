const { chromium, devices } = require('playwright');
const fs = require('fs');
const SP = process.argv[2];
const LIB = fs.readFileSync(SP + '/xlsxpkg/package/dist/xlsx.full.min.js', 'utf8');
const URL = 'file://' + SP + '/gestao-contratos.html';
const erros = [];
const ok = (c, m) => { console.log((c ? 'OK   ' : 'FALHA ') + m); if (!c) erros.push(m); };
(async () => {
  const b = await chromium.launch();
  for (const [nome, opts] of [['pc', { viewport: { width: 1440, height: 900 } }], ['cel', devices['iPhone 13']]]) {
    const c = await b.newContext(opts);
    await c.route(/jsdelivr/, r => r.fulfill({ body: LIB, contentType: 'application/javascript' }));
    await c.route(/fonts\./, r => r.fulfill({ body: '' }));
    const pg = await c.newPage(); pg.on('pageerror', e => erros.push('JS: ' + e.message));
    await pg.goto(URL); await pg.click('#f-login button[type=submit]'); await pg.waitForSelector('.kpis');
    await pg.goto(URL + '#contratos'); await pg.waitForSelector('#tb tr.link');
    const tipos = await pg.$$eval('#tb .col-tipo .pill', ps => ps.map(p => p.textContent));
    ok(tipos.includes('Novo') && tipos.includes('Renovação') && tipos.includes('Aditivado'), nome + ': coluna Tipo com Novo, Renovação e Aditivado');
    await pg.selectOption('#f-tp', 'renovacao');
    ok((await pg.$$('#tb tr.link')).length === 1 && (await pg.textContent('#tb')).includes('Aluguel da loja Shopping'), nome + ': filtro Renovação');
    await pg.selectOption('#f-tp', 'distrato');
    ok((await pg.$$('#tb tr.link')).length === 1 && (await pg.textContent('#tb')).includes('Frete de entregas'), nome + ': filtro Distrato mostra o encerrado por distrato');
    await pg.selectOption('#f-tp', 'aditivo');
    ok((await pg.textContent('#tb')).includes('Limpeza das lojas'), nome + ': filtro Aditivado');
    await pg.selectOption('#f-tp', '');
    if (nome === 'cel') {
      ok((await pg.evaluate(() => document.documentElement.scrollWidth)) <= 390, 'celular: lista sem passar da tela');
      ok(await pg.isVisible('#tb tr.link .col-tipo .pill'), 'celular: tipo aparece no cartão');
      await pg.screenshot({ path: SP + '/tp-cel.png' });
    } else {
      await pg.screenshot({ path: SP + '/tp-pc.png' });
      await pg.goto(URL + '#contrato/10'); await pg.waitForSelector('#c-tipo');
      await pg.selectOption('#c-tipo', 'renovacao'); await pg.click('#b-salvar'); await pg.waitForTimeout(400);
      ok((await pg.textContent('.card-h')).includes('Renovação'), 'contrato marcado como renovação mostra o selo');
    }
  }
  await b.close();
  console.log('\nERROS/FALHAS:', erros.length ? erros : 'nenhum');
})();
