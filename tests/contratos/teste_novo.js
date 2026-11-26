const { chromium, devices } = require('playwright'); const fs = require('fs'); const SP = process.argv[2];
const LIB = fs.readFileSync(SP + '/xlsxpkg/package/dist/xlsx.full.min.js', 'utf8'); const URL = 'file://' + SP + '/gestao-contratos.html';
const erros = []; const ok = (c, m) => { console.log((c ? 'OK   ' : 'FALHA ') + m); if (!c) erros.push(m); };
(async () => { const b = await chromium.launch();
  for (const [nome, opts] of [['pc', { viewport: { width: 1440, height: 900 } }], ['iphone', devices['iPhone 13']]]) {
    const ctx = await b.newContext(opts); await ctx.route(/jsdelivr/, r => r.fulfill({ body: LIB, contentType: 'application/javascript' })); await ctx.route(/fonts\./, r => r.fulfill({ body: '' }));
    const pg = await ctx.newPage(); pg.on('pageerror', e => erros.push('JS: ' + e.message));
    await pg.goto(URL); await pg.evaluate(() => sessionStorage.setItem('proto-papel', 'gestor')); await pg.goto(URL);
    await pg.click('#f-login button[type=submit]'); await pg.waitForSelector('.kpis');
    await pg.goto(URL + '#novo'); await pg.waitForSelector('#n-dz');
    ok(!(await pg.textContent('#main')).includes('Importar planilha'), nome + ': sem botão "Importar planilha" no novo contrato');
    ok(await pg.isHidden('#n-val-f'), nome + ': contrato assinado não pede validade');
    await pg.selectOption('#c-natsel', '702'); await pg.fill('#c-titulo', 'Aluguel loja teste'); await pg.fill('#c-forn', 'Locadora Teste'); await pg.selectOption('#c-uni', '2');
    await pg.setInputFiles('#n-file', { name: 'contrato-assinado.pdf', mimeType: 'application/pdf', buffer: Buffer.alloc(9000, 1) });
    ok((await pg.textContent('#n-lista')).includes('contrato-assinado.pdf'), nome + ': arquivo escolhido aparece na lista');
    if (nome === 'iphone') { ok((await pg.evaluate(() => document.documentElement.scrollWidth)) <= 390, 'iphone: novo contrato sem passar da tela'); await pg.screenshot({ path: SP + '/n-iphone.png', fullPage: true }); }
    else await pg.screenshot({ path: SP + '/n-pc.png', fullPage: true });
    await pg.click('#b-salvar'); await pg.waitForFunction(() => location.hash.startsWith('#contrato/'));
    await pg.waitForFunction(() => /Arquivo salvo|Salvo no SharePoint/.test((document.querySelector('#fila') || {}).textContent || ''), null, { timeout: 15000 });
    await pg.waitForTimeout(800);
    const chk = await pg.textContent('#chk');
    ok(/Contrato principal assinado\s*Em dia/.test(chk), nome + ': depois de salvar, contrato assinado anexado (Em dia)');
    ok((await pg.textContent('#d-lista')).includes('contrato-assinado'), nome + ': arquivo na lista do contrato');
  }
  await b.close(); console.log('\nERROS/FALHAS:', erros.length ? erros : 'nenhum'); })();
