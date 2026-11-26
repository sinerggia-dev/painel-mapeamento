const { chromium, devices } = require('playwright');
const SP = process.argv[2], fs = require('fs');
const LIB = fs.readFileSync(SP + '/xlsxpkg/package/dist/xlsx.full.min.js', 'utf8');
const XLSX = require(SP + '/xlsxpkg/package/xlsx.js');
const URL = 'file://' + SP + '/gestao-contratos.html';
(async () => {
  const b = await chromium.launch(); const erros = [];
  for (const [nome, dev] of [['pixel', devices['Pixel 7']], ['iphone', devices['iPhone SE']]]) {
    const ctx = await b.newContext({ ...dev });
    await ctx.route(/cdn\.jsdelivr\.net\/npm\/xlsx/, r => r.fulfill({ body: LIB, contentType: 'application/javascript' }));
    await ctx.route(/fonts\.(googleapis|gstatic)\.com/, r => r.fulfill({ body: '', contentType: 'text/css' }));
    const pg = await ctx.newPage(); pg.on('pageerror', e => erros.push(nome + ' JS: ' + e.message));
    await pg.goto(URL); await pg.click('#f-login button[type=submit]'); await pg.waitForSelector('.kpis');
    const W = dev.viewport.width; const small = [];
    for (const h of ['painel','contratos','contrato/1','contrato/5','novo','importar','categorias','unidades','acessos','identidade','naturezas','sharepoint','perfil']) {
      await pg.goto(URL + '#' + h); await pg.waitForTimeout(500);
      const r = await pg.evaluate(() => ({ w: document.documentElement.scrollWidth,
        // campos com letra menor que 16px fazem o iPhone dar zoom
        z: [...document.querySelectorAll('input,select,textarea')].filter(e => e.offsetParent && !['radio', 'checkbox', 'color', 'file', 'hidden'].includes(e.type) && parseFloat(getComputedStyle(e).fontSize) < 16).map(e => e.id || e.type),
        // alvos de toque pequenos
        t: [...document.querySelectorAll('button,a.btn,select,input:not([type=checkbox])')].filter(e => e.offsetParent && e.getBoundingClientRect().height < 36).map(e => e.id || e.textContent.trim().slice(0, 20)) }));
      if (r.w > W) erros.push(`${nome} #${h} largura ${r.w}>${W}`);
      r.z = r.z.filter(x => !['proto-papel','c-conf','c-enc','checkbox','id-cor','id-fundo','nt-usa','nt-hash'].includes(x)); if (r.z.length) erros.push(`${nome} #${h} campos <16px: ${r.z.join(',')}`);
      if (r.t.length) small.push(`#${h}: ${r.t.slice(0, 6).join(' | ')}`);
    }
    // modal categoria e importacao com erro
    await pg.goto(URL + '#categorias'); await pg.click('#b-nova-cat'); await pg.waitForTimeout(400);
    await pg.screenshot({ path: `${SP}/${nome}-modal.png` });
    await pg.click('[data-fechar].btn');
    await pg.goto(URL + '#importar'); await pg.waitForTimeout(300);
    const w = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(w, XLSX.utils.aoa_to_sheet([['Unidade *','Fornecedor *','CPF/CNPJ do fornecedor','Código do fornecedor no seu sistema','Categoria *','Código da natureza no seu sistema','Objeto do contrato *'],['99','Fulano de Tal com nome comprido','123','','Categoria inexistente','','']]), 'Contratos');
    XLSX.writeFile(w, SP + '/m.xlsx'); await pg.setInputFiles('#imp-file', SP + '/m.xlsx'); await pg.waitForSelector('#imp-res .msg');
    const wi = await pg.evaluate(() => document.documentElement.scrollWidth); if (wi > W) erros.push(`${nome} importar com erros largura ${wi}`);
    await pg.screenshot({ path: `${SP}/${nome}-imp.png`, fullPage: true });
    await pg.goto(URL + '#contrato/1'); await pg.waitForSelector('#chk .chk-i'); await pg.screenshot({ path: `${SP}/${nome}-contrato.png`, fullPage: true });
    console.log(nome, 'alvos pequenos:', small.length ? small : 'nenhum');
  }
  console.log('PROBLEMAS:', erros.length ? erros : 'nenhum'); await b.close();
})();
