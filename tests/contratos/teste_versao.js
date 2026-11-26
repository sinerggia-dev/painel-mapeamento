const { chromium, devices } = require('playwright');
const fs = require('fs');
const SP = process.argv[2];
const LIB = fs.readFileSync(SP + '/xlsxpkg/package/dist/xlsx.full.min.js', 'utf8');
const URL = 'file://' + SP + '/gestao-contratos.html';
const erros = [];
const ok = (c, m) => { console.log((c ? 'OK   ' : 'FALHA ') + m); if (!c) erros.push(m); };
async function entrar(b, papel, opts = { viewport: { width: 1440, height: 900 } }) {
  const c = await b.newContext(opts);
  await c.route(/jsdelivr/, r => r.fulfill({ body: LIB, contentType: 'application/javascript' }));
  await c.route(/fonts\./, r => r.fulfill({ body: '', contentType: 'text/css' }));
  const pg = await c.newPage(); pg.on('pageerror', e => erros.push('JS: ' + e.message));
  await pg.goto(URL); await pg.evaluate(p => sessionStorage.setItem('proto-papel', p), papel); await pg.goto(URL);
  await pg.click('#f-login button[type=submit]'); await pg.waitForSelector('.kpis');
  return pg;
}
(async () => {
  const b = await chromium.launch();
  // usuaria da empresa (acesso completo nas filiais dela)
  let pg = await entrar(b, 'gestor');
  await pg.goto(URL + '#contrato/3'); await pg.waitForSelector('#u-cat');
  ok(!(await pg.$('.sp')), 'empresa: sem o quadro azul do SharePoint');
  ok((await pg.textContent('.anexar h3')) === 'Anexar documento', 'área com título "Anexar documento"');
  const opc = await pg.$$eval('#u-cat option', os => os.map(o => o.textContent));
  ok(opc.includes('Contrato atualizado (nova versão do contrato assinado)'), 'opção clara: "Contrato atualizado (nova versão do contrato assinado)"');
  await pg.selectOption('#u-cat', 'contrato-assinado');
  ok(!(await pg.isHidden('#u-mud')) && (await pg.textContent('#u-expl')).includes('versão vigente'), 'nova versão: explica e pede o que mudou');
  await pg.click('#dz'); await pg.waitForTimeout(200);
  ok((await pg.textContent('#toast')).includes('o que mudou'), 'sem dizer o que mudou, não deixa enviar');
  await pg.fill('#u-oque', 'Valor reajustado e nova cláusula de manutenção'); await pg.fill('#u-novoval', '33.000,00'); await pg.fill('#u-novofim', '2029-12-31');
  await pg.setInputFiles('#u-file', { name: 'contrato-v2.pdf', mimeType: 'application/pdf', buffer: Buffer.alloc(8000, 2) });
  await pg.waitForFunction(() => /Nova versão do contrato/.test((document.querySelector('#tlv') || {}).textContent || ''), null, { timeout: 15000 });
  await pg.waitForSelector('#d-lista .doc');
  const docs = await pg.textContent('#d-lista');
  ok(docs.includes('versão vigente') && docs.includes('versão anterior · guardada'), 'arquivos: "versão vigente" e "versão anterior · guardada"');
  ok(docs.includes('Nova versão do contrato: Valor reajustado'), 'arquivo mostra o que mudou');
  ok(!docs.includes('SharePoint') && !(await pg.$('.sp')), 'nenhuma menção ao SharePoint para a empresa');
  ok((await pg.inputValue('#c-fim')) === '2029-12-31' && (await pg.inputValue('#c-valor')).startsWith('33000'), 'contrato passa a ter o novo valor e vencimento');
  const tl = await pg.textContent('#tlv'); ok(tl.includes('Aditivo de valor') && tl.includes('32.000,00') && tl.includes('33.000,00'), 'linha do tempo: valor anterior → novo');
  ok(!(await pg.$('[data-apagar-doc]')) && !(await pg.textContent('#d-lista')).includes('Apagar'), 'não existe botão de apagar arquivo');
  await pg.screenshot({ path: SP + '/v-empresa.png', fullPage: true });
  // documento com validade continua simples
  await pg.selectOption('#u-cat', 'cndt'); ok(await pg.isHidden('#u-mud') && !(await pg.isHidden('#u-val-f')), 'certidão: pede validade, não pede "o que mudou"');
  // admin ve o quadro do SharePoint
  pg = await entrar(b, 'admin');
  await pg.goto(URL + '#contrato/3'); await pg.waitForSelector('#u-cat');
  ok(!!(await pg.$('.sp')) && (await pg.textContent('.sp')).includes('Só a Sinergia vê'), 'admin Sinergia: vê o quadro do SharePoint');
  // admin vendo como a empresa: nao ve
  await pg.selectOption('#sel-emp', 'millena'); await pg.waitForTimeout(700);
  await pg.goto(URL + '#contrato/3'); await pg.waitForSelector('#u-cat');
  ok(!(await pg.$('.sp')), 'admin vendo como a empresa: sem o quadro do SharePoint');
  // celular
  pg = await entrar(b, 'gestor', { ...devices['iPhone 13'] });
  await pg.goto(URL + '#contrato/3'); await pg.waitForSelector('#u-cat'); await pg.selectOption('#u-cat', 'aditivo');
  ok((await pg.evaluate(() => document.documentElement.scrollWidth)) <= 390, 'celular: área de anexo sem passar da tela');
  await pg.screenshot({ path: SP + '/v-iphone.png', fullPage: true });
  await b.close();
  console.log('\nERROS/FALHAS:', erros.length ? erros : 'nenhum');
})();
