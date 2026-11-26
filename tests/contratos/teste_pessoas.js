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
  const pg = await c.newPage();
  pg.on('pageerror', e => erros.push('JS: ' + e.message));
  await pg.exposeFunction('__abriu', u => { pg.__abertos = (pg.__abertos || []).concat(u); });
  await pg.addInitScript(() => { window.open = (u) => { window.__abriu(String(u)); return null; }; });
  await pg.goto(URL); await pg.evaluate(p => sessionStorage.setItem('proto-papel', p), papel); await pg.goto(URL);
  await pg.click('#f-login button[type=submit]');
  return pg;
}
(async () => {
  const b = await chromium.launch();
  let pg = await entrar(b, 'admin'); await pg.waitForSelector('.kpis');
  // busca rapida
  await pg.fill('#bg-q', 'vigil'); await pg.waitForTimeout(150);
  ok((await pg.textContent('#bg-res')).includes('Vigilância patrimonial'), 'busca rápida acha "Vigilância patrimonial"');
  await pg.fill('#bg-q', '11.222.333'); await pg.waitForTimeout(150);
  ok((await pg.textContent('#bg-res')).includes('Aluguel da loja Centro'), 'busca rápida por CNPJ');
  await pg.press('#bg-q', 'Enter'); await pg.waitForFunction(() => location.hash.startsWith('#contrato/'));
  ok(true, 'Enter abre o primeiro resultado');
  await pg.fill('#bg-q', 'aluguel'); await pg.waitForTimeout(150); await pg.click('.bg-res a.todos'); await pg.waitForSelector('#tb tr.link');
  ok((await pg.inputValue('#f-q')) === 'aluguel' && (await pg.$$('#tb tr.link')).length >= 4, '"Ver todos" leva para a lista já filtrada');
  // acessos: busca de pessoas
  await pg.goto(URL + '#acessos'); await pg.waitForSelector('#ac-tb tr');
  await pg.fill('#ac-q', 'supervisora'); await pg.waitForTimeout(150);
  const t1 = await pg.textContent('#ac-tb');
  ok(t1.includes('Bruna Costa') && !t1.includes('Carla') && t1.includes('1º acesso pendente') && t1.includes('2 contrato(s) escolhido(s)'), 'busca de pessoa pelo cargo; mostra 1º acesso pendente e contratos escolhidos');
  await pg.fill('#ac-q', '');
  // cadastrar pessoa com contratos escolhidos
  await pg.click('#b-novo-user'); await pg.waitForSelector('#pe-nome');
  await pg.fill('#pe-nome', 'Marcos Teste'); await pg.fill('#pe-email', 'marcos@millena.exemplo'); await pg.fill('#pe-cargo', 'Gerente de loja'); await pg.fill('#pe-funcao', 'Loja Centro');
  await pg.fill('#pe-fone', '81 98888-7777'); await pg.selectOption('#m-nivel', '2');
  await pg.check('[name=m-modo][value=lista]'); await pg.fill('#m-cbusca', 'loja centro'); await pg.waitForTimeout(100);
  const vis = await pg.$$eval('#m-cts [data-txt]', ls => ls.filter(l => !l.hidden).length); ok(vis >= 1 && vis < 5, 'busca de contrato dentro do cadastro filtra a lista: ' + vis);
  await pg.click('#m-ok'); await pg.waitForTimeout(200);
  ok((await pg.textContent('#m-msg')).includes('pelo menos um contrato'), 'contratos escolhidos sem marcar nenhum: avisa');
  await pg.check('#m-cts [data-ct="1"]'); await pg.click('#m-ok'); await pg.waitForSelector('#ea-zap');
  const q = await pg.textContent('.acesso-q');
  ok(/marcos@millena\.exemplo/.test(q) && /[a-z0-9]{4}-[a-z0-9]{4}/.test(q) && q.includes('?e=millena'), 'tela de envio com usuário, senha automática e link da empresa');
  ok((await pg.inputValue('#ea-fone')) === '(81) 98888-7777', 'WhatsApp do cadastro já preenchido');
  await pg.click('#ea-zap'); await pg.waitForTimeout(150);
  ok((pg.__abertos || []).some(u => u.startsWith('https://wa.me/5581988887777?text=')), 'Enviar pelo WhatsApp abre wa.me com o número e a mensagem');
  await pg.screenshot({ path: SP + '/pe-envio.png' });
  await pg.click('#m-ok'); await pg.waitForTimeout(400);
  const t2 = await pg.textContent('#ac-tb');
  ok((t2.match(/Marcos Teste/g) || []).length === 1 && t2.includes('Gerente de loja') && t2.includes('1 contrato(s) escolhido(s)'), 'pessoa nova na lista com cargo e o que vê');
  // editar e gerar nova senha
  await pg.click('[data-ed="u-ges"]'); await pg.waitForSelector('#pe-cargo');
  ok((await pg.inputValue('#pe-cargo')) === 'Gerente administrativa' && (await pg.isDisabled('#pe-email')), 'editar: carrega cargo; e-mail não muda');
  await pg.click('#pe-reenviar'); await pg.waitForTimeout(150); await pg.click('#m-ok'); await pg.waitForSelector('#ea-zap');
  ok(true, 'gerar nova senha abre o envio do acesso');
  await pg.click('#m-ok'); await pg.waitForTimeout(300);
  ok((await pg.textContent('#ac-tb')).match(/Carla Mendes[^]*?1º acesso pendente/) !== null, 'depois da nova senha, a Carla volta a ter 1º acesso pendente');
  // primeiro acesso da Bruna
  pg = await entrar(b, 'primeiro', { ...devices['iPhone 13'] });
  await pg.waitForSelector('#bloco-nova:not([hidden])');
  ok((await pg.textContent('#login-titulo')).includes('Crie a sua senha'), 'primeiro acesso: pede para criar a senha');
  await pg.screenshot({ path: SP + '/pe-primeiro.png' });
  await pg.fill('#n-senha', 'minhasenha1'); await pg.fill('#n-senha2', 'minhasenha1'); await pg.click('#btn-nova'); await pg.waitForSelector('.kpis');
  await pg.goto(URL + '#contratos'); await pg.waitForSelector('#tb tr.link');
  const nb = (await pg.$$('#tb tr.link')).length; ok(nb === 2, 'Bruna (2 contratos escolhidos) vê só 2 contratos: ' + nb);
  ok(!(await pg.textContent('#tabbar')).includes('Novo'), 'Bruna não tem "Novo contrato"');
  ok((await pg.evaluate(() => document.documentElement.scrollWidth)) <= 390, 'celular sem passar da tela');
  await b.close();
  console.log('\nERROS/FALHAS:', erros.length ? erros : 'nenhum');
})();
