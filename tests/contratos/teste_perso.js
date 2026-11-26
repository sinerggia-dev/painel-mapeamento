const { chromium, devices } = require('playwright');
const fs = require('fs');
const SP = process.argv[2];
const LIB = fs.readFileSync(SP + '/xlsxpkg/package/dist/xlsx.full.min.js', 'utf8');
const URL = 'file://' + SP + '/gestao-contratos.html';
const erros = [];
const ok = (c, m) => { console.log((c ? 'OK   ' : 'FALHA ') + m); if (!c) erros.push(m); };
const LOGO = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAYAAADED76LAAAAHElEQVQoU2NkYPj/n4EIwDiqEF8YjIYC/jAYDQUAbBMIAf3Yq5kAAAAASUVORK5CYII=', 'base64');
async function ctx(b, opts = {}) {
  const c = await b.newContext({ viewport: { width: 1440, height: 900 }, ...opts });
  await c.route(/jsdelivr/, r => r.fulfill({ body: LIB, contentType: 'application/javascript' }));
  await c.route(/fonts\./, r => r.fulfill({ body: '', contentType: 'text/css' }));
  const pg = await c.newPage();
  pg.on('pageerror', e => erros.push('JS: ' + e.message));
  return pg;
}
(async () => {
  const b = await chromium.launch();
  // tela de entrada pelo link da empresa
  let pg = await ctx(b);
  await pg.goto(URL + '?e=millena'); await pg.waitForTimeout(500);
  ok(!!(await pg.$('#login-brand img[alt="Millena Móveis"]')) && (await pg.getAttribute('html', 'data-theme')) === 'dark', 'link da empresa: tela de entrada com a logo e o tema escuro da Millena');
  ok((await pg.evaluate(() => document.documentElement.style.getPropertyValue('--acc'))) === '#F8B900', 'link da empresa: cor da Millena na entrada');
  await pg.screenshot({ path: SP + '/p-login.png' });
  await pg.goto(URL); await pg.waitForTimeout(300);
  await pg.click('#f-login button[type=submit]'); await pg.waitForSelector('.kpis');
  // personalizacao (administracao)
  await pg.goto(URL + '#identidade'); await pg.waitForSelector('#id-tema');
  ok((await pg.textContent('#nav')).includes('Personalização'), 'menu com "Personalização"');
  ok((await pg.inputValue('#id-link')).endsWith('?e=millena'), 'link de entrada da empresa na tela');
  await pg.setInputFiles('#id-logo', { name: 'logo.png', mimeType: 'image/png', buffer: LOGO }); await pg.waitForTimeout(300);
  await pg.check('[name=id-menu][value=marca]'); await pg.selectOption('#id-tema', 'escuro');
  ok((await pg.getAttribute('#id-prev', 'data-menu')) === 'marca', 'prévia mostra o menu na cor da marca');
  await pg.click('#id-salvar'); await pg.waitForTimeout(400);
  // como a empresa ve
  await pg.selectOption('#sel-emp', 'millena'); await pg.waitForTimeout(700);
  ok(await pg.evaluate(() => document.body.classList.contains('menu-marca')), 'como a empresa vê: menu na cor da marca');
  ok((await pg.getAttribute('html', 'data-theme')) === 'dark', 'como a empresa vê: tema escuro da empresa');
  ok((await pg.getAttribute('#favicon', 'href')).startsWith('data:image'), 'ícone da aba = logo da empresa');
  await pg.screenshot({ path: SP + '/p-marca.png' });
  // pessoa escolhe tema so para ela
  await pg.goto(URL + '#perfil'); await pg.waitForSelector('#pf-tema');
  await pg.selectOption('#pf-tema', 'claro'); await pg.waitForTimeout(200);
  ok((await pg.getAttribute('html', 'data-theme')) === 'light', 'Meu perfil: tema claro vale mais que o da empresa');
  await pg.fill('#pf-s1', '1234'); await pg.click('#pf-senha'); await pg.waitForTimeout(200);
  ok((await pg.textContent('#pf-msg')).includes('8 caracteres'), 'Meu perfil: senha curta recusada');
  await pg.fill('#pf-s1', 'senhanova123'); await pg.fill('#pf-s2', 'senhanova123'); await pg.click('#pf-senha'); await pg.waitForTimeout(300);
  ok((await pg.textContent('#pf-msg')).includes('Senha alterada'), 'Meu perfil: senha alterada');
  await pg.screenshot({ path: SP + '/p-perfil.png' });
  // menu claro
  await pg.selectOption('#sel-emp', '__adm'); await pg.waitForTimeout(600);
  ok(!(await pg.evaluate(() => document.body.classList.contains('menu-marca'))), 'administração: menu volta ao padrão Sinergia');
  await pg.goto(URL + '#identidade'); await pg.waitForSelector('#id-tema'); await pg.check('[name=id-menu][value=claro]'); await pg.click('#id-salvar'); await pg.waitForTimeout(400);
  await pg.selectOption('#sel-emp', 'millena'); await pg.waitForTimeout(600);
  ok(await pg.evaluate(() => document.body.classList.contains('menu-claro')), 'como a empresa vê: menu claro');
  await pg.screenshot({ path: SP + '/p-claro.png' });
  // usuaria da empresa: perfil sem link de personalizacao
  pg = await ctx(b, { ...devices['iPhone 13'] });
  await pg.goto(URL); await pg.evaluate(() => sessionStorage.setItem('proto-papel', 'gestor')); await pg.goto(URL);
  await pg.click('#f-login button[type=submit]'); await pg.waitForSelector('.kpis');
  await pg.goto(URL + '#perfil'); await pg.waitForSelector('#pf-tema');
  const t = await pg.textContent('#main');
  ok(t.includes('Carla Mendes') && !t.includes('Personalização</a>'), 'usuária da empresa: Meu perfil sem acesso à personalização');
  ok((await pg.evaluate(() => document.documentElement.scrollWidth)) <= 390, 'Meu perfil no celular sem passar da tela');
  await pg.goto(URL + '#identidade'); await pg.waitForTimeout(200); ok((await pg.textContent('#main')).includes('Sem permissão'), 'usuária da empresa não abre Personalização');
  await b.close();
  console.log('\nERROS/FALHAS:', erros.length ? erros : 'nenhum');
})();
