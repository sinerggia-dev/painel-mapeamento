const { chromium } = require('playwright');
const SP = process.argv[2];
const XLSX = require(SP + '/xlsxpkg/package/xlsx.js');
const fs = require('fs');
const LIB = fs.readFileSync(SP + '/xlsxpkg/package/dist/xlsx.full.min.js', 'utf8');
const URL = 'file://' + SP + '/gestao-contratos.html';
const erros = [];
const ok = (cond, msg) => { console.log((cond ? 'OK   ' : 'FALHA ') + msg); if (!cond) erros.push(msg); };
async function nova(browser, opts, papel = 'admin') {
  const ctx = await browser.newContext({ ...opts, acceptDownloads: true });
  await ctx.route(/cdn\.jsdelivr\.net\/npm\/xlsx/, r => r.fulfill({ body: LIB, contentType: 'application/javascript' }));
  await ctx.route(/fonts\.(googleapis|gstatic)\.com/, r => r.fulfill({ body: '', contentType: 'text/css' }));
  const pg = await ctx.newPage();
  pg.on('pageerror', e => erros.push('JS: ' + e.message));
  pg.on('console', m => { if (m.type() === 'error') erros.push('console: ' + m.text()); });
  await pg.goto(URL);
  await pg.evaluate(p => { sessionStorage.setItem('proto-papel', p); }, papel);
  await pg.goto(URL);
  await pg.click('#f-login button[type=submit]');
  await pg.waitForSelector('.kpis, .card');
  await pg.waitForFunction(() => window.XLSX);
  return pg;
}
const txt = async (pg, s) => (await pg.textContent(s)) || '';
(async () => {
  const browser = await chromium.launch();
  // ───── ADMIN, computador ─────
  let pg = await nova(browser, { viewport: { width: 1440, height: 900 } });
  ok((await txt(pg, '.kpis')).includes('Documentos pendentes'), 'painel com KPI de documentos pendentes');
  const docCard = await txt(pg, '.dash');
  ok(docCard.includes('AVCB') && docCard.includes('Certificado de regularidade do FGTS'), 'painel lista documentos vencidos/vencendo (AVCB, FGTS)');
  ok(docCard.includes('Custo mensal por classe') && docCard.includes('7 - MANUTENCAO DE PREDIOS'), 'painel custo por classe da natureza');
  await pg.screenshot({ path: SP + '/s-painel.png', fullPage: true });
  // lista
  await pg.goto(URL + '#contratos'); await pg.waitForSelector('#tb tr');
  ok((await pg.$$('#tb tr.link')).length === 13, 'lista mostra 13 contratos ativos (o frete foi encerrado por distrato)');
  await pg.selectOption('#f-gr', 'nt:702'); ok((await pg.$$('#tb tr.link')).length === 4, 'filtro pela natureza 702 = 4');
  ok((await txt(pg, '#tb')).includes('702 · ALUGUEL DE IMOVEIS'), 'lista mostra a natureza');
  await pg.selectOption('#f-gr', ''); await pg.selectOption('#f-st', 'pendente');
  const nPend = (await pg.$$('#tb tr.link')).length; ok(nPend >= 4, 'filtro documentos pendentes: ' + nPend);
  await pg.selectOption('#f-st', ''); await pg.fill('#f-q', '11.222.333'); ok((await pg.$$('#tb tr.link')).length === 1, 'busca por CNPJ formatado');
  // detalhe
  await pg.goto(URL + '#contrato/1'); await pg.waitForSelector('#chk .chk-i');
  const chk = await txt(pg, '#chk');
  ok(chk.includes('Vencido') && chk.includes('Vence logo') && chk.includes('Em dia'), 'checklist do aluguel: vencido, vence logo, em dia');
  ok((await txt(pg, '.detail')).includes('Próximo reajuste (IGP-M)'), 'resumo mostra próximo reajuste');
  ok((await pg.inputValue('#u-cat')) === 'avcb', 'upload já sugere o primeiro documento pendente (AVCB)');
  await pg.screenshot({ path: SP + '/s-contrato.png', fullPage: true });
  // envio de AVCB novo com validade
  await pg.click('#dz'); await pg.waitForTimeout(300);
  ok((await txt(pg, '#toast')).includes('validade'), 'sem validade: pede a data antes de enviar');
  await pg.fill('#u-val', '2028-05-10');
  await pg.setInputFiles('#u-file', { name: 'avcb.pdf', mimeType: 'application/pdf', buffer: Buffer.alloc(5000, 1) });
  await pg.waitForFunction(() => document.querySelector('#fila') && /Salvo no SharePoint/.test(document.querySelector('#fila').textContent), null, { timeout: 15000 });
  await pg.waitForTimeout(800);
  const chk2 = await txt(pg, '#chk'); ok(chk2.includes('10/05/2028'), 'depois do envio, AVCB válido até 10/05/2028');
  ok((await txt(pg, '#fila')).includes('702 - ALUGUEL DE IMOVEIS/CTR-2026-00001'), 'pasta no SharePoint pela natureza');
  // novo contrato
  await pg.goto(URL + '#novo'); await pg.waitForSelector('#c-natsel');
  const natOpts = await pg.$$eval('#c-natsel option', os => os.map(o => o.value).filter(Boolean));
  ok(natOpts.includes('721') && !natOpts.includes('727') && !natOpts.includes('212') && !natOpts.includes('1830'), 'cadastro só oferece naturezas com contrato (sem #, folha, receita)');
  await pg.selectOption('#c-natsel', '721');
  ok((await txt(pg, '#c-catinfo')).includes('Vigilância e segurança'), 'natureza mostra a categoria Sinergia');
  ok((await pg.inputValue('#c-reaj')) === 'Convenção coletiva', 'categoria preenche reajuste padrão');
  ok(!(await pg.$eval('#c-dica', e => e.hidden)) && (await txt(pg, '#c-dica')).includes('Lei 6.019'), 'dica da categoria aparece');
  await pg.fill('#c-titulo', 'Vigilância loja Agreste'); await pg.fill('#c-forn', 'Vigia Teste Ltda'); await pg.fill('#c-doc', '11.222.333/0001-00');
  await pg.selectOption('#c-uni', '8'); await pg.fill('#c-ini', '2026-01-10'); await pg.fill('#c-valor', '3.500,00');
  await pg.click('#b-salvar'); await pg.waitForTimeout(200);
  ok((await txt(pg, '#c-msg')).includes('CNPJ inválido'), 'CNPJ errado é recusado');
  await pg.fill('#c-doc', '11.222.333/0001-81'); await pg.click('#b-salvar');
  await pg.waitForFunction(() => location.hash.startsWith('#contrato/'));
  await pg.waitForSelector('#chk .chk-i');
  ok((await txt(pg, '.detail')).includes('janeiro'), 'mês do reajuste = mês do início (janeiro)');
  ok((await pg.$$('#chk .chk-i')).length === 5, 'vigilância exige 5 documentos (contrato + 4)');
  // CNPJ alfanumerico valido (calculado)
  const cnpjAlfa = (() => { const b = '12ABC34501DE'; const val = c => c.charCodeAt(0) - 48; const dv = (s, p) => { let t = 0; for (let i = 0; i < p.length; i++) t += val(s[i]) * p[i]; const r = t % 11; return r < 2 ? 0 : 11 - r; }; const d1 = dv(b, [5,4,3,2,9,8,7,6,5,4,3,2]); const d2 = dv(b + d1, [6,5,4,3,2,9,8,7,6,5,4,3,2]); return b + d1 + d2; })();
  await pg.fill('#c-doc', cnpjAlfa); await pg.click('#b-salvar'); await pg.waitForTimeout(400);
  ok(!(await txt(pg, '#c-msg')).includes('inválido'), 'CNPJ com letras (novo formato) é aceito: ' + cnpjAlfa);
  // planilha modelo
  await pg.goto(URL + '#importar'); await pg.waitForSelector('#b-modelo');
  const [dl] = await Promise.all([pg.waitForEvent('download'), pg.click('#b-modelo')]);
  const p = SP + '/modelo.xlsx'; await dl.saveAs(p);
  const wb = XLSX.readFile(p);
  ok(['Contratos','Exemplo','Instruções','Naturezas','Filiais'].every(n => wb.SheetNames.includes(n)), 'modelo tem as 5 abas: ' + wb.SheetNames.join(', '));
  const cab = XLSX.utils.sheet_to_json(wb.Sheets.Contratos, { header: 1 })[0];
  ok(cab.length === 15 && cab[0] === 'Filial *' && cab[4] === 'Natureza *' && cab[14] === 'Tipo (Novo ou Renovação)', 'aba Contratos com 15 colunas, Natureza e Tipo no fim');
  ok(XLSX.utils.sheet_to_json(wb.Sheets.Naturezas).length === 15, 'aba Naturezas com as 15 que têm contrato: ' + XLSX.utils.sheet_to_json(wb.Sheets.Naturezas).length);
  // preenche com erros
  const linhasRuins = [cab,
    ['1', 'Locadora Boa Ltda', '11.222.333/0001-81', 'X1', '702', 'Aluguel matriz anexo', '5.000,00', '01/02/2025', '31/01/2030', '', '', '', 'Não', ''],
    ['99', 'Fulano', '123', '', '999', '', 'abc', '31/02/2025', '01/01/2020', '', 'XPTO', '13', 'talvez', ''],
    ['1', 'Locadora Boa Ltda', '11.222.333/0001-81', 'X1', 702, 'Aluguel matriz anexo', 5000, 45700, 47513, '', '', '', '', ''],
    ['2', 'Imobiliária Exemplo Ltda', '11.222.333/0001-81', '', '702', 'Aluguel da loja Centro', '', '', '', '', '', '', '', ''],
    ['3', 'Vigia Antigo', '', '', '727', 'Seg antiga', '', '', '', '', '', '', '', ''],
    ['3', 'Eletricista', '', '', '808', 'Elétrica', '', '', '', '', '', '', '', ''],
  ];
  const mk = (rows, nome) => { const w = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(w, XLSX.utils.aoa_to_sheet(rows), 'Contratos'); const f = SP + '/' + nome; XLSX.writeFile(w, f); return f; };
  await pg.setInputFiles('#imp-file', mk(linhasRuins, 'ruim.xlsx'));
  await pg.waitForSelector('#imp-res .msg');
  const r1 = await txt(pg, '#imp-res');
  ok(r1.includes('Nada foi gravado'), 'planilha com erro: nada gravado');
  for (const e of ['Natureza "999" não existe', 'está desativada', 'não está marcada como "tem contrato"', 'CPF tem 11 números', 'Valor mensal inválido', 'Data de início inválida', 'Índice de reajuste "XPTO"', 'Mês do reajuste inválido', 'Renovação automática', 'Linha repetida (igual à linha 2)', 'Este contrato já está cadastrado', 'Objeto do contrato em branco'])
    ok(r1.includes(e), 'erro detectado: ' + e);
  ok(!r1.includes('Unidade "99"'), 'admin: unidade 99 nova não vira erro (será criada)');
  // planilha boa
  const boas = [cab,
    ['1', 'Locadora Boa Ltda', '11.222.333/0001-81', 'X1', '702', 'Aluguel matriz anexo', '5.000,00', '01/02/2025', '31/01/2030', '', '', '', 'Não', ''],
    ['Loja Norte (exemplo)', 'Frete Rápido', '', '', 102, 'Entregas zona norte', 1200.5, '15/03/2026', '', '', 'ipca', 'março', 'sim', 'teste'],
    ['50', 'Nova Loja Advocacia', '', '', '1501', 'Advogado filial 50', '', '', '', '45', '', '', '', ''],
  ];
  await pg.setInputFiles('#imp-file', mk(boas, 'boa.xlsx'));
  await pg.waitForSelector('#imp-ok');
  const r2 = await txt(pg, '#imp-res');
  ok(r2.includes('3 contrato(s) prontos') && r2.includes('filial(s) nova(s)') && r2.includes('sem vencimento'), 'planilha boa: 3 prontos, 1 unidade nova, avisos');
  await pg.click('#imp-ok'); await pg.waitForFunction(() => /gravados/.test(document.querySelector('#imp-res').textContent));
  await pg.goto(URL + '#contratos'); await pg.waitForSelector('#tb tr.link');
  ok((await pg.$$('#tb tr.link')).length === 17, 'depois da importação: 17 contratos ativos');
  await pg.fill('#f-q', 'Entregas zona norte'); await pg.click('#tb tr.link'); await pg.waitForSelector('#c-aviso');
  ok((await pg.inputValue('#c-aviso')) === '30' && (await pg.inputValue('#c-mes')) === '3' && (await pg.inputValue('#c-reaj')) === 'IPCA' && (await pg.inputValue('#c-uni')) === '4', 'importado: unidade por nome, aviso padrão, reajuste IPCA, mês março');
  // categorias
  await pg.goto(URL + '#categorias'); await pg.waitForSelector('#b-nova-cat');
  ok((await txt(pg, '#cat-lista')).includes('Locação de empilhadeira') && (await txt(pg, '#cat-lista')).includes('Padrão'), 'tela de categorias: padrão + da empresa');
  await pg.click('#b-nova-cat'); await pg.fill('#k-nome', 'aluguel de imóvel comercial'); await pg.click('#m-ok'); await pg.waitForTimeout(200);
  ok((await txt(pg, '#m-msg')).includes('Já existe'), 'não deixa criar categoria repetida');
  await pg.fill('#k-nome', 'Manutenção de geradores'); await pg.selectOption('#k-grupo', 'manutencao'); await pg.check('[data-doc="seguro"]'); await pg.click('#m-ok');
  await pg.waitForTimeout(400); ok((await txt(pg, '#cat-lista')).includes('Manutenção de geradores'), 'categoria nova criada');
  await pg.screenshot({ path: SP + '/s-categorias.png', fullPage: false });
  // unidades
  await pg.goto(URL + '#unidades'); await pg.waitForSelector('#uni-tb');
  ok((await txt(pg, '#uni-tb')).includes('50'), 'unidade 50 criada pela importação');
  await pg.click('[data-uni="50"]'); await pg.fill('#un-nome', 'Loja Nova 50'); await pg.click('#m-ok'); await pg.waitForTimeout(300);
  ok((await txt(pg, '#uni-tb')).includes('Loja Nova 50'), 'renomear unidade');
  // acessos
  await pg.goto(URL + '#acessos'); await pg.waitForSelector('[data-ed]');
  ok((await txt(pg, '#ac-tb')).includes('Imóveis e ocupação'), 'acessos mostram grupos');
  await pg.click('[data-ed="u-jur"]'); await pg.waitForSelector('[data-gr]'); ok((await pg.$$('[data-un]')).length === 9, 'modal de acesso lista 9 unidades');
  await pg.click('[data-fechar].btn');
  await pg.goto(URL + '#importar'); await pg.screenshot({ path: SP + '/s-importar.png' });
  // administracao: sempre com a cara da Sinergia
  await pg.goto(URL + '#painel'); await pg.waitForSelector('.kpis');
  ok((await txt(pg, '#side-brand')).includes('SINERGIA') && !(await pg.evaluate(() => document.documentElement.style.getPropertyValue('--acc'))) && await pg.isVisible('#selo-admin'), 'administração: identidade da Sinergia e selo de admin');
  ok((await txt(pg, '#nav')).includes('Acessos') && (await pg.inputValue('#sel-adm')) === 'millena', 'administração: menu de admin e Millena em gestão');
  // ver como a empresa: exatamente o que a Millena ve
  await pg.selectOption('#sel-emp', 'millena'); await pg.waitForTimeout(700);
  ok(!!(await pg.$('#side-brand img[alt="Millena Móveis"]')), 'como a empresa vê: logo da Millena Móveis no menu');
  ok((await pg.evaluate(() => document.documentElement.style.getPropertyValue('--acc'))) === '#F8B900', 'como a empresa vê: cor da Millena');
  const navEmp = await txt(pg, '#nav');
  ok(!navEmp.includes('Acessos') && !navEmp.includes('Administração') && !(await pg.isVisible('#selo-admin')) && !(await pg.isVisible('#box-adm')), 'como a empresa vê: sem nada de administração');
  ok(navEmp.includes('Novo contrato') && navEmp.includes('Importar planilha'), 'como a empresa vê: funções de quem tem acesso completo');
  await pg.goto(URL + '#acessos'); await pg.waitForTimeout(200); ok((await txt(pg, '#main')).includes('Sem permissão'), 'como a empresa vê: tela de acessos bloqueada');
  await pg.goto(URL + '#painel'); await pg.waitForSelector('.kpis'); await pg.screenshot({ path: SP + '/s-como-empresa.png' });
  await pg.selectOption('#sel-emp', '__adm'); await pg.waitForTimeout(700);
  await pg.goto(URL + '#identidade'); await pg.waitForSelector('#id-termo');
  await pg.selectOption('#id-termo', 'Loja|Lojas'); await pg.fill('#id-hex', '#7c3aed'); await pg.dispatchEvent('#id-hex', 'change');
  await pg.setInputFiles('#id-logo', { name: 'logo.png', mimeType: 'image/png', buffer: Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAYAAADED76LAAAAHElEQVQoU2NkYPj/n4EIwDiqEF8YjIYC/jAYDQUAbBMIAf3Yq5kAAAAASUVORK5CYII=', 'base64') });
  await pg.waitForTimeout(400); await pg.click('#id-salvar'); await pg.waitForTimeout(400);
  ok((await txt(pg, '#nav')).includes('Lojas'), 'identidade salva: "Lojas" no menu');
  await pg.selectOption('#sel-emp', 'millena'); await pg.waitForTimeout(700);
  ok((await pg.$('#side-brand img')) !== null && (await pg.evaluate(() => document.documentElement.style.getPropertyValue('--acc'))) === '#7c3aed', 'como a empresa vê: logo e nova cor');
  await pg.selectOption('#sel-emp', '__adm'); await pg.waitForTimeout(700);
  await pg.goto(URL + '#identidade'); await pg.waitForSelector('#id-termo');
  await pg.screenshot({ path: SP + '/s-identidade.png' });
  await pg.selectOption('#id-termo', 'Filial|Filiais'); await pg.click('#id-salvar'); await pg.waitForTimeout(300);
  // naturezas do ERP: importar a planilha real (local) e revisar
  await pg.goto(URL + '#naturezas'); await pg.waitForSelector('#nt-file', { state: 'attached' });
  if (fs.existsSync(SP + '/up/nat.xlsx')) {   // planilha real: so na maquina de quem testa, nunca no repositorio
  await pg.setInputFiles('#nt-file', SP + '/up/nat.xlsx'); await pg.waitForSelector('#nt-ok');
  const rn = await txt(pg, '#nt-res'); console.log('   naturezas:', rn.replace(/\s+/g, ' ').slice(0, 260));
  ok(/1145 naturezas/.test(rn), 'importação lê as 1145 naturezas');
  await pg.click('#nt-ok'); await pg.waitForFunction(() => { const v = document.querySelector('.kpis4 .val'); return v && Number(v.textContent) > 1000; }, null, { timeout: 30000 });
  const kp = await txt(pg, '.kpis4'); console.log('   kpis:', kp.replace(/\s+/g, ' '));
  const sug = await pg.$$eval('#nt-tb tr[data-cod]', trs => trs.map(tr => tr.children[0].textContent + ' ' + tr.children[1].querySelector('.t').textContent.replace(' sugerida', '') + ' → ' + tr.querySelector('[data-cat]').selectedOptions[0].textContent));
  fs.writeFileSync(SP + '/sugestoes.txt', sug.join('\n'));
  ok(sug.length > 30, 'sugestões para revisar: ' + sug.length);
  } else console.log('   (pulado: coloque a planilha de naturezas em <pasta>/up/nat.xlsx para testar a importação real)');
  await pg.selectOption('#nt-f', 'sem'); await pg.fill('#nt-q', 'CONDOMINIO'); await pg.waitForTimeout(200);
  await pg.check('#nt-tb tr[data-cod="742"] [data-tc]'); await pg.waitForTimeout(300); await pg.selectOption('#nt-f', 'contrato'); await pg.waitForTimeout(200);
  await pg.selectOption('#nt-tb tr[data-cod="742"] [data-cat]', { label: 'Aluguel de imóvel comercial' }); await pg.waitForTimeout(400);
  await pg.goto(URL + '#novo'); await pg.waitForSelector('#c-natsel');
  ok((await pg.$$eval('#c-natsel option', os => os.map(o => o.value))).includes('742'), 'natureza revisada (742 condomínio) passa a aparecer no cadastro');
  await pg.screenshot({ path: SP + '/s-novo-nat.png' });
  await pg.goto(URL + '#naturezas'); await pg.waitForSelector('#nt-tb'); await pg.screenshot({ path: SP + '/s-naturezas.png', fullPage: false });
  // ver como usuario (dentro da visao da empresa)
  await pg.selectOption('#sel-emp', 'millena'); await pg.waitForTimeout(700);
  await pg.goto(URL + '#painel'); await pg.waitForSelector('#sel-ver option[value="u-ges"]', { state: 'attached' });
  await pg.selectOption('#sel-ver', 'u-ges'); await pg.waitForSelector('.modo-bar');
  ok((await txt(pg, '.modo-bar')).includes('Carla Mendes') && !(await txt(pg, '#nav')).includes('Acessos'), 'ver como Carla: aviso no topo e sem menu de administração');
  await pg.goto(URL + '#contratos'); await pg.waitForSelector('#tb tr');
  const nCarla = (await pg.$$('#tb tr.link')).length; ok(nCarla > 0 && nCarla < 17, 'ver como Carla: só os contratos das filiais dela (' + nCarla + ')');
  await pg.click('#tb tr.link'); await pg.waitForSelector('#b-salvar'); await pg.click('#b-salvar'); await pg.waitForTimeout(300);
  ok((await txt(pg, '#c-msg')).includes('só visualização'), 'ver como: salvar é bloqueado');
  await pg.click('[data-act="sair-sim"]'); await pg.waitForTimeout(500);
  ok(!(await pg.$('.modo-bar')) && !(await txt(pg, '#nav')).includes('Acessos'), 'volta para a empresa toda, ainda sem administração');
  await pg.selectOption('#sel-emp', '__adm'); await pg.waitForTimeout(700);
  ok((await txt(pg, '#nav')).includes('Acessos'), 'volta para a administração');
  // visao consolidada
  await pg.selectOption('#sel-adm', '*'); await pg.waitForSelector('.modo-bar.info');
  await pg.goto(URL + '#painel'); await pg.waitForSelector('[data-emp]'); const pc = await txt(pg, '#main');
  ok(pc.includes('Por empresa') && pc.includes('Novatec') && pc.includes('Millena'), 'consolidado: painel por empresa');
  ok((await txt(pg, '#side-brand')).includes('SINERGIA') && !(await pg.evaluate(() => document.documentElement.style.getPropertyValue('--acc'))), 'consolidado: identidade da Sinergia');
  ok(await (await pg.$('#box-ver')).isHidden(), 'consolidado: sem "ver como usuário"');
  await pg.goto(URL + '#contratos'); await pg.waitForSelector('#tb tr.link');
  const nTot = (await pg.$$('#tb tr.link')).length; ok(nTot === 20, 'consolidado: 20 contratos ativos (17 Millena + 3 Novatec): ' + nTot);
  await pg.selectOption('#f-fi', 'novatec'); ok((await pg.$$('#tb tr.link')).length === 3, 'consolidado: filtro por empresa');
  await pg.goto(URL + '#novo'); await pg.waitForTimeout(200); ok((await txt(pg, '#main')).includes('escolha a empresa'), 'consolidado: cadastro pede para escolher empresa');
  await pg.goto(URL + '#painel'); await pg.waitForSelector('[data-emp="novatec"]'); await pg.screenshot({ path: SP + '/s-consolidado.png', fullPage: true });
  await pg.click('[data-emp="novatec"]'); await pg.waitForTimeout(600);
  ok((await pg.inputValue('#sel-adm')) === 'novatec' && (await txt(pg, '#nav')).includes('Unidades'), 'clicar na empresa põe a Novatec em gestão ("Unidades")');
  await pg.selectOption('#sel-adm', 'millena'); await pg.waitForTimeout(600);
  // ───── GESTORA (unidades 1-6) ─────
  pg = await nova(browser, { viewport: { width: 1440, height: 900 } }, 'gestor');
  await pg.goto(URL + '#novo'); await pg.waitForSelector('#c-uni');
  const opts = await pg.$$eval('#c-uni option', os => os.map(o => o.value).filter(Boolean));
  ok(opts.join(',') === '1,2,3,4,5,6', 'gestora só vê as unidades dela no cadastro: ' + opts);
  await pg.goto(URL + '#importar'); await pg.waitForSelector('#imp-file', { state: 'attached' });
  await pg.setInputFiles('#imp-file', mk([cab, ['7', 'Fornecedor X', '', '', '1501', 'Adv loja 7', '', '', '', '', '', '', '', ''], ['77', 'Fornecedor Y', '', '', '1501', 'Adv loja 77', '', '', '', '', '', '', '', '']], 'gest.xlsx'));
  await pg.waitForSelector('#imp-res .msg'); const r3 = await txt(pg, '#imp-res');
  ok(r3.includes('não permite cadastrar na filial "7"') && r3.includes('Filial "77" não está cadastrada'), 'gestora: unidade fora do acesso e inexistente são barradas');
  ok(!(await txt(pg, '#nav')).includes('Categorias'), 'gestora não vê telas de administração');
  // ───── LEITOR (só imóveis) ─────
  pg = await nova(browser, { viewport: { width: 1440, height: 900 } }, 'leitor');
  await pg.goto(URL + '#contratos'); await pg.waitForSelector('#tb tr.link');
  ok((await pg.$$('#tb tr.link')).length === 4, 'leitor vê só os 4 de Imóveis');
  ok(!(await txt(pg, '#nav')).includes('Importar'), 'leitor não importa');
  // ───── CELULAR ─────
  const { devices } = require('playwright');
  pg = await nova(browser, { ...devices['iPhone 13'] });
  for (const h of ['painel', 'contratos', 'contrato/1', 'novo', 'importar', 'categorias', 'unidades', 'acessos', 'identidade']) {
    await pg.goto(URL + '#' + h); await pg.waitForTimeout(400);
    const w = await pg.evaluate(() => document.documentElement.scrollWidth);
    ok(w <= 390, `celular #${h}: largura ${w}`);
    if (['painel', 'contrato/1', 'importar'].includes(h)) await pg.screenshot({ path: SP + `/m-${h.replace('/', '-')}.png`, fullPage: true });
  }
  await browser.close();
  console.log('\nERROS/FALHAS:', erros.length ? erros : 'nenhum');
})();
