const { chromium, devices } = require('playwright');
const fs = require('fs');
const SP = process.argv[2];
const LIB = fs.readFileSync(SP + '/xlsxpkg/package/dist/xlsx.full.min.js', 'utf8');
const URL = 'file://' + SP + '/gestao-contratos.html';
const erros = [];
const ok = (c, m) => { console.log((c ? 'OK   ' : 'FALHA ') + m); if (!c) erros.push(m); };
const brlNum = s => Number(String(s).replace(/[^\d,]/g, '').replace(',', '.'));
const ymd = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
(async () => {
  const b = await chromium.launch();
  const ctx = await b.newContext({ viewport: { width: 1440, height: 900 } });
  await ctx.route(/jsdelivr/, r => r.fulfill({ body: LIB, contentType: 'application/javascript' }));
  await ctx.route(/fonts\./, r => r.fulfill({ body: '', contentType: 'text/css' }));
  const pg = await ctx.newPage(); pg.on('pageerror', e => erros.push('JS: ' + e.message));
  await pg.goto(URL); await pg.click('#f-login button[type=submit]'); await pg.waitForSelector('.kpis');
  // painel: reajustes
  const painel = await pg.textContent('#main');
  ok(painel.includes('Reajustes') && painel.includes('Falta registrar'), 'painel: cartão Reajustes com "Falta registrar"');
  // contrato 1: ja teve reajuste; proximo em ~30 dias
  await pg.goto(URL + '#contrato/1'); await pg.waitForSelector('#rj');
  const rj = await pg.textContent('#rj');
  ok(rj.includes('IGP-M') && rj.includes('Último reajuste') && /4,20%/.test(rj), 'contrato: regra e último reajuste (IGP-M 4,20%)');
  ok((await pg.textContent('#tlv')).includes('Reajuste IGP-M 4,20%'), 'linha do tempo mostra o reajuste anterior');
  ok((await pg.textContent('.detail')).includes('Próximo reajuste (IGP-M):'), 'resumo com o próximo reajuste pela mesma regra');
  // data-base futura: indice ainda nao divulgado
  await pg.click('#rj-calcular'); await pg.waitForSelector('#rj-res .msg, #rj-res .rj-res');
  ok((await pg.textContent('#rj-res')).includes('ainda não foi divulgado'), 'data-base futura: avisa que o índice ainda não saiu');
  // data-base no passado: calcula
  const d = new Date(); d.setMonth(d.getMonth() - 1); d.setDate(1);
  await pg.fill('#rj-data', ymd(d)); await pg.click('#rj-calcular'); await pg.waitForSelector('#rj-res .rj-res');
  const res = await pg.textContent('#rj-res');
  const atual = brlNum(await pg.textContent('.rj-res >> div:has-text("Valor atual") >> b')), novo = brlNum(await pg.textContent('.rj-res .novo b'));
  const aplicado = Number((await pg.textContent('.rj-res >> div:has-text("Percentual aplicado") >> b')).replace('%', '').replace(',', '.'));
  ok(res.includes('Período do índice') && Math.abs(novo - Math.round(atual * (1 + aplicado / 100) * 100) / 100) < 0.011, `cálculo: ${atual} × (1 + ${aplicado}%) = ${novo}`);
  ok((await pg.$$('.rj-meses b')).length === 12, 'mostra os 12 meses do índice');
  ok(res.includes('Atenção: pela regra do contrato'), 'data-base diferente do aniversário: avisa antes de registrar');
  await pg.click('#rj-registrar'); await pg.click('#m-ok'); await pg.waitForTimeout(500);
  ok(brlNum(await pg.inputValue('#c-valor')) === novo, 'depois de registrar, o valor do contrato é o novo');
  const tl = await pg.textContent('#tlv'); ok((tl.match(/Reajuste IGP-M/g) || []).length === 2, 'linha do tempo com os 2 reajustes');
  ok(await pg.waitForFunction(() => /reajuste registrado/.test(document.querySelector('#hist').textContent), null, { timeout: 5000 }).then(() => true, () => false), 'histórico: reajuste registrado');
  await pg.screenshot({ path: SP + '/rj-contrato.png', fullPage: true });
  // teto: contrato 3 (CD) com teto de 0,5%
  await pg.goto(URL + '#contrato/3'); await pg.waitForSelector('#c-teto');
  ok(!(await pg.isHidden('#c-teto-f')) && await pg.isHidden('#c-pct-f'), 'IGP-M: mostra teto/negativo/substituto e esconde percentual fixo');
  await pg.fill('#c-teto', '0,5'); await pg.click('#b-salvar'); await pg.waitForTimeout(400);
  ok((await pg.textContent('#rj-pill')).includes('Devido desde'), 'contrato sem reajuste registrado: "Devido desde"');
  await pg.click('#rj-calcular'); await pg.waitForSelector('#rj-res .rj-res, #rj-res .msg');
  const rt = await pg.textContent('#rj-res');
  ok(/Percentual aplicado\s*0,5000%/.test(rt) || /índice negativo/.test(rt), 'teto de 0,5% aplicado (ou índice negativo mantido)');
  // aditivo de prazo
  await pg.click('#b-aditivo'); await pg.selectOption('#ad-tipo', 'aditivo_prazo'); await pg.fill('#ad-fim', '2030-12-31'); await pg.fill('#ad-obs', 'Prorrogação por 4 anos'); await pg.click('#m-ok'); await pg.waitForTimeout(500);
  ok((await pg.inputValue('#c-fim')) === '2030-12-31' && (await pg.textContent('#tlv')).includes('Aditivo de prazo'), 'aditivo de prazo muda o vencimento e entra na linha do tempo');
  // percentual fixo: contrato 10 (exames)
  await pg.goto(URL + '#contrato/10'); await pg.waitForSelector('#c-reaj');
  await pg.selectOption('#c-reaj', 'Percentual fixo'); ok(!(await pg.isHidden('#c-pct-f')), 'percentual fixo: mostra o campo do percentual');
  await pg.click('#b-salvar'); await pg.waitForTimeout(200); ok((await pg.textContent('#c-msg')).includes('percentual fixo'), 'percentual fixo sem número: avisa');
  await pg.fill('#c-pct', '5'); await pg.click('#b-salvar'); await pg.waitForTimeout(400);
  await pg.fill('#rj-data', ymd(d)); await pg.click('#rj-calcular'); await pg.waitForSelector('#rj-res .rj-res');
  const v10 = brlNum(await pg.textContent('.rj-res >> div:has-text("Valor atual") >> b')), n10 = brlNum(await pg.textContent('.rj-res .novo b'));
  ok(Math.abs(n10 - Math.round(v10 * 1.05 * 100) / 100) < 0.011, `percentual fixo 5%: ${v10} → ${n10}`);
  // convencao coletiva: contrato 5 (vigilancia), percentual digitado
  await pg.goto(URL + '#contrato/5'); await pg.waitForSelector('#rj-calcular');
  await pg.click('#rj-calcular'); await pg.waitForSelector('#rj-res .msg'); ok((await pg.textContent('#rj-res')).includes('Informe o percentual'), 'convenção coletiva: pede o percentual');
  await pg.fill('#rj-man', '6,5'); await pg.click('#rj-calcular'); await pg.waitForSelector('#rj-res .rj-res');
  const v5 = brlNum(await pg.textContent('.rj-res >> div:has-text("Valor atual") >> b')), n5 = brlNum(await pg.textContent('.rj-res .novo b'));
  ok(Math.abs(n5 - Math.round(v5 * 1.065 * 100) / 100) < 0.011, `convenção coletiva 6,5%: ${v5} → ${n5}`);
  // lista: filtro reajuste devido
  await pg.goto(URL + '#contratos?st=reajuste'); await pg.waitForSelector('#tb tr');
  const nr = (await pg.$$('#tb tr.link')).length; ok(nr >= 1, 'lista: filtro "Reajuste devido" (' + nr + ')');
  // celular
  const m = await b.newContext({ ...devices['iPhone 13'] });
  await m.route(/jsdelivr/, r => r.fulfill({ body: LIB, contentType: 'application/javascript' })); await m.route(/fonts\./, r => r.fulfill({ body: '' }));
  const mp = await m.newPage(); mp.on('pageerror', e => erros.push('JS: ' + e.message));
  await mp.goto(URL); await mp.click('#f-login button[type=submit]'); await mp.waitForSelector('.kpis');
  await mp.goto(URL + '#contrato/1'); await mp.waitForSelector('#rj-calcular');
  await mp.fill('#rj-data', ymd(d)); await mp.click('#rj-calcular'); await mp.waitForSelector('#rj-res .rj-res');
  ok((await mp.evaluate(() => document.documentElement.scrollWidth)) <= 390, 'celular: cálculo do reajuste sem passar da tela');
  await mp.screenshot({ path: SP + '/rj-iphone.png', fullPage: true });
  await b.close();
  console.log('\nERROS/FALHAS:', erros.length ? erros : 'nenhum');
})();
