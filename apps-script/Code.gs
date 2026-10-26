/**
 * Obra Ferreira · API da página de planejamento
 * Banco de dados: planilha "Obra Ferreira · Checklist e Cronograma".
 * Arquivos enviados pela página ficam na pasta do Drive "Obra Ferreira · Arquivos".
 * Leitura é pública (para a página). Qualquer alteração exige o PIN.
 */
const SHEET_ID = '1Se7KXt4meU_Cd5eX0rutU9uennIJENLOTUag5QghvbE';
const PIN = '1234';
const FIRST_ROW = 5;
const PASTA_RAIZ = 'Obra Ferreira · Arquivos';

const TABELAS = {
  checklist:  { aba: 'Checklist',  cols: ['num','disciplina','item','responsavel','fase','status','pedido','obs'], texto: ['pedido'] },
  cronograma: { aba: 'Cronograma', cols: ['num','etapa','responsavel','inicio','fim','dias','status','obs'], datas: ['inicio','fim'], formula: 'dias' },
  medicoes:   { aba: 'Medições',   cols: ['num','data','servico','executor','parcela','valor','status','obs'], datas: ['data'], numeros: ['valor'] },
  pagamentos: { aba: 'Pagamentos', cols: ['num','data','favorecido','referencia','valor','forma','status','obs'], datas: ['data'], numeros: ['valor'] },
  contratos:  { aba: 'Contratos',  cols: ['num','contratado','objeto','valor','assinatura','prazo','status','obs'], datas: ['assinatura','prazo'], numeros: ['valor'] },
  arquivos:   { aba: 'Arquivos',   cols: ['num','secao','pasta','nome','link','fileId','enviado','obs'], datas: ['enviado'] },
  pastas:     { aba: 'Pastas',     cols: ['num','secao','caminho','criada'], datas: ['criada'] },
};

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.arquivo) return json(baixar(String(p.arquivo)));
  const out = { ok: true, atualizado: new Date().toISOString() };
  Object.keys(TABELAS).forEach(k => out[k] = ler(TABELAS[k]));
  return json(out);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = JSON.parse(e.postData.contents || '{}');
    if (String(req.pin) !== PIN) return json({ ok: false, erro: 'PIN incorreto.' });
    if (req.acao === 'upload') return json(upload(req.dados || {}));
    const t = TABELAS[req.tabela];
    if (!t) return json({ ok: false, erro: 'Tabela desconhecida.' });
    const sh = aba(t);
    if (req.acao === 'adicionar') {
      const num = proximo(t);
      gravarLinha(sh, t, primeiraLivre(sh), Object.assign({}, req.dados, { num: num }));
      return json({ ok: true, num: num });
    }
    const row = acharLinha(sh, req.num);
    if (!row) return json({ ok: false, erro: 'Registro não encontrado. Toque em Atualizar.' });
    if (req.acao === 'atualizar') { gravarLinha(sh, t, row, req.dados || {}, true); return json({ ok: true }); }
    if (req.acao === 'excluir') {
      if (req.tabela === 'arquivos') {
        const id = sh.getRange(row, 6).getValue();
        try { if (id) DriveApp.getFileById(id).setTrashed(true); } catch (err) {}
      }
      sh.deleteRow(row);
      return json({ ok: true });
    }
    return json({ ok: false, erro: 'Ação desconhecida.' });
  } catch (err) {
    return json({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- planilha ---------- */
function aba(t) { return SpreadsheetApp.openById(SHEET_ID).getSheetByName(t.aba); }

function ler(t) {
  const sh = aba(t);
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < FIRST_ROW) return [];
  const n = t.cols.length;
  const vals = sh.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, n).getValues();
  const tz = Session.getScriptTimeZone();
  return vals
    .filter(r => String(r[0]).trim() !== '' && r.slice(1).some(v => String(v).trim() !== ''))
    .map(r => {
      const o = {};
      t.cols.forEach((c, i) => {
        let v = r[i];
        if (v instanceof Date) v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
        o[c] = v;
      });
      return o;
    });
}

function proximo(t) { return ler(t).reduce((m, r) => Math.max(m, Number(r.num) || 0), 0) + 1; }

function primeiraLivre(sh) {
  const last = Math.max(sh.getLastRow(), FIRST_ROW - 1);
  if (last < FIRST_ROW) return FIRST_ROW;
  const col = sh.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, 2).getValues();
  for (let i = 0; i < col.length; i++) if (String(col[i][0]).trim() === '' && String(col[i][1]).trim() === '') return FIRST_ROW + i;
  return last + 1;
}

function acharLinha(sh, num) {
  const last = sh.getLastRow();
  if (last < FIRST_ROW) return 0;
  const nums = sh.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, 1).getValues();
  for (let i = 0; i < nums.length; i++) if (String(nums[i][0]) === String(num)) return FIRST_ROW + i;
  return 0;
}

function gravarLinha(sh, t, row, dados, parcial) {
  t.cols.forEach((c, i) => {
    const cell = sh.getRange(row, i + 1);
    if (c === t.formula) {
      if (!parcial) cell.setFormula('=IF(OR(D' + row + '="",E' + row + '=""),"",E' + row + '-D' + row + '+1)');
      return;
    }
    if (parcial && !(c in dados)) return;
    let v = dados[c] == null ? '' : dados[c];
    if ((t.datas || []).indexOf(c) >= 0) {
      cell.setNumberFormat('dd/mm/yyyy');
      if (v) { const p = String(v).split('-'); v = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])); }
    } else if ((t.numeros || []).indexOf(c) >= 0) {
      v = v === '' ? '' : Number(String(v).replace(',', '.'));
    } else if ((t.texto || []).indexOf(c) >= 0 || c === 'caminho' || c === 'pasta') {
      cell.setNumberFormat('@');
    }
    cell.setValue(v);
  });
}

/* ---------- arquivos no Drive ---------- */
function pastaDrive(caminho) {
  let it = DriveApp.getFoldersByName(PASTA_RAIZ);
  let f = it.hasNext() ? it.next() : DriveApp.createFolder(PASTA_RAIZ);
  String(caminho).split('/').map(s => s.trim()).filter(Boolean).forEach(nome => {
    const sub = f.getFoldersByName(nome);
    f = sub.hasNext() ? sub.next() : f.createFolder(nome);
  });
  return f;
}

function upload(d) {
  if (!d.base64 || !d.nome) return { ok: false, erro: 'Arquivo vazio.' };
  const bytes = Utilities.base64Decode(d.base64);
  const blob = Utilities.newBlob(bytes, d.mime || 'application/octet-stream', d.nome);
  const pasta = pastaDrive((d.secao || 'Documentos') + '/' + (d.pasta || ''));
  const file = pasta.createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) {}
  const t = TABELAS.arquivos, sh = aba(t), num = proximo(t);
  const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  gravarLinha(sh, t, primeiraLivre(sh), {
    num: num, secao: d.secao || 'Documentos', pasta: d.pasta || '', nome: d.nome,
    link: file.getUrl(), fileId: file.getId(), enviado: hoje, obs: d.obs || ''
  });
  return { ok: true, num: num };
}

/* Só entrega arquivos que estão listados na aba Arquivos. */
function baixar(id) {
  const lista = ler(TABELAS.arquivos);
  const reg = lista.filter(r => String(r.fileId) === id)[0];
  if (!reg) return { ok: false, erro: 'Arquivo não encontrado.' };
  const blob = DriveApp.getFileById(id).getBlob();
  return { ok: true, nome: reg.nome, mime: blob.getContentType(), base64: Utilities.base64Encode(blob.getBytes()) };
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
