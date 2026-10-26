/**
 * Obra Ferreira · API da página de planejamento
 * Lê e grava o checklist e o cronograma na planilha "Obra Ferreira · Checklist e Cronograma".
 * Leitura é pública (para a página). Gravação exige o PIN.
 */
const SHEET_ID = '1Se7KXt4meU_Cd5eX0rutU9uennIJENLOTUag5QghvbE';
const PIN = '1234';
const FIRST_ROW = 5;

const CK = { aba: 'Checklist', cols: ['num', 'disciplina', 'item', 'responsavel', 'fase', 'status', 'pedido', 'obs'] };
const CR = { aba: 'Cronograma', cols: ['num', 'etapa', 'responsavel', 'inicio', 'fim', 'dias', 'status', 'obs'] };

function doGet() {
  return json({ ok: true, checklist: ler(CK), cronograma: ler(CR), atualizado: new Date().toISOString() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const req = JSON.parse(e.postData.contents || '{}');
    if (String(req.pin) !== PIN) return json({ ok: false, erro: 'PIN incorreto.' });
    const t = req.tabela === 'cronograma' ? CR : CK;
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(t.aba);
    if (req.acao === 'adicionar') {
      const linhas = ler(t);
      const num = linhas.reduce((m, r) => Math.max(m, Number(r.num) || 0), 0) + 1;
      const row = FIRST_ROW + linhas.length;
      gravarLinha(sh, t, row, Object.assign({}, req.dados, { num: num }));
      return json({ ok: true, num: num });
    }
    const row = acharLinha(sh, req.num);
    if (!row) return json({ ok: false, erro: 'Item não encontrado. Recarregue a página.' });
    if (req.acao === 'atualizar') {
      gravarLinha(sh, t, row, req.dados, true);
      return json({ ok: true });
    }
    if (req.acao === 'excluir') {
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

function ler(t) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(t.aba);
  const last = sh.getLastRow();
  if (last < FIRST_ROW) return [];
  const vals = sh.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, 8).getValues();
  const tz = Session.getScriptTimeZone();
  return vals
    .filter(r => String(r[1]).trim() !== '' && String(r[0]).trim() !== '')
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

function acharLinha(sh, num) {
  const last = sh.getLastRow();
  if (last < FIRST_ROW) return 0;
  const nums = sh.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, 1).getValues();
  for (let i = 0; i < nums.length; i++) if (String(nums[i][0]) === String(num)) return FIRST_ROW + i;
  return 0;
}

function gravarLinha(sh, t, row, dados, parcial) {
  t.cols.forEach((c, i) => {
    if (c === 'dias') {
      if (!parcial) sh.getRange(row, i + 1).setFormula('=IF(OR(D' + row + '="",E' + row + '=""),"",E' + row + '-D' + row + '+1)');
      return;
    }
    if (parcial && !(c in dados)) return;
    let v = dados[c] == null ? '' : dados[c];
    const cell = sh.getRange(row, i + 1);
    if ((c === 'inicio' || c === 'fim') && v) {
      const p = String(v).split('-');
      v = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
      cell.setNumberFormat('dd/mm/yyyy');
    }
    if (c === 'pedido') cell.setNumberFormat('@');
    cell.setValue(v);
  });
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
