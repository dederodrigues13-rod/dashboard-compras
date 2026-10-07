const XLSX = require("xlsx");

const norm = s =>
  String(s ?? "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");

// nomes aceitos para cada coluna (já normalizados: sem acento, sem símbolos)
const ALIASES = {
  rc: ["nchamado", "nochamado", "numerochamado", "chamado", "rc", "nrc", "numerorc"],
  item: ["descricaodoitemservico", "descricaodoitem", "descricao", "item"],
  setor: ["maquinasetor", "maquina", "setor"],
  supplier: ["fornecedor"],
  qty: ["qtde", "qtd", "quantidade"],
  unit: ["valorunitario", "vlunitario"],
  total: ["valortotal", "vltotal", "valor"],
  date: ["dataaberturarequisicao", "dataabertura", "dataderequisicao", "data"],
  status: ["status"],
};

function pad(n) {
  return String(n).padStart(2, "0");
}

function toDateStr(v) {
  if (v == null || v === "") return "";
  if (typeof v === "number") {
    const p = XLSX.SSF.parse_date_code(v);
    if (!p) return "";
    return `${pad(p.d)}/${pad(p.m)}/${p.y}`;
  }
  if (v instanceof Date) {
    return `${pad(v.getUTCDate())}/${pad(v.getUTCMonth() + 1)}/${v.getUTCFullYear()}`;
  }
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) {
    const y = m[3].length === 2 ? "20" + m[3] : m[3];
    return `${pad(m[1])}/${pad(m[2])}/${y}`;
  }
  return "";
}

function toNumber(v) {
  if (v == null || v === "") return null;
  if (typeof v === "number") return v;
  let s = String(v).replace(/[R$\s]/g, "");
  if (s.includes(",")) s = s.replace(/\./g, "").replace(",", ".");
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function parseWorkbook(buffer, sheetName) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  let name = sheetName && wb.SheetNames.includes(sheetName) ? sheetName : null;

  // se não foi informada a aba, procura a primeira que tenha a coluna de chamado
  const tryParse = n => {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: "" });
    for (let i = 0; i < Math.min(rows.length, 20); i++) {
      const cells = rows[i].map(norm);
      const hasRc = cells.some(c => ALIASES.rc.includes(c));
      const hasStatus = cells.some(c => ALIASES.status.includes(c));
      if (hasRc && hasStatus) return { rows, headerIdx: i, cells };
    }
    return null;
  };

  let found = null;
  if (name) found = tryParse(name);
  if (!found) {
    for (const n of wb.SheetNames) {
      const r = tryParse(n);
      if (r) { found = r; name = n; break; }
    }
  }
  if (!found) {
    throw new Error(
      "Não encontrei a linha de cabeçalho (colunas 'Nº Chamado' e 'Status'). Abas lidas: " +
        wb.SheetNames.join(", ")
    );
  }

  const { rows, headerIdx, cells } = found;
  const col = {};
  for (const [key, list] of Object.entries(ALIASES)) {
    const idx = cells.findIndex(c => list.includes(c));
    if (idx >= 0) col[key] = idx;
  }

  const out = [];
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    const rc = String(r[col.rc] ?? "").trim();
    if (!rc || !/\d/.test(rc)) continue;
    const qty = toNumber(r[col.qty]);
    let value = toNumber(r[col.total]);
    if (value == null && col.unit != null) {
      const u = toNumber(r[col.unit]);
      if (u != null) value = qty ? u * qty : u;
    }
    const date = toDateStr(r[col.date]);
    out.push({
      rc,
      item: String(r[col.item] ?? "").trim(),
      setor: col.setor != null ? String(r[col.setor] ?? "").trim() : "",
      supplier: String(r[col.supplier] ?? "").trim(),
      qty,
      value,
      date,
      status: String(r[col.status] ?? "").trim().toUpperCase(),
    });
  }
  return { sheet: name, columns: Object.keys(col), rows: out };
}

module.exports = { parseWorkbook };
