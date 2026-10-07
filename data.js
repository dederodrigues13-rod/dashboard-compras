const { parseWorkbook } = require("../lib/parse");

// Converte o link de compartilhamento do OneDrive pessoal em link de download direto
function shareToDownloadUrl(link) {
  const b64 = Buffer.from(link.trim()).toString("base64").replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
  return `https://api.onedrive.com/v1.0/shares/u!${b64}/root/content`;
}

module.exports = async (req, res) => {
  try {
    const link = process.env.ONEDRIVE_LINK;
    if (!link) {
      return res.status(500).json({ error: "Variável ONEDRIVE_LINK não configurada no Vercel." });
    }
    const r = await fetch(shareToDownloadUrl(link), { redirect: "follow" });
    if (!r.ok) {
      return res.status(502).json({
        error: `OneDrive respondeu ${r.status}. Confira se o link está como 'Qualquer pessoa com o link pode ver'.`,
      });
    }
    const buf = Buffer.from(await r.arrayBuffer());
    const parsed = parseWorkbook(buf, process.env.SHEET_NAME);
    res.setHeader("Cache-Control", "s-maxage=60, stale-while-revalidate=120");
    res.status(200).json({ updatedAt: new Date().toISOString(), ...parsed });
  } catch (e) {
    res.status(500).json({ error: String(e.message || e) });
  }
};
