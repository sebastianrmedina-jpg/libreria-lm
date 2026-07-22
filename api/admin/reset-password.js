const { requireAdmin } = require("../_lib/requireAdmin");

module.exports = async (req, res) => {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Método no permitido." });
    return;
  }

  const { error: authError, adminClient } = await requireAdmin(req);
  if (authError) {
    res.status(authError.status).json({ error: authError.message });
    return;
  }

  const { userId, newPassword } = req.body || {};
  if (!userId || !newPassword) {
    res.status(400).json({ error: "Falta el usuario o la contraseña nueva." });
    return;
  }

  const { error } = await adminClient.auth.admin.updateUserById(userId, { password: newPassword });
  if (error) {
    res.status(400).json({ error: `No se pudo resetear la contraseña: ${error.message}` });
    return;
  }

  res.status(200).json({ ok: true });
};
