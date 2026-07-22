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

  const { userId } = req.body || {};
  if (!userId) {
    res.status(400).json({ error: "Falta el usuario a borrar." });
    return;
  }

  // No dejar el sistema sin ningún admin, aunque el cliente ya valide esto antes de llamar acá.
  const { data: target } = await adminClient.from("lm_users").select("role").eq("id", userId).maybeSingle();
  if (target?.role === "admin") {
    const { count } = await adminClient
      .from("lm_users")
      .select("id", { count: "exact", head: true })
      .eq("role", "admin");
    if ((count || 0) <= 1) {
      res.status(400).json({ error: "Debe haber al menos un administrador." });
      return;
    }
  }

  const { error: profileError } = await adminClient.from("lm_users").delete().eq("id", userId);
  if (profileError) {
    res.status(400).json({ error: `No se pudo borrar el perfil: ${profileError.message}` });
    return;
  }

  const { error: authDeleteError } = await adminClient.auth.admin.deleteUser(userId);
  if (authDeleteError) {
    res.status(400).json({ error: `Se borró el perfil pero no el login: ${authDeleteError.message}` });
    return;
  }

  res.status(200).json({ ok: true });
};
