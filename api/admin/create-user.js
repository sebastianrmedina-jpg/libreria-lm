const { requireAdmin } = require("../_lib/requireAdmin");

const EMAIL_DOMAIN = "libreria-lm.local";

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

  const body = req.body || {};
  const username = (body.username || "").trim();
  const password = body.password || "";
  const name = (body.name || "").trim();
  if (!username || !password || !name) {
    res.status(400).json({ error: "Completá usuario, contraseña y nombre." });
    return;
  }

  const email = `${username}@${EMAIL_DOMAIN}`;

  const { data: created, error: createError } = await adminClient.auth.admin.createUser({
    email,
    password,
    email_confirm: true,
    user_metadata: { username },
  });
  if (createError) {
    res.status(400).json({ error: `No se pudo crear el usuario: ${createError.message}` });
    return;
  }

  const uid = created.user.id;
  const { error: profileError } = await adminClient.from("lm_users").insert({
    id: uid,
    username,
    password: "",
    name,
    role: body.role || "vendedor",
    email: body.email || "",
    vendedor: body.vendedor || "",
    price_list: body.priceList || "default",
    can_see_all: body.canSeeAll !== false,
    phone: body.phone || "",
    cargo: body.cargo || "",
    avatar: body.avatar || "",
    barcode_enabled: body.barcodeEnabled || false,
    tabs_deshabilitados: body.tabsDeshabilitados || [],
  });
  if (profileError) {
    // El login ya existe en auth.users pero el perfil falló — deshacemos para no dejar una cuenta fantasma.
    await adminClient.auth.admin.deleteUser(uid);
    res.status(400).json({ error: `No se pudo guardar el perfil: ${profileError.message}` });
    return;
  }

  res.status(200).json({ id: uid, username, email });
};
