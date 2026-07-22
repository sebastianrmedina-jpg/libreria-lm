const { createClient } = require("@supabase/supabase-js");

const EMAIL_DOMAIN = "libreria-lm.local";
const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

// De un solo uso: crea una cuenta real de Supabase Auth para cada fila de
// lm_users que todavía no la tenga, y actualiza el id de esa fila para que
// coincida con el uid nuevo (así auth.uid() la encuentra directo).
//
// No usa requireAdmin — todavía no existe ninguna sesión real de Supabase
// Auth la primera vez que esto se corre (es precisamente lo que este script
// crea). Se protege con un secreto de un solo propósito, MIGRATION_SECRET,
// que se borra de las variables de entorno de Vercel apenas se termina de usar.
//
// Idempotente: una fila con id que YA tiene forma de UUID se asume migrada y
// se saltea, así que correr esto dos veces no rompe ni duplica nada.
module.exports = async (req, res) => {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Método no permitido." });
    return;
  }

  const secretHeader = req.headers["x-migration-secret"];
  if (!process.env.MIGRATION_SECRET || secretHeader !== process.env.MIGRATION_SECRET) {
    res.status(401).json({ error: "Secreto de migración inválido." });
    return;
  }

  const url = process.env.SUPABASE_URL;
  const secretKey = process.env.SUPABASE_SECRET_KEY;
  if (!url || !secretKey) {
    res.status(500).json({ error: "Faltan variables de entorno de Supabase." });
    return;
  }
  const adminClient = createClient(url, secretKey);

  const { data: users, error: usersError } = await adminClient.from("lm_users").select("*");
  if (usersError) {
    res.status(500).json({ error: `No se pudo leer lm_users: ${usersError.message}` });
    return;
  }

  const results = [];
  for (const u of users) {
    if (UUID_RE.test(u.id)) {
      results.push({ username: u.username, status: "ya-migrado" });
      continue;
    }

    const email = `${u.username}@${EMAIL_DOMAIN}`;
    const { data: created, error: createError } = await adminClient.auth.admin.createUser({
      email,
      password: u.password || Math.random().toString(36).slice(2, 12),
      email_confirm: true,
      user_metadata: { username: u.username },
    });
    if (createError) {
      results.push({ username: u.username, status: "error", detail: createError.message });
      continue;
    }

    const newId = created.user.id;
    const { error: updateError } = await adminClient.from("lm_users").update({ id: newId }).eq("id", u.id);
    if (updateError) {
      results.push({ username: u.username, status: "error", detail: `usuario creado pero no se pudo actualizar el id: ${updateError.message}` });
      continue;
    }

    results.push({ username: u.username, status: "migrado", newId });
  }

  res.status(200).json({ results });
};
