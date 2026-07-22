const { createClient } = require("@supabase/supabase-js");

// No es un endpoint — helper que usan los archivos bajo api/admin/*.
// Valida el JWT del que llama (con la clave pública, nunca la secreta) y
// confirma que su fila en lm_users tiene role="admin" antes de dar acceso al
// cliente con la clave secreta, que es el único que puede hacer operaciones
// privilegiadas (auth.admin.*).
async function requireAdmin(req) {
  const authHeader = req.headers.authorization || "";
  const token = authHeader.startsWith("Bearer ") ? authHeader.slice(7) : null;
  if (!token) {
    return { error: { status: 401, message: "Falta el token de autorización." } };
  }

  const url = process.env.SUPABASE_URL;
  const publishableKey = process.env.SUPABASE_PUBLISHABLE_KEY;
  const secretKey = process.env.SUPABASE_SECRET_KEY;
  if (!url || !publishableKey || !secretKey) {
    return { error: { status: 500, message: "Faltan variables de entorno de Supabase." } };
  }

  const callerClient = createClient(url, publishableKey, {
    global: { headers: { Authorization: `Bearer ${token}` } },
  });

  const { data: userData, error: userError } = await callerClient.auth.getUser();
  if (userError || !userData?.user) {
    return { error: { status: 401, message: "Sesión inválida o vencida." } };
  }

  const { data: profile, error: profileError } = await callerClient
    .from("lm_users")
    .select("role")
    .eq("id", userData.user.id)
    .maybeSingle();
  if (profileError || !profile || profile.role !== "admin") {
    return { error: { status: 403, message: "Necesitás ser administrador para hacer esto." } };
  }

  const adminClient = createClient(url, secretKey);
  return { user: userData.user, adminClient };
}

module.exports = { requireAdmin };
