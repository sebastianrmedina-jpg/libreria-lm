const { createClient } = require("@supabase/supabase-js");

// Proxy del envío real a OneSignal. Hace falta un endpoint de servidor para
// esto a propósito: la API REST de OneSignal (onesignal.com/api/v1/notifications)
// bloquea las llamadas hechas directo desde un navegador (CORS) — está
// pensada para que la llame un servidor, nunca el cliente. Esto es lo que
// hacía que las notificaciones nunca llegaran, sin importar qué tan bien
// estuviera configurado el permiso o el service worker. De paso, la clave de
// OneSignal deja de viajar en el bundle público de la app (antes estaba
// hardcodeada ahí, visible para cualquiera).
module.exports = async (req, res) => {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Método no permitido." });
    return;
  }

  const authHeader = req.headers.authorization || "";
  const token = authHeader.startsWith("Bearer ") ? authHeader.slice(7) : null;
  if (!token) {
    res.status(401).json({ error: "Falta el token de autorización." });
    return;
  }

  const supabaseUrl = process.env.SUPABASE_URL;
  const publishableKey = process.env.SUPABASE_PUBLISHABLE_KEY;
  if (!supabaseUrl || !publishableKey) {
    res.status(500).json({ error: "Faltan variables de entorno de Supabase." });
    return;
  }

  // Cualquier usuario logueado puede disparar una notificación (un vendedor
  // avisándole a admin de una venta, admin confirmando un pago, etc.) — a
  // diferencia de api/admin/*, acá no hace falta ser admin, solo tener una
  // sesión real.
  const callerClient = createClient(supabaseUrl, publishableKey, {
    global: { headers: { Authorization: `Bearer ${token}` } },
  });
  const { data: userData, error: userError } = await callerClient.auth.getUser();
  if (userError || !userData?.user) {
    res.status(401).json({ error: "Sesión inválida o vencida." });
    return;
  }

  const { title, body, targetUsername } = req.body || {};
  if (!title || !body || !targetUsername) {
    res.status(400).json({ error: "Faltan datos del envío (title, body, targetUsername)." });
    return;
  }

  const appId = process.env.ONESIGNAL_APP_ID;
  const apiKey = process.env.ONESIGNAL_REST_API_KEY;
  if (!appId || !apiKey) {
    res.status(500).json({ error: "Faltan variables de entorno de OneSignal." });
    return;
  }

  // Si el destinatario es "admin", enviar a todos con rol admin.
  // Si es un username específico, enviar solo a ese usuario.
  const filters = targetUsername === "admin"
    ? [{ field: "tag", key: "role", relation: "=", value: "admin" }]
    : [{ field: "tag", key: "username", relation: "=", value: targetUsername }];

  try {
    const osRes = await fetch("https://onesignal.com/api/v1/notifications", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Key ${apiKey}`,
      },
      body: JSON.stringify({
        app_id: appId,
        headings: { en: title, es: title },
        contents: { en: body, es: body },
        filters,
        // Sin esto Android puede demorar o no mostrar el push bajo ahorro de batería.
        priority: 10,
      }),
    });
    const json = await osRes.json().catch(() => null);
    if (!osRes.ok) {
      res.status(502).json({ error: "OneSignal rechazó el envío.", detail: json });
      return;
    }
    res.status(200).json({ ok: true, recipients: json?.recipients ?? 0 });
  } catch (e) {
    res.status(502).json({ error: "No se pudo contactar a OneSignal." });
  }
};
