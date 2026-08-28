// Conexion de la app web a Supabase.
//
// La llave que va aca es la PUBLICABLE (anon / sb_publishable_...), NUNCA la
// secreta: esta pagina la descarga cualquiera que la abra. Con las politicas
// de schema_cuadre.sql esa llave solo puede LEER, nunca escribir ni borrar.
//
// Los valores estan en Supabase -> Project Settings -> API Keys,
// pestana "Publishable and secret API keys" -> Publishable key.

window.CONFIG = {
  url: "https://mdkytnimqcivgypwqqqp.supabase.co",
  anonKey: "sb_publishable_xiwomgwl3iMdk2VRA6Xl-Q_reTcM8iV"
};
