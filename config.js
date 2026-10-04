// Configuración de la app. Solo tienes que cambiar CLIENT_ID (ver LEEME.md, paso 1).
window.CONFIG = {
  // "Id. de aplicación (cliente)" que te da Microsoft Entra al registrar la app
  CLIENT_ID: "8b0c61bf-c214-4325-9290-ed7bdf3b83cc",

  // Ruta del archivo dentro de tu OneDrive (sin "C:\Users\...\OneDrive")
  RUTA_ARCHIVO: "Calender/Inversiones/Presupuesto Personal.xlsx",

  // Listas por defecto (la app usa las de la hoja "Listas" del Excel cuando puede leerlas)
  GASTOS: ["Salud", "Mercado", "Jorge", "Carro", "Salidas", "Compras", "Obligaciones", "Inversion"],
  INGRESOS: ["Salario", "Honorarios", "Rendimientos", "Otros ingresos"],
  MEDIOS: ["Efectivo", "Débito", "Crédito"],

  GRAPH: "https://graph.microsoft.com/v1.0",
  LOGIN: "https://login.microsoftonline.com/consumers/oauth2/v2.0",
};
