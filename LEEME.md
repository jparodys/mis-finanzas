# Mis Finanzas – app para el celular

Registra gastos e ingresos desde el celular directo en **Presupuesto Parodys.xlsx** (hoja *Registro*) en tu OneDrive.
La app descarga el archivo, escribe la nueva fila en Fecha / Descripción / Movimiento / Tipo de gasto / Medio de pago / Valor / Notas y lo vuelve a subir.
Tus fórmulas, gráficos, listas desplegables y formatos no se tocan; Excel recalcula todo la próxima vez que abras el archivo.

Configuración única: unos 20–30 minutos.

---

## Paso 1 – Registrar la app en Microsoft (para poder iniciar sesión)

1. Entra a **https://entra.microsoft.com** con la **misma cuenta de Microsoft de tu OneDrive**.
   - Si te pide crear un directorio o una *cuenta gratuita de Azure*, créala. Puede pedir una tarjeta solo para verificar identidad; registrar apps no tiene costo.
2. Menú **Aplicaciones → Registros de aplicaciones → + Nuevo registro**.
   - **Nombre:** `Mis Finanzas`
   - **Tipos de cuenta admitidos:** *Solo cuentas personales de Microsoft*
   - **URI de redirección:** plataforma **Aplicación de página única (SPA)** y la dirección
     `https://TU-USUARIO.github.io/mis-finanzas/` (la defines en el paso 2; puedes volver a editarla después en *Autenticación*).
   - Clic en **Registrar**.
3. Copia el **Id. de aplicación (cliente)** (algo como `1a2b3c4d-....`).
4. En **Permisos de API → + Agregar un permiso → Microsoft Graph → Permisos delegados**, marca **Files.ReadWrite** → *Agregar permisos*.
5. Abre `config.js` y pega el Id:
   ```js
   CLIENT_ID: "1a2b3c4d-....",
   ```
   Si moviste el Excel de carpeta, cambia también `RUTA_ARCHIVO` (ruta dentro de OneDrive).

> El Id de cliente **no es una contraseña**; puede ir en el código sin problema.

## Paso 2 – Publicar la app gratis en GitHub Pages

1. Crea una cuenta en **https://github.com** (si no tienes).
2. **New repository** → nombre `mis-finanzas` → **Public** → *Create repository*.
3. **Add file → Upload files** y arrastra **todo el contenido** de esta carpeta (`index.html`, `app.js`, `xlsx-registro.js`, `config.js`, `sw.js`, `manifest.webmanifest` y la carpeta `icons`) → *Commit changes*.
4. **Settings → Pages** → *Source: Deploy from a branch* → rama `main`, carpeta `/ (root)` → **Save**.
5. En 1–2 minutos queda en `https://TU-USUARIO.github.io/mis-finanzas/`.
   Verifica que esa dirección sea **exactamente** la URI de redirección del paso 1 (con la `/` final).

> El repositorio es público, pero **solo contiene el código**. Tus datos siguen privados en tu OneDrive y solo se ven después de iniciar sesión con tu cuenta.

## Paso 3 – Instalarla en el celular

- **Android (Chrome):** abre la dirección → menú ⋮ → **Instalar app** / *Agregar a pantalla principal*.
- **iPhone (Safari):** abre la dirección → botón Compartir → **Agregar a inicio**.

Ábrela → **Entrar con Microsoft** → acepta el permiso de *leer y escribir tus archivos*. Listo.

---

## Cómo se usa

1. Elige **Gasto** o **Ingreso**, escribe el valor, toca el tipo, pon una descripción → **Guardar**.
2. Verás “✓ Movimiento guardado en Excel (fila N)”.
3. **↻** vuelve a leer el Excel (resumen del mes y últimos movimientos).
4. **Sin señal:** el movimiento queda *pendiente* y se sube solo cuando vuelve la conexión (o con *Subir ahora*).

## Importante

- **Cierra el Excel en el PC mientras registras desde el celular**, o vuelve a abrirlo después.
  Si está abierto y lo guardas, OneDrive puede crear una copia en conflicto o sobrescribir lo que subió la app.
- La app escribe en la **fila siguiente al último registro con datos** (Fecha, Descripción o Valor). No dejes datos sueltos muy abajo en Registro.
- No cambies los títulos de la fila 6 de Registro (Fecha, Descripción, Movimiento, Tipo de gasto, Medio de pago, Valor, Notas) ni los de la fila 1 de *Listas* (“Tipos de gasto”, “Tipos de ingreso”, “Medios de pago”). Puedes agregar columnas nuevas sin problema.
- Las categorías de la app salen de la hoja **Listas**: si agregas o renombras una ahí, toca ↻ en la app.
- La sesión dura alrededor de 24 horas; después, solo vuelve a tocar *Entrar con Microsoft*.
- Si cambias archivos de la app y el celular sigue mostrando la versión vieja, sube el número en `sw.js` (`mis-finanzas-vN`) y el `?v=N` de los `<script>` en `index.html`.
