# Sintropía Social — PRD

## Problem statement (original, ES)
Modificaciones sobre el ZIP `Investigacion-social-main.zip`:
1. Quitar el plan de descarga de $18 MXN y la suscripción mensual de $49 MXN. Dejar únicamente donaciones, con el botón hacia: https://www.paypal.com/donate/?hosted_button_id=84RDSWN99GWCN
2. Hacer que funcionen las preguntas que se despliegan en `diagnostico.html` después de elegir una herramienta (Árbol de problemas, Línea base, Metodología SMART, Marco lógico, Talleres, Propuesta de intervención).
3. Subir los cambios a GitHub (vía la opción "Save to Github" del chat).

## Project type
Sitio estático HTML/JS (sin backend Node/Python). Backend de datos vía Google Apps Script (`Code.gs`, `app/CODIGO_COMPLETO.gs`).

## Files (en /app root)
- index.html — landing + repositorio bibliográfico (modificado: sin $18/$49, donación PayPal)
- diagnostico.html — diagnóstico social rápido (modificado: bug onclick corregido)
- registro.html — registro de usuarios (modificado: quitada referencia a $18 MXN)
- pago.html — convertido a página de donación (ya no procesa pagos)
- blog.html, contribuir.html, admin.html — sin cambios
- config.js, Code.gs, CODE_TEST_SIMPLE.gs, app/CODIGO_COMPLETO.gs — sin cambios
- logo.png

## What was implemented (this session — 2026-01)
- ✅ index.html: removidos botones "$18 MXN" y "Suscribirme $49 MXN/mes". Reemplazada la card "Planes de acceso" por "Apoya el proyecto" con botón Donar con PayPal (hosted_button_id=84RDSWN99GWCN).
- ✅ index.html: modal de descarga simplificado — descarga de PDF libre y gratuita + CTA de donación.
- ✅ index.html: eliminado SDK de PayPal y toda la lógica de cobro/suscripción (createOrder, createSubscription, paypal.Buttons render).
- ✅ diagnostico.html: corregidos atributos `onclick` rotos en `selectHerr()` y `pickOpc()` que impedían el despliegue de las preguntas tras elegir una herramienta. (Era un bug de escape `\"` dentro de atributos HTML con comillas dobles → cambiado a `\'`.)
- ✅ registro.html: actualizada la lista de beneficios — descarga gratuita en PDF.
- ✅ pago.html: reemplazada por una página simple de donación (PayPal donate link).
- ✅ Pruebas locales (Playwright): clic en "Árbol de problemas" abre el cuestionario; clic en "Línea base" + opciones múltiples seleccionables; modal de descarga muestra "Descarga libre y gratuita" + botón de donar.

## Backlog / Future
- P1: Verificar visualmente todas las herramientas (Marco lógico, SMART, Talleres, Propuesta) — flujo es idéntico, ya validado para Árbol y Línea base.
- P2: Considerar agregar el botón "Donar con PayPal" también en `blog.html` y `contribuir.html` para mayor visibilidad.
- P2: `pago.html` ahora es solo informativa — eventualmente podría eliminarse del repo si ningún enlace apunta a ella.

## Next action items
- Usuario: usar el botón "Save to Github" del chat para publicar los cambios en su repositorio `Investigacion-social`.
