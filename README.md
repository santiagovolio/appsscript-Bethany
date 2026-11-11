# Proyecto Bethany · Código de Apps Script

Este repositorio contiene el **código fuente** de la aplicación de Proyecto Bethany para **salidas de fin de semana** y **solicitudes de mantenimiento**.

Los dos archivos de este repositorio son exactamente los que corren dentro del proyecto de **Google Apps Script** de la cuenta de Google de Proyecto Bethany. GitHub solo los guarda como respaldo y referencia; la aplicación no se ejecuta desde aquí.

🔗 **Página de la aplicación:** https://proyectobethany.github.io/proyectobethany/
📁 **Repositorio de la página (GitHub Pages):** https://github.com/proyectobethany/proyectobethany

---

## Archivos

| Archivo | Dónde va en Apps Script | Qué hace |
|---|---|---|
| `Code.gs` | Archivo de script `Código.gs` | Servidor: cuentas y contraseñas, sesiones, reglas de salidas y tokens, semanas del cuatrimestre, boletas de mantenimiento, notificaciones por correo, libros de Google Sheets y generación de la boleta en Word (.docx). |
| `Index.html` | Archivo HTML llamado `Index` | Interfaz: todas las pantallas de las mamás, la administración y el personal de mantenimiento, adaptadas a celular y computadora. |

El manifiesto `appsscript.json` del proyecto usa la zona horaria `America/Costa_Rica`, el entorno `V8` y la aplicación web configurada para ejecutarse como la cuenta propietaria y estar disponible para cualquier usuario.

## Cómo funciona

```
Celular o computadora
        │
        ▼
Página de GitHub Pages (proyectobethany.github.io/proyectobethany)
        │  muestra la app a pantalla completa (iframe)
        ▼
Aplicación web de Google Apps Script  ←  Code.gs + Index.html (este repositorio)
        │  lee y escribe
        ▼
Libros de Google Sheets en el Drive de Proyecto Bethany
```

## Perfiles

- **Mamás residentes:** crean su cuenta y su perfil, solicitan salidas usando sus tokens del mes (y el token de excepción cuando lo necesitan), llenan la boleta de mantenimiento con su firma y reciben notificaciones.
- **Administración:** aprueba o rechaza salidas, ve la semana del cuatrimestre, gestiona cuentas, revisa y completa las boletas de mantenimiento, configura las reglas y descarga los Excel.
- **Personal de mantenimiento:** solo ve las boletas de mantenimiento y completa la sección 4. Sus cuentas las crea la administración.

## Libros de Google Sheets que crea

| Libro | Contenido |
|---|---|
| Bethany - Sistema | Usuarios, configuración, notificaciones y firmas. **No compartir.** |
| Bethany - Solicitudes | Todas las solicitudes de salida, por estado. |
| Bethany - Solicitudes aprobadas | Registro de quiénes salen. |
| Bethany - Residentes | Perfiles de las mamás y de sus hijos e hijas. **No compartir.** |
| Bethany - Mantenimiento | Boletas de mantenimiento activas en la app. |
| Bethany - Registro de mantenimiento | Historial permanente de todas las boletas, ordenadas por estado. |

## Instalación nueva

1. En script.google.com, con la cuenta de Google de Proyecto Bethany, crear un proyecto nuevo.
2. Pegar `Code.gs` en `Código.gs` y crear un archivo HTML llamado `Index` con el contenido de `Index.html`.
3. Ejecutar la función `instalar` y autorizar los permisos. En el **Registro de ejecución** aparece un **código de instalación**.
4. **Implementar → Nueva implementación → Aplicación web** (Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario).
5. Abrir la URL de la aplicación, entrar a **Configuración inicial** y crear la primera cuenta de administración con ese código. El código sirve una sola vez y ninguna contraseña queda escrita en el código.
6. Poner la URL `/exec` en el `index.html` del repositorio de la página.

## Actualizar el código

1. Reemplazar `Código.gs` e `Index` en el proyecto de Apps Script con las versiones nuevas de este repositorio.
2. **Implementar → Administrar implementaciones → ✏️ → Versión: Nueva versión → Implementar.** La URL no cambia.
3. Los cambios de estructura en los libros de Excel se aplican solos la primera vez que se usa la app.

## Seguridad y privacidad

- Este repositorio **no contiene datos personales, contraseñas, IDs de los libros ni la URL `/exec`**. Esos datos viven solo en la cuenta de Google de Proyecto Bethany.
- Las contraseñas se guardan con hash y sal; el ingreso se bloquea temporalmente tras varios intentos fallidos.
- La app guarda datos personales de las mamás y de sus hijos e hijas, protegidos por la **Ley N.º 8968** de Costa Rica.

## Créditos

- **Desarrollo:** Santiago Volio Argüello, Trabajo Comunal Universitario (TCU), LEAD University.
- **Organización:** Proyecto Bethany, Costa Rica.
- Desarrollado con apoyo de herramientas de inteligencia artificial.
