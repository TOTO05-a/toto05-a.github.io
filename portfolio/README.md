# Portafolio de Eduardo Antonio Cruz Bolaños

Sitio estático bilingüe (ES/EN) hecho con HTML5, CSS3 y JavaScript Vanilla. Sin frameworks, sin dependencias de JS y sin proceso de build.

## Ejecutarlo

Abre `index.html` en el navegador. También puedes usar un servidor estático:

```bash
cd portfolio
python3 -m http.server 8000   # http://localhost:8000
```

## Estructura

```
portfolio/
├── index.html
├── css/
│   ├── style.css          Variables de color, componentes y secciones
│   └── responsive.css     Ajustes para laptop, tablet y móvil
├── js/
│   ├── translations.js    Todos los textos en ES y EN
│   └── main.js            Idioma, menú, pestañas, animaciones
├── assets/
│   ├── images/
│   │   ├── profile.jpg                      ← tu foto
│   │   └── projects/
│   │       ├── tv-programs.jpg              ← captura TV Programs App
│   │       ├── freelance-manager.jpg        ← captura Gestor de Proyectos Freelance
│   │       └── inventario-gamer.jpg         ← captura Inventario Gamer
│   ├── icons/logo.svg                       ← tu logo
│   └── documents/CV-Eduardo-Cruz.pdf        ← tu CV
└── README.md
```

Ninguno de esos archivos existe todavía. Mientras falten, el sitio muestra placeholders (foto, capturas, logo neutro). El botón de CV descargará el PDF cuando lo coloques en esa ruta.

## Dónde colocar cada archivo

- **Foto:** `assets/images/profile.jpg` (cuadrada, mínimo 600×600 px).
- **CV:** `assets/documents/CV-Eduardo-Cruz.pdf`.
- **Logo:** `assets/icons/logo.svg`. También se usa como favicon.
- **Capturas de proyectos:** `assets/images/projects/` con los nombres indicados arriba (proporción 16:9).

## GitHub y Demo de los proyectos

En `index.html`, dentro de cada `<article class="glass project">`, hay dos enlaces `GitHub` y `Demo` marcados como pendientes. Para activarlos:

1. Agrega `href="URL"` (y `target="_blank" rel="noopener noreferrer"`).
2. Elimina `aria-disabled="true"`, la clase `is-disabled` y el atributo `title`/`data-i18n-title`.

Si un proyecto no tiene demo, elimina ese enlace.

## Agregar un proyecto

Duplica un bloque `<article class="glass project">`, cambia el título, la descripción, los chips y la ruta de la imagen. Si el texto debe traducirse, agrega las claves en `js/translations.js` (ES y EN) y usa `data-i18n="clave"`.

## Cambiar textos e idioma

Los textos están en `js/translations.js`. El HTML contiene la versión en español como respaldo, así que si cambias un texto también cámbialo en `index.html`. El idioma elegido se guarda en `localStorage` (clave `lang`) y el idioma inicial es español.

## Cambiar la paleta

Todos los colores están en las variables `:root` al inicio de `css/style.css` (`--color-background`, `--color-surface`, `--color-primary`, `--color-secondary`, `--color-text`, `--color-muted`).

## Desplegar

Es un sitio estático: sirve en GitHub Pages, Netlify o Vercel subiendo la carpeta tal cual.
