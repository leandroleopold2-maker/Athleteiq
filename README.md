# ⚡ AthleteIQ — Performance Tracker

App para análisis de rendimiento físico en fútbol. Importa datos de **MyJump Lab** y **PhotoFinish**.

## Cómo subir a Vercel (paso a paso)

### 1. Subir el código a GitHub

1. Entrá a [github.com](https://github.com) y creá una cuenta si no tenés
2. Hacé clic en **New repository**, llamalo `athleteiq`, dejalo en Public
3. En la página del repo recién creado, hacé clic en **uploading an existing file**
4. Arrastrá **todos los archivos de esta carpeta** (incluyendo la carpeta `src/`)
5. Hacé clic en **Commit changes**

### 2. Conectar con Vercel

1. Entrá a [vercel.com](https://vercel.com) y creá una cuenta (podés usar tu cuenta de GitHub)
2. Hacé clic en **Add New Project**
3. Seleccioná tu repositorio `athleteiq`
4. Vercel detecta Vite automáticamente — no hay que cambiar nada
5. Hacé clic en **Deploy**

En 1-2 minutos vas a tener tu app en un link propio (ej: `athleteiq.vercel.app`).

### 3. Actualizar la app en el futuro

Cuando yo te genere una versión nueva de `App.jsx`, solo tenés que:
1. Ir a tu repo en GitHub
2. Abrir el archivo `src/App.jsx`
3. Hacer clic en el ícono del lápiz (editar)
4. Pegar el nuevo contenido y guardar
5. Vercel se actualiza solo en segundos

## Datos guardados

Los datos se guardan en el navegador (localStorage). Esto significa:
- Persisten aunque cierres y abras la app
- Son por dispositivo (celu y compu tienen sus propios datos)
- Si borrás caché del navegador se pierden → usá el botón **Exportar sesión** para hacer backup

## Stack

- React 18 + Vite
- Recharts para gráficos
- Sin base de datos — todo en el navegador
