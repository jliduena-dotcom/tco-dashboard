# 📊 Dashboard TCO — ONNET Fibra

Dashboard de seguimiento de TCOs. Se actualiza automáticamente cada vez que subes el archivo Excel.

---

## 🚀 Configuración inicial (solo una vez)

### Paso 1 — Habilitar GitHub Pages

1. Entra a tu repositorio en GitHub
2. Haz clic en **Settings** (arriba a la derecha)
3. En el menú izquierdo busca **Pages**
4. En **Source** selecciona **Deploy from a branch**
5. En **Branch** selecciona `main` y carpeta `/docs`
6. Haz clic en **Save**

Después de 1-2 minutos tu dashboard estará en:
```
https://TU_USUARIO.github.io/tco-dashboard/
```

### Paso 2 — Dar permisos al workflow

1. Ve a **Settings → Actions → General**
2. Baja hasta **Workflow permissions**
3. Selecciona **Read and write permissions**
4. Haz clic en **Save**

---

## 🔄 Cómo actualizar el dashboard (uso diario)

Cada vez que tengas un Excel nuevo:

1. Entra a tu repositorio en GitHub (en el navegador)
2. Haz clic en la carpeta **`data/`**
3. Haz clic en el botón **Add file → Upload files**
4. Arrastra tu archivo Excel *(debe llamarse exactamente igual)*
5. Haz clic en **Commit changes**

✅ GitHub genera el dashboard automáticamente (tarda ~1 minuto).  
✅ Tu URL pública se actualiza sola.

---

## 📁 Estructura del proyecto

```
tco-dashboard/
│
├── data/
│   └── Resultado_respuesta_ONNET_SN.xls   ← Aquí subes tu Excel
│
├── docs/
│   └── index.html                          ← Dashboard (generado automático)
│
├── scripts/
│   └── generar_dashboard.py                ← Script que procesa el Excel
│
└── .github/
    └── workflows/
        └── generar_dashboard.yml           ← Automatización de GitHub
```

---

## ❓ Preguntas frecuentes

**¿Puedo cambiar el nombre del archivo Excel?**  
Sí, pero debes actualizar la línea `EXCEL_PATH` en `scripts/generar_dashboard.py`.

**¿El dashboard es público?**  
Sí, con GitHub Pages gratuito es público. Si necesitas privacidad, avísame.

**¿Funciona con .xlsx también?**  
Sí, el script detecta la extensión automáticamente.

**¿Cuánto tiempo tarda en actualizarse?**  
Aproximadamente 1 minuto después de subir el archivo.

**¿Tiene costo?**  
No. GitHub Actions en repositorios públicos es 100% gratuito.
