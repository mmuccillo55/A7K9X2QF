# Rack Index — C12

Documentación de cableado del estudio.

## Estructura

```
Conexiones.xlsx        ← fuente de verdad (editá este)
parse.py               ← convierte xlsx → json
connections.json       ← generado automáticamente (no editar)
index.html             ← interfaz web
.github/workflows/     ← CI/CD automático
```

## Cómo actualizar

1. Editá `Conexiones.xlsx`
2. Hacé commit y push
3. GitHub Actions corre `parse.py` automáticamente y actualiza `connections.json`
4. El index refleja los cambios al recargar

## Uso local

```bash
# Regenerar el JSON manualmente
python parse.py Conexiones.xlsx connections.json

# Ver el index localmente (necesita servidor HTTP)
python -m http.server 8000
# abrir http://localhost:8000
```

## GitHub Pages

El index vive en: `https://<usuario>.github.io/<repo>/`

Para activarlo: Settings → Pages → Source: Deploy from branch → `main` → `/root`
