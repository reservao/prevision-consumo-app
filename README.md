# Previsión de Consumo – Transformador de datos

App web que transforma archivos `.xls` de previsión de consumo (con datos apilados) en una tabla Excel ordenada y categorizada por Unidad Agregada.

## ¿Qué hace?

- Recibe un archivo `.xls` o `.xlsx` con el formato de previsión de consumo
- Detecta automáticamente cada Unidad Agregada y sus productos
- Genera un Excel descargable con tabla formateada, colores y filtros

## Cómo usarla localmente

```bash
pip install -r requirements.txt
python app.py
```

Luego abre `http://localhost:5000` en tu navegador.

## Despliegue en Render

1. Conecta este repositorio en [render.com](https://render.com)
2. Tipo: **Web Service**
3. Build command: `pip install -r requirements.txt`
4. Start command: `gunicorn app:app`
5. ¡Listo!
