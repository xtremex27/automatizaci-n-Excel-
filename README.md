# Corrector de Hoja de Ruta — Excel HR

Servidor Python (Flask) que recibe archivos Excel con formato físico multi-fila, los corrige automáticamente y devuelve el archivo con el formato correcto (6 columnas).

## ¿Qué hace?

- Recibe un `.xls` o `.xlsx` vía `POST /corregir-hr`
- Auto-detecta el número de HR y el Distrito
- Extrae todas las entradas con código de barras `LS...CW`
- Devuelve el archivo corregido con 6 columnas

## Uso local

```bash
pip install -r requirements.txt
python servidor.py
```

Abrir `index.html` en el navegador y usar `http://localhost:5000/corregir-hr`

## Estructura

```
├── servidor.py       # Servidor Flask con la lógica de corrección
├── index.html        # Interfaz web para subir el Excel
├── requirements.txt  # Dependencias Python
├── vercel.json       # Configuración para despliegue en Vercel
└── .gitignore
```

## Despliegue en Vercel

Conectar el repositorio GitHub en [vercel.com](https://vercel.com) y hacer deploy directamente.
