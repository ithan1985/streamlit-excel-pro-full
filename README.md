# Excel Explorer Pro

Aplicación web construida con Streamlit para cargar, explorar y filtrar archivos Excel.
Incluye soporte para Docker y Docker Compose, ideal para pruebas locales, clases y despliegues simples.

## Características

- Carga de archivos Excel desde la interfaz
- Soporte para múltiples hojas
- Búsqueda global
- Filtros por columnas de texto
- Filtros por columnas numéricas
- Métricas rápidas
- Resumen estadístico
- Descarga de resultados filtrados en CSV

## Estructura del proyecto

```bash
streamlit-excel-pro-full/
├── .gitignore
├── README.md
├── app.py
├── requirements.txt
├── Dockerfile
├── docker-compose.yml
├── .streamlit/
│   └── config.toml
├── data/
│   └── ejemplo_productos.xlsx
└── screenshots/
```

## Requisitos

- Docker Desktop
- WSL2
- Docker Compose

## Ejecución con Docker Compose

```bash
docker compose up --build
```

Abrir en navegador:

```bash
http://localhost:8501
```

## Detener contenedores

```bash
docker compose down
```

## Ejecución local sin Docker

Instalar dependencias:

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Uso

1. Sube un archivo Excel desde la interfaz.
2. Selecciona la hoja.
3. Usa la barra lateral para buscar y filtrar.
4. Revisa métricas, resumen y descarga.

## GitHub

```bash
git init
git add .
git commit -m "Proyecto inicial Streamlit Excel Pro"
git branch -M main
git remote add origin <URL_DEL_REPO>
git push -u origin main
```
