# Kruger PPTX API

Backend Flask + python-pptx para generación de presentaciones Kruger.

## Archivos

- `app.py` — API Flask con todos los slides
- `requirements.txt` — dependencias
- `render.yaml` — configuración de deploy en Render

## Deploy en Render

1. Sube este repositorio a GitHub
2. En Render: New → Web Service → conecta el repo
3. Render detecta `render.yaml` automáticamente
4. Deploy tarda ~2 minutos

## Endpoint

`POST /generate`

Body JSON:
```json
{
  "vertical": {
    "nm": "Telecomunicaciones",
    "pain": ["..."],
    "prods": [{"n":"...","cat":"...","d":"...","ico":"..."}],
    "servs": [{"n":"01","t":"...","d":"..."}],
    "cases": [{"n":"...","sub":"...","ch":"...","ap":"...","out":"..."}]
  },
  "cfg": {
    "name": "Martín Almirati",
    "title": "Gerente",
    "email": "malmirati@kruger.com",
    "phone": "0983972068",
    "client": "Claro"
  }
}
```

Respuesta: archivo `.pptx` como descarga binaria.

`GET /health` — health check
