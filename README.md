# Notador - Generador de Boletines

Aplicación de escritorio para generar boletines escolares a partir de archivos Excel y plantillas Word.

## Características

- Interfaz gráfica moderna y fácil de usar
- Procesamiento por grado y grupo
- Selección múltiple de estudiantes mediante checkboxes
- Generación de boletines en formato Word y PDF
- Cálculo automático de promedios y materias perdidas
- Soporte para múltiples periodos académicos

## Requisitos

- Python 3.x
- Microsoft Word (instalado en el sistema)
- Bibliotecas Python requeridas (ver requirements.txt)

## Instalación

1. Clonar el repositorio:
```bash
git clone [URL_DEL_REPOSITORIO]
cd Notador
```

2. Crear y activar entorno virtual:
```bash
python -m venv .venv
.venv\Scripts\activate
```

3. Instalar dependencias:
```bash
pip install -r requirements.txt
```

## Uso

1. Ejecutar la aplicación:
```bash
python notador.py
```

2. O usar el ejecutable compilado:
- Ejecutar `dist/notador.exe`

## Estructura del Archivo Excel

El archivo Excel debe contener las siguientes columnas:
- ID y nombre del estudiante (formato: "123456789 - APELLIDOS NOMBRES")
- Columnas de materias con notas
- GRUPO (opcional)
- PERIODO (opcional)

## Compilación

Para generar el ejecutable:
```bash
pyinstaller --onefile --noconsole notador.py
```

## Licencia

[Especificar licencia]
