# Prueba de Impresión

Este proyecto permite probar distintos métodos de impresión para una impresora
EPSON TM-U950. Proporciona una interfaz gráfica donde se puede seleccionar el
método de impresión y la impresora instalada en el sistema, incluso si se
encuentra fuera de línea.

## Requisitos

- Python 3
- [pywin32](https://pypi.org/project/pywin32/) para acceder a `win32print` en
  sistemas Windows.

Instala las dependencias con:

```bash
pip install -r requirements.txt
```

## Uso

Ejecuta el menú de impresión con:

```bash
python menu_impresion.py
```

Selecciona la impresora y el método de impresión deseado para enviar una
factura de prueba o ver una vista previa.
