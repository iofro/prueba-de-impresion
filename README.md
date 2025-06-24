# Prueba de Impresión

Este proyecto permite probar la impresión de facturas en una impresora
EPSON TM-U950 utilizando únicamente la API `win32ui`. La aplicación ofrece una
interfaz gráfica donde se selecciona la impresora instalada en el sistema,
incluso si se encuentra fuera de línea, y el tipo de alineación deseado.

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

Selecciona la impresora y el método de alineación deseado para enviar una
factura de prueba.

Los métodos disponibles incluyen únicamente variantes de `win32ui`:

- Coordenadas absolutas
- Alineado por espacios
- Alineado con tabulaciones
- Alineado con CRLF

## Fuente de impresión

La aplicación intenta usar de forma predeterminada una fuente de matriz de
puntos (`Dot Matrix`) para asemejar la salida de las impresoras de impacto.
Puedes cambiar el nombre de la fuente modificando la constante
`DEFAULT_FONT_NAME` en `menu_impresion.py`. Si la fuente no está instalada en
el sistema se utilizará `Courier New` como alternativa.
