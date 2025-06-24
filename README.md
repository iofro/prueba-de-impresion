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
Es necesario que dicha fuente esté instalada en el sistema; de lo contrario
se utilizará `Courier New` como alternativa.

Si notas que los caracteres se ven cortados o poco legibles, verifica qué
fuentes incluye el driver de tu impresora (por ejemplo, "EPSON Draft 10cpi" o
"EPSON Roman 12cpi") e indica el nombre exacto en la constante
`DEFAULT_FONT_NAME` de `menu_impresion.py`.

## Configuración de la página

La función `configurar_mapeo` asume que la factura ocupa un área de
27.5 × 16.6 cm. Estos valores funcionan para la TM‑U950 pero pueden variar
según el modelo y el tamaño de papel. El origen de la vista (offsets
`PHYSICALOFFSETX` y `PHYSICALOFFSETY`) se consulta directamente a la
impresora porque cada driver aplica desplazamientos diferentes.
