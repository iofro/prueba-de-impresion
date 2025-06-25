import tkinter as tk
from tkinter import ttk, messagebox

# Fuente predeterminada para las impresiones. Cambia el valor si tu impresora
# ofrece un nombre distinto para su tipo de letra de matriz de puntos.
DEFAULT_FONT_NAME = "Dot Matrix"
# Escala horizontal para la fuente. Un valor menor a ``1.0`` estrecha el texto
# sin modificar la altura.
DEFAULT_FONT_WIDTH_SCALE = 0.4

# Intenta importar win32print, si no está disponible muestra un error al intentar imprimir
try:
    import win32print
except ImportError:
    win32print = None
try:
    import win32con
except ImportError:
    win32con = None
try:
    import win32ui
except ImportError:
    win32ui = None


# Comandos de la EPSON TM-U950 para seleccionar el modo de impresi\u00f3n.
# SLIP4_MODE indica que se utilizará la posición "slip 4" recomendada
# para facturas en la bandeja de formularios.
SLIP4_MODE = b"\x1B\x69\x04"

def cm_a_twips(valor_cm: float) -> int:
    """Convierte cent\u00edmetros a TWIPS (1/1440 pulgadas)."""
    return round(valor_cm * 566.93)

# Coordenadas est\u00e1ticas para la factura (en cent\u00edmetros)
HEADER_COORDS = {
    "cliente": (4.45, 4.80),
    "direccion": (4.45, 5.40),
    "fecha": (3.81, 6.40),
    "giro": (3.81, 6.90),
    "fecha_remision": (3.81, 7.50),
    "condicion_pago": (4.45, 8.00),
    "vendedor": (4.45, 8.50),
    "nrc": (7.62, 6.40),
    "no_rem": (7.62, 7.50),
    "nit": (11.43, 6.40),
    "orden_no": (12.07, 7.50),
    "venta_cuenta_de": (11.43, 8.00),
    "fecha_nota_ant": (11.75, 8.50),
}

PRODUCT_HEADER = {
    "cantidad": (2.22, 10.10),
    "descripcion": (3.90, 10.10),
    "precio_unitario": (9.21, 10.10),
    "ventas_exentas": (11.11, 10.10),
    "ventas_no_sujetas": (12.70, 10.10),
    "ventas_gravadas": (14.10, 10.10),
}

TOTALS_COORDS = {
    "literal": (2.22, 22.23),
    "sumas": (14.10, 21.59),
    "iva": (14.10, 22.23),
    "subtotal": (14.10, 22.86),
    "iva_retenido": (14.10, 23.45),
    "no_sujetas": (14.10, 24.00),
    "ventas_exentas": (14.10, 24.60),
    "total": (14.10, 25.08),
}

PRODUCT_ROW_START_Y = 10.70
ROW_HEIGHT = 0.6


def seleccionar_fuente(
    dc,
    puntos: int = 12,
    nombre: str = DEFAULT_FONT_NAME,
    escala_ancho: float = DEFAULT_FONT_WIDTH_SCALE,
):
    """Crea y selecciona en el DC una fuente con las dimensiones indicadas.

    ``escala_ancho`` permite modificar la proporción horizontal del texto sin
    afectar su altura. Un valor menor que ``1.0`` hará que la fuente sea más
    estrecha.
    """
    if win32ui is None:
        return None, None
    alto = -puntos * 20
    ancho = 0 if escala_ancho == 1.0 else int(abs(alto) * escala_ancho)
    try:
        font = win32ui.CreateFont({"name": nombre, "height": alto, "width": ancho})
    except Exception:
        font = win32ui.CreateFont({"name": "Courier New", "height": alto, "width": ancho})
    old = dc.SelectObject(font)
    return font, old

def eliminar_fuente(font):
    """Libera una fuente de forma segura."""
    if not font:
        return
    try:
        font.DeleteObject()
    except AttributeError:
        try:
            import win32gui
            win32gui.DeleteObject(font.GetHandle())
        except Exception:
            pass

def configurar_mapeo(dc):
    """Configura el mapeo del DC para coincidir con el tamaño de la factura.

    Se asume un papel de 27.5 × 16.6 cm. Los desplazamientos (`PHYSICALOFFSETX`,
    `PHYSICALOFFSETY`) se leen de la impresora y pueden variar entre modelos.
    """
    if not win32con:
        return
    dc.SetMapMode(win32con.MM_TWIPS)
    ancho = dc.GetDeviceCaps(win32con.HORZRES)
    alto = dc.GetDeviceCaps(win32con.VERTRES)
    offset_x = dc.GetDeviceCaps(win32con.PHYSICALOFFSETX)
    offset_y = dc.GetDeviceCaps(win32con.PHYSICALOFFSETY)
    # PyCDC no expone los métodos *Ex, por lo que utilizamos las variantes
    # simples que aceptan una tupla como parámetro.
    dc.SetWindowExt((cm_a_twips(27.5), cm_a_twips(16.6)))
    dc.SetViewportExt((ancho, -alto))
    dc.SetViewportOrg((-offset_x, offset_y))


def activar_modo_slip(printer_name: str) -> bool:
    """Activa el modo SLIP4 en la impresora para usar la bandeja de formularios."""
    if win32print is None:
        messagebox.showerror("Error", "win32print no est\u00e1 instalado. Instala pywin32.")
        return False
    try:
        hprinter = win32print.OpenPrinter(printer_name)
        win32print.StartDocPrinter(hprinter, 1, ("Modo SLIP", None, "RAW"))
        win32print.StartPagePrinter(hprinter)
        win32print.WritePrinter(hprinter, SLIP4_MODE)
        win32print.EndPagePrinter(hprinter)
        win32print.EndDocPrinter(hprinter)
        win32print.ClosePrinter(hprinter)
        return True
    except Exception as e:
        messagebox.showerror("Error de impresi\u00f3n", str(e))
        return False


def generar_factura_datos():
    """Devuelve la informaci\u00f3n de la factura de prueba.

    Retorna una tupla con tres elementos:

    ``encabezado``
        Diccionario de campos del encabezado.

    ``productos``
        Lista de tuplas con las columnas de la tabla de productos.

    ``totales``
        Diccionario con los importes de resumen y totales.
    """

    encabezado = {
        "cliente": "Ferreter\u00eda El Martillo S.A. de C.V.",
        "direccion": "Calle El Progreso #23, San Miguel",
        "fecha": "22/06/2025",
        "giro": "Venta de materiales de construcci\u00f3n",
        "fecha_remision": "21/06/2025",
        "condicion_pago": "Cr\u00e9dito 30 d\u00edas",
        "vendedor": "Mar\u00eda L\u00f3pez",
        "nrc": "129847-3",
        "no_rem": "REM-4567",
        "nit": "0614-150385-101-8",
        "orden_no": "ORD-98231",
        "venta_cuenta_de": "Constructora Innovar S.A. de C.V.",
        "fecha_nota_ant": "20/06/2025",
    }

    productos = [
        ("10", "Bolsa de cemento gris", "6.50", "0.00", "0.00", "65.00"),
    ]

    totales = {
        "literal": "Sesenta y cinco dólares con 00/100",
        "sumas": "65.00",
        "iva": "8.45",
        "subtotal": "73.45",
        "iva_retenido": "0.00",
        "no_sujetas": "0.00",
        "ventas_exentas": "0.00",
        "total": "73.45",
    }

    return encabezado, productos, totales

def imprimir_factura_win32ui(printer_name):
    """Imprime usando win32ui para enviar texto a coordenadas absolutas."""
    if win32print is None or win32ui is None:
        messagebox.showerror(
            "Error",
            "win32print/win32ui no están disponibles. Instala pywin32 en Windows.",
        )
        return
    try:
        encabezado, productos, totales = generar_factura_datos()
        hprinter = win32print.OpenPrinter(printer_name)
        win32print.StartDocPrinter(hprinter, 1, ("Factura win32ui", None, "RAW"))
        win32print.StartPagePrinter(hprinter)
        win32print.WritePrinter(hprinter, SLIP4_MODE)
        win32print.EndPagePrinter(hprinter)
        win32print.EndDocPrinter(hprinter)

        dc = win32ui.CreateDC()
        dc.CreatePrinterDC(printer_name)
        configurar_mapeo(dc)
        dc.StartDoc("Factura win32ui")
        dc.StartPage()
        font, old_font = seleccionar_fuente(dc, 12)

        def draw(x_cm, y_cm, texto):
            dc.TextOut(cm_a_twips(x_cm), -cm_a_twips(y_cm), texto)

        # Encabezado (solo datos)
        for campo, (x, y) in HEADER_COORDS.items():
            draw(x, y, encabezado.get(campo, ""))

        # Encabezado de la tabla de productos
        draw(PRODUCT_HEADER["cantidad"][0], PRODUCT_HEADER["cantidad"][1], "Cantidad")
        draw(PRODUCT_HEADER["descripcion"][0], PRODUCT_HEADER["descripcion"][1], "Descripción")
        draw(PRODUCT_HEADER["precio_unitario"][0], PRODUCT_HEADER["precio_unitario"][1], "Precio unitario")
        draw(PRODUCT_HEADER["ventas_exentas"][0], PRODUCT_HEADER["ventas_exentas"][1], "Ventas exentas")
        draw(PRODUCT_HEADER["ventas_no_sujetas"][0], PRODUCT_HEADER["ventas_no_sujetas"][1], "Ventas no sujetas")
        draw(PRODUCT_HEADER["ventas_gravadas"][0], PRODUCT_HEADER["ventas_gravadas"][1], "Ventas gravadas")

        # Productos
        for i, (cant, desc, prec, ex, ns, grav) in enumerate(productos):
            y = PRODUCT_ROW_START_Y + i * ROW_HEIGHT
            draw(PRODUCT_HEADER["cantidad"][0], y, cant)
            draw(PRODUCT_HEADER["descripcion"][0], y, desc)
            draw(PRODUCT_HEADER["precio_unitario"][0], y, prec)
            draw(PRODUCT_HEADER["ventas_exentas"][0], y, ex)
            draw(PRODUCT_HEADER["ventas_no_sujetas"][0], y, ns)
            draw(PRODUCT_HEADER["ventas_gravadas"][0], y, grav)

        # Totales
        for campo, (x, y) in TOTALS_COORDS.items():
            draw(x, y, totales.get(campo, ""))

        if old_font:
            dc.SelectObject(old_font)
        if font:
            eliminar_fuente(font)
        dc.EndPage()
        dc.EndDoc()
        dc.DeleteDC()
        win32print.ClosePrinter(hprinter)
        messagebox.showinfo("Éxito", "Factura enviada a la impresora.")
    except Exception as e:
        messagebox.showerror("Error de impresión", str(e))


def imprimir_factura_win32ui_espacios(printer_name):
    """Imprime la factura alineando el texto con espacios."""
    if win32print is None or win32ui is None:
        messagebox.showerror(
            "Error",
            "win32print/win32ui no están disponibles. Instala pywin32 en Windows.",
        )
        return
    try:
        encabezado, productos, totales = generar_factura_datos()
        hprinter = win32print.OpenPrinter(printer_name)
        win32print.StartDocPrinter(hprinter, 1, ("Factura win32ui espacios", None, "RAW"))
        win32print.StartPagePrinter(hprinter)
        win32print.WritePrinter(hprinter, SLIP4_MODE)
        win32print.EndPagePrinter(hprinter)
        win32print.EndDocPrinter(hprinter)

        dc = win32ui.CreateDC()
        dc.CreatePrinterDC(printer_name)
        configurar_mapeo(dc)
        dc.StartDoc("Factura win32ui espacios")
        dc.StartPage()
        font, old_font = seleccionar_fuente(dc, 12)

        def draw(x_cm, y_cm, texto):
            dc.TextOut(cm_a_twips(x_cm), -cm_a_twips(y_cm), texto)

        # Dibujar encabezado en las posiciones indicadas
        for campo, (x, y) in HEADER_COORDS.items():
            draw(x, y, encabezado.get(campo, ""))

        # Encabezado de columnas utilizando espacios
        draw(
            PRODUCT_HEADER["cantidad"][0],
            PRODUCT_HEADER["cantidad"][1],
            "Cant  Descripción             Precio  Exentas  NoSuj  Gravadas",
        )

        # Filas de productos
        for i, (cant, desc, prec, ex, ns, grav) in enumerate(productos):
            y_line = PRODUCT_ROW_START_Y + i * ROW_HEIGHT
            linea = f"{cant:<5}{desc:<23}{prec:>7}    {ex:>4}     {ns:>4}   {grav:>4}"
            draw(PRODUCT_HEADER["cantidad"][0], y_line, linea)

        # Totales
        for campo, (x, y) in TOTALS_COORDS.items():
            draw(x, y, totales.get(campo, ""))

        if old_font:
            dc.SelectObject(old_font)
        if font:
            eliminar_fuente(font)


        dc.EndPage()
        dc.EndDoc()
        dc.DeleteDC()
        win32print.ClosePrinter(hprinter)
        messagebox.showinfo("Éxito", "Factura enviada a la impresora.")
    except Exception as e:
        messagebox.showerror("Error de impresión", str(e))


def imprimir_factura_win32ui_tabs(printer_name):
    """Imprime la factura usando tabulaciones para alinear columnas."""
    if win32print is None or win32ui is None:
        messagebox.showerror(
            "Error",
            "win32print/win32ui no están disponibles. Instala pywin32 en Windows.",
        )
        return
    try:
        encabezado, productos, totales = generar_factura_datos()
        hprinter = win32print.OpenPrinter(printer_name)
        win32print.StartDocPrinter(hprinter, 1, ("Factura win32ui tabs", None, "RAW"))
        win32print.StartPagePrinter(hprinter)
        win32print.WritePrinter(hprinter, SLIP4_MODE)
        win32print.EndPagePrinter(hprinter)
        win32print.EndDocPrinter(hprinter)

        dc = win32ui.CreateDC()
        dc.CreatePrinterDC(printer_name)
        configurar_mapeo(dc)
        dc.StartDoc("Factura win32ui tabs")
        dc.StartPage()
        font, old_font = seleccionar_fuente(dc, 12)

        def draw(x_cm, y_cm, texto):
            dc.TextOut(cm_a_twips(x_cm), -cm_a_twips(y_cm), texto)

        # Encabezado de la factura en sus coordenadas
        for campo, (x, y) in HEADER_COORDS.items():
            draw(x, y, encabezado.get(campo, ""))

        # Encabezado de la tabla de productos en sus coordenadas
        draw(PRODUCT_HEADER["cantidad"][0], PRODUCT_HEADER["cantidad"][1], "Cantidad")
        draw(PRODUCT_HEADER["descripcion"][0], PRODUCT_HEADER["descripcion"][1], "Descripci\u00f3n")
        draw(
            PRODUCT_HEADER["precio_unitario"][0],
            PRODUCT_HEADER["precio_unitario"][1],
            "Precio Unitario",
        )
        draw(
            PRODUCT_HEADER["ventas_exentas"][0],
            PRODUCT_HEADER["ventas_exentas"][1],
            "V. Exentas",
        )
        draw(
            PRODUCT_HEADER["ventas_no_sujetas"][0],
            PRODUCT_HEADER["ventas_no_sujetas"][1],
            "V. No Sujetas",
        )
        draw(
            PRODUCT_HEADER["ventas_gravadas"][0],
            PRODUCT_HEADER["ventas_gravadas"][1],
            "V. Gravadas",
        )

        # Filas de productos usando tabulaciones para separar texto pero con coordenadas fijas
        for i, (cant, desc, prec, ex, ns, grav) in enumerate(productos):
            y_line = PRODUCT_ROW_START_Y + i * ROW_HEIGHT
            draw(PRODUCT_HEADER["cantidad"][0], y_line, cant)
            draw(PRODUCT_HEADER["descripcion"][0], y_line, desc)
            draw(PRODUCT_HEADER["precio_unitario"][0], y_line, prec)
            draw(PRODUCT_HEADER["ventas_exentas"][0], y_line, ex)
            draw(PRODUCT_HEADER["ventas_no_sujetas"][0], y_line, ns)
            draw(PRODUCT_HEADER["ventas_gravadas"][0], y_line, grav)

        # Totales en sus coordenadas
        for campo, (x, y) in TOTALS_COORDS.items():
            draw(x, y, totales.get(campo, ""))

        if old_font:
            dc.SelectObject(old_font)
        if font:
            eliminar_fuente(font)

        dc.EndPage()
        dc.EndDoc()
        dc.DeleteDC()
        win32print.ClosePrinter(hprinter)
        messagebox.showinfo("Éxito", "Factura enviada a la impresora.")
    except Exception as e:
        messagebox.showerror("Error de impresión", str(e))


def imprimir_factura_win32ui_crlf(printer_name):
    """Imprime la factura usando CRLF para los saltos de línea."""
    if win32print is None or win32ui is None:
        messagebox.showerror(
            "Error",
            "win32print/win32ui no están disponibles. Instala pywin32 en Windows.",
        )
        return
    try:
        encabezado, productos, totales = generar_factura_datos()
        hprinter = win32print.OpenPrinter(printer_name)
        win32print.StartDocPrinter(hprinter, 1, ("Factura win32ui CRLF", None, "RAW"))
        win32print.StartPagePrinter(hprinter)
        win32print.WritePrinter(hprinter, SLIP4_MODE)
        win32print.EndPagePrinter(hprinter)
        win32print.EndDocPrinter(hprinter)

        dc = win32ui.CreateDC()
        dc.CreatePrinterDC(printer_name)
        configurar_mapeo(dc)
        dc.StartDoc("Factura win32ui CRLF")
        dc.StartPage()
        font, old_font = seleccionar_fuente(dc, 12)

        def draw(x_cm, y_cm, texto):
            dc.TextOut(cm_a_twips(x_cm), -cm_a_twips(y_cm), texto)

        header_order = [
            "cliente",
            "direccion",
            "fecha",
            "giro",
            "fecha_remision",
            "condicion_pago",
            "vendedor",
            "nrc",
            "no_rem",
            "nit",
            "orden_no",
            "venta_cuenta_de",
            "fecha_nota_ant",
        ]

        lines = [encabezado.get(campo, "") for campo in header_order]
        lines.append("")
        lines.append("Cant  Descripción             Precio  Exentas  NoSuj  Gravadas")
        for cant, desc, prec, ex, ns, grav in productos:
            lines.append(f"{cant:<5}{desc:<23}{prec:>7}    {ex:>4}     {ns:>4}   {grav:>4}")
        lines.append("")
        lines.append(totales["literal"])
        lines.append(f"Sumas: {totales['sumas']}")
        lines.append(f"13% IVA: {totales['iva']}")
        lines.append(f"Subtotal: {totales['subtotal']}")
        lines.append(f"IVA retenido: {totales['iva_retenido']}")
        lines.append(f"Vtas no sujetas: {totales['no_sujetas']}")
        lines.append(f"Ventas exentas: {totales['ventas_exentas']}")
        lines.append(f"Venta total: {totales['total']}")

        y = 4.8
        line_height = 0.6
        for campo in header_order:
            draw(0, y, encabezado.get(campo, ""))
            y += line_height


        if old_font:
            dc.SelectObject(old_font)
        if font:
            eliminar_fuente(font)

        draw(0, 10.10, "Cant  Descripción             Precio  Exentas  NoSuj  Gravadas")

        y_base = 10.70
        for i, (cant, desc, prec, ex, ns, grav) in enumerate(productos):
            y_line = y_base + i * line_height
            draw(0, y_line, f"{cant:<5}{desc:<23}{prec:>7}    {ex:>4}     {ns:>4}   {grav:>4}")

        y = y_base + len(productos) * line_height + line_height
        draw(0, y, totales["literal"])
        y += line_height
        draw(0, y, f"Sumas: {totales['sumas']}")
        y += line_height
        draw(0, y, f"IVA: {totales['iva']}")
        y += line_height
        draw(0, y, f"Subtotal: {totales['subtotal']}")
        y += line_height
        draw(0, y, f"IVA retenido: {totales['iva_retenido']}")
        y += line_height
        draw(0, y, f"Vtas no sujetas: {totales['no_sujetas']}")
        y += line_height
        draw(0, y, f"Ventas exentas: {totales['ventas_exentas']}")
        y += line_height
        draw(0, y, f"Venta total: {totales['total']}")


        dc.EndPage()
        dc.EndDoc()
        dc.DeleteDC()
        win32print.ClosePrinter(hprinter)
        messagebox.showinfo("Éxito", "Factura enviada a la impresora.")
    except Exception as e:
        messagebox.showerror("Error de impresión", str(e))


def imprimir_segun_metodo():
    metodo = metodo_var.get()
    printer_name = printer_var.get()
    if not printer_name:
        messagebox.showwarning("Advertencia", "Debes ingresar el nombre de la impresora.")
        return
    if metodo == "win32ui (coordenadas absolutas)":
        imprimir_factura_win32ui(printer_name)
    elif metodo == "win32ui (alineado por espacios)":
        imprimir_factura_win32ui_espacios(printer_name)
    elif metodo == "win32ui (alineado con tabulaciones)":
        imprimir_factura_win32ui_tabs(printer_name)
    elif metodo == "win32ui (alineado con CRLF)":
        imprimir_factura_win32ui_crlf(printer_name)
    else:
        messagebox.showinfo("Info", f"Método '{metodo}' aún no implementado.")


if __name__ == "__main__":
    root = tk.Tk()
    root.title("Menú de Métodos de Impresión de Prueba")
    root.geometry("420x220")

    frame = ttk.Frame(root, padding=20)
    frame.pack(fill=tk.BOTH, expand=True)

    # Selección de método
    ttk.Label(frame, text="Selecciona el método de impresión de prueba:").pack(anchor=tk.W)
    metodo_var = tk.StringVar(value="win32ui (coordenadas absolutas)")
    metodos = [
        "win32ui (coordenadas absolutas)",
        "win32ui (alineado por espacios)",
        "win32ui (alineado con tabulaciones)",
        "win32ui (alineado con CRLF)"
    ]
    metodo_menu = ttk.Combobox(frame, textvariable=metodo_var, values=metodos, state="readonly")
    metodo_menu.pack(fill=tk.X, pady=5)

    # Selección de impresora disponible (offline o desconectada también)
    ttk.Label(frame, text="Nombre de la impresora:").pack(anchor=tk.W, pady=(10,0))
    def obtener_impresoras():
        if win32print:
            flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            try:
                printers = win32print.EnumPrinters(flags)
                return [p[2] for p in printers]
            except Exception:
                return []
        return []

    printer_var = tk.StringVar()
    printer_list = obtener_impresoras()
    if printer_list:
        printer_var.set(printer_list[0])
    else:
        printer_var.set("EPSON TM-U950")
    printer_menu = ttk.Combobox(frame, textvariable=printer_var, values=printer_list, state="readonly")
    printer_menu.pack(fill=tk.X, pady=5)

    # Botón de imprimir
    imprimir_btn = ttk.Button(frame, text="Imprimir factura de prueba", command=imprimir_segun_metodo)
    imprimir_btn.pack(pady=15)

    root.mainloop()
