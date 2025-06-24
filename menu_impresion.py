import tkinter as tk
from tkinter import ttk, messagebox

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

def configurar_mapeo(dc):
    """Configura el mapeo para que 27.5 cm x 16.6 cm coincidan con el área imprimible."""
    if not win32con:
        return
    dc.SetMapMode(win32con.MM_TWIPS)
    ancho = dc.GetDeviceCaps(win32con.HORZRES)
    alto = dc.GetDeviceCaps(win32con.VERTRES)
    dc.SetWindowExtEx(cm_a_twips(27.5), cm_a_twips(16.6))
    dc.SetViewportExtEx(ancho, -alto)

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
        "cliente": "Francisco L\u00f3pez",
        "direccion": "Col. Escal\u00f3n, San Salvador",
        "fecha": "2025-06-12",
        "giro": "Comercio",
        "vence": "2025-06-10",
        "pago": "30 D\u00cdAS",
        "atencion": "Mar\u00eda P\u00e9rez",
        "nrc": "123456-7",
        "remision": "REM-00123",
        "nit": "0614-250786-102-3",
        "orden": "ORD-789",
        "proveedor": "Distribuidora S.A.",
        "fecha_doc": "2025-05-30",
    }

    productos = [
        ("2", "Paracetamol 500mg", "0.50", "0.00", "0.00", "1.00"),
        ("1", "Ibuprofeno 200mg", "0.75", "0.00", "0.00", "0.75"),
        ("3", "Vitamina C 1000mg", "0.60", "0.00", "0.00", "1.80"),
    ]

    totales = {
        "literal": "Cuatro dólares con cincuenta centavos",
        "sumas": "3.55",
        "iva": "0.46",
        "subtotal": "4.01",
        "iva_retenido": "0.00",
        "no_sujetas": "0.00",
        "ventas_exentas": "0.00",
        "total": "4.01",
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

        def draw(x_cm, y_cm, texto):
            dc.TextOut(cm_a_twips(x_cm), -cm_a_twips(y_cm), texto)

        # Encabezado (solo datos)
        header_pos = [
            (4.45, 4.80),
            (4.45, 5.40),
            (3.81, 6.40),
            (3.81, 6.90),
            (3.81, 7.50),
            (4.45, 8.00),
            (4.45, 8.50),
            (7.62, 6.40),
            (7.62, 7.50),
            (11.43, 6.40),
            (12.07, 7.50),
            (11.43, 8.00),
            (11.75, 8.50),
        ]
        header_vals = [
            encabezado["cliente"],
            encabezado["direccion"],
            encabezado["fecha"],
            encabezado["giro"],
            encabezado["vence"],
            encabezado["pago"],
            encabezado["atencion"],
            encabezado["nrc"],
            encabezado["remision"],
            encabezado["nit"],
            encabezado["orden"],
            encabezado["proveedor"],
            encabezado["fecha_doc"],
        ]
        for (x, y), text in zip(header_pos, header_vals):
            draw(x, y, text)

        # Encabezado de la tabla de productos
        draw(2.22, 10.10, "Cantidad")
        draw(3.90, 10.10, "Descripción")
        draw(9.21, 10.10, "Precio unitario")
        draw(11.11, 10.10, "Ventas exentas")
        draw(12.70, 10.10, "Ventas no sujetas")
        draw(14.10, 10.10, "Ventas gravadas")

        # Productos
        y_base = 10.70
        row_height = 0.6
        for i, (cant, desc, prec, ex, ns, grav) in enumerate(productos):
            y = y_base + i * row_height
            draw(2.22, y, cant)
            draw(3.90, y, desc)
            draw(9.21, y, prec)
            draw(11.11, y, ex)
            draw(12.70, y, ns)
            draw(14.10, y, grav)

        # Totales
        totals_pos = [
            (2.22, 22.23),
            (14.10, 21.59),
            (14.10, 22.23),
            (14.10, 22.86),
            (14.10, 23.45),
            (14.10, 24.00),
            (14.10, 24.60),
            (14.10, 25.08),
        ]
        totals_vals = [
            totales["literal"],
            totales["sumas"],
            totales["iva"],
            totales["subtotal"],
            totales["iva_retenido"],
            totales["no_sujetas"],
            totales["ventas_exentas"],
            totales["total"],
        ]
        for (x, y), text in zip(totals_pos, totals_vals):
            draw(x, y, text)

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

        def draw(x_cm, y_cm, texto):
            dc.TextOut(cm_a_twips(x_cm), -cm_a_twips(y_cm), texto)

        header_order = [
            "cliente",
            "direccion",
            "fecha",
            "giro",
            "vence",
            "pago",
            "atencion",
            "nrc",
            "remision",
            "nit",
            "orden",
            "proveedor",
            "fecha_doc",
        ]


        lines = [encabezado[campo] for campo in header_order]
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
            draw(0, y, encabezado[campo])
            y += line_height

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
        draw(0, y, f"Exentas: {totales['exentas']}")
        y += line_height
        draw(0, y, f"No sujetas: {totales['no_sujetas']}")
        y += line_height
        draw(0, y, f"Descuentos: {totales['descuentos']}")
        y += line_height
        draw(0, y, f"Total: {totales['total']}")

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

        def draw(x_cm, y_cm, texto):
            dc.TextOut(cm_a_twips(x_cm), -cm_a_twips(y_cm), texto)

        header_order = [
            "cliente",
            "direccion",
            "fecha",
            "giro",
            "vence",
            "pago",
            "atencion",
            "nrc",
            "remision",
            "nit",
            "orden",
            "proveedor",
            "fecha_doc",
        ]


        lines = [encabezado[campo] for campo in header_order]
        lines.append("")
        lines.append("Cant\tDescripción\t\t\tPrecio\tExentas\tNoSuj\tGravadas")
        for cant, desc, prec, ex, ns, grav in productos:
            lines.append(f"{cant}\t{desc}\t\t{prec}\t{ex}\t{ns}\t{grav}")
        lines.append("")
        lines.append(totales["literal"])
        lines.append(f"Sumas:\t{totales['sumas']}")
        lines.append(f"13% IVA:\t{totales['iva']}")
        lines.append(f"Subtotal:\t{totales['subtotal']}")
        lines.append(f"IVA retenido:\t{totales['iva_retenido']}")
        lines.append(f"Vtas no sujetas:\t{totales['no_sujetas']}")
        lines.append(f"Ventas exentas:\t{totales['ventas_exentas']}")
        lines.append(f"Venta total:\t{totales['total']}")

        y = 4.8
        line_height = 0.6
        for campo in header_order:
            draw(0, y, encabezado[campo])
            y += line_height

        draw(0, 10.10, "Cant\tDescripción\t\t\tPrecio\tExentas\tNoSuj\tGravadas")

        y_base = 10.70
        for i, (cant, desc, prec, ex, ns, grav) in enumerate(productos):
            y_line = y_base + i * line_height
            draw(0, y_line, f"{cant}\t{desc}\t\t{prec}\t{ex}\t{ns}\t{grav}")

        y = y_base + len(productos) * line_height + line_height
        draw(0, y, totales["literal"])
        y += line_height
        draw(0, y, f"Sumas:\t{totales['sumas']}")
        y += line_height
        draw(0, y, f"IVA:\t{totales['iva']}")
        y += line_height
        draw(0, y, f"Subtotal:\t{totales['subtotal']}")
        y += line_height
        draw(0, y, f"Exentas:\t{totales['exentas']}")
        y += line_height
        draw(0, y, f"No sujetas:\t{totales['no_sujetas']}")
        y += line_height
        draw(0, y, f"Descuentos:\t{totales['descuentos']}")
        y += line_height
        draw(0, y, f"Total:\t{totales['total']}")

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

        def draw(x_cm, y_cm, texto):
            dc.TextOut(cm_a_twips(x_cm), -cm_a_twips(y_cm), texto)

        header_order = [
            "cliente",
            "direccion",
            "fecha",
            "giro",
            "vence",
            "pago",
            "atencion",
            "nrc",
            "remision",
            "nit",
            "orden",
            "proveedor",
            "fecha_doc",
        ]

        lines = [encabezado[campo] for campo in header_order]
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
            draw(0, y, encabezado[campo])
            y += line_height

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
        draw(0, y, f"Exentas: {totales['exentas']}")
        y += line_height
        draw(0, y, f"No sujetas: {totales['no_sujetas']}")
        y += line_height
        draw(0, y, f"Descuentos: {totales['descuentos']}")
        y += line_height
        draw(0, y, f"Total: {totales['total']}")

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
