import os
import tempfile
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

try:
    from escpos.printer import Serial
except Exception:
    Serial = None

# Comandos de la EPSON TM-U950 para seleccionar el modo de impresi\u00f3n.
# SLIP4_MODE indica que se utilizará la posición "slip 4" recomendada
# para facturas en la bandeja de formularios.
SLIP4_MODE = b"\x1B\x69\x04"

try:
    from reportlab.lib.pagesizes import LETTER
    from reportlab.pdfgen import canvas
except Exception:
    canvas = None

def coordenada_a_texto_raw(prev_y_cm, x_cm, y_cm, texto, ancho_char_cm=0.25, alto_linea_cm=0.40):
    """Convierte coordenadas en centimetros a texto usando saltos de linea.

    ``prev_y_cm`` es la última coordenada Y impresa. El valor se actualiza al
    devolver el texto formateado para que las siguientes llamadas puedan
    calcular la diferencia de líneas correctamente.
    """
    espacios = int(x_cm / ancho_char_cm)
    saltos = int((y_cm - prev_y_cm) / alto_linea_cm)
    return ("\n" * saltos) + (" " * espacios) + texto + "\n", y_cm

def cm_a_twips(valor_cm: float) -> int:
    """Convierte cent\u00edmetros a TWIPS (1/1440 pulgadas)."""
    return round(valor_cm * 566.93)

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
        "literal": "Cuatro d\u00f3lares con cincuenta centavos",
        "sumas": "3.55",
        "iva": "0.46",
        "subtotal": "4.01",
        "exentas": "0.00",
        "no_sujetas": "0.00",
        "descuentos": "0.00",
        "total": "4.01",
    }

    return encabezado, productos, totales

def imprimir_factura_raw(printer_name):
    """Imprime una factura de prueba directamente en la impresora RAW."""
    if win32print is None:
        messagebox.showerror("Error", "win32print no está instalado. Instala pywin32.")
        return
    encabezado, productos, totales = generar_factura_datos()

    factura_raw = ""
    prev_y_cm = 0.0

    # Encabezado (solo datos, sin etiquetas)
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
        line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, x, y, text)
        factura_raw += line

    # Tabla de productos
    y_base = 10.10
    row_height = 0.6
    for i, (cantidad, descripcion, precio, exentas, no_sujetas, gravadas) in enumerate(productos):
        y = y_base + i * row_height
        line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 2.22, y, cantidad)
        factura_raw += line
        line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 3.90, y, descripcion)
        factura_raw += line
        line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 9.21, y, precio)
        factura_raw += line
        line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 11.11, y, exentas)
        factura_raw += line
        line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 12.70, y, no_sujetas)
        factura_raw += line
        line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 14.10, y, gravadas)
        factura_raw += line
    # Totales y resumen fiscal
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
        totales["exentas"],
        totales["no_sujetas"],
        totales["descuentos"],
        totales["total"],
    ]
    for (x, y), text in zip(totals_pos, totals_vals):
        line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, x, y, text)
        factura_raw += line
    # Forzar impresi\u00f3n en modo slip 4 para que la impresora utilice
    # la bandeja de formularios en lugar del recibo continuo.
    slip_cmd = SLIP4_MODE
    try:
        hprinter = win32print.OpenPrinter(printer_name)
        win32print.StartDocPrinter(hprinter, 1, ("Factura RAW", None, "RAW"))
        win32print.StartPagePrinter(hprinter)
        win32print.WritePrinter(hprinter, slip_cmd)
        win32print.WritePrinter(hprinter, factura_raw.encode("utf-8"))
        win32print.EndPagePrinter(hprinter)
        win32print.EndDocPrinter(hprinter)
        win32print.ClosePrinter(hprinter)
        messagebox.showinfo("Éxito", "Factura enviada a la impresora.")
    except Exception as e:
        messagebox.showerror("Error de impresión", str(e))

def imprimir_factura_raw_simple(printer_name):
    """Imprime una factura de prueba en texto plano, sin coordenadas, solo alineado por espacios."""
    if win32print is None:
        messagebox.showerror("Error", "win32print no está instalado. Instala pywin32.")
        return
    encabezado, productos, totales = generar_factura_datos()

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
    lines.append(f"IVA: {totales['iva']}")
    lines.append(f"Subtotal: {totales['subtotal']}")
    lines.append(f"Exentas: {totales['exentas']}")
    lines.append(f"No sujetas: {totales['no_sujetas']}")
    lines.append(f"Descuentos: {totales['descuentos']}")
    lines.append(f"Total: {totales['total']}")
    factura = "\n".join(lines) + "\n"
    slip_cmd = SLIP4_MODE
    try:
        hprinter = win32print.OpenPrinter(printer_name)
        win32print.StartDocPrinter(hprinter, 1, ("Factura RAW Simple", None, "RAW"))
        win32print.StartPagePrinter(hprinter)
        win32print.WritePrinter(hprinter, slip_cmd)
        win32print.WritePrinter(hprinter, factura.encode("utf-8"))
        win32print.EndPagePrinter(hprinter)
        win32print.EndDocPrinter(hprinter)
        win32print.ClosePrinter(hprinter)
        messagebox.showinfo("Éxito", "Factura enviada a la impresora.")
    except Exception as e:
        messagebox.showerror("Error de impresión", str(e))

def imprimir_factura_raw_tabs(printer_name):
    """Imprime una factura de prueba usando tabulaciones para alinear columnas."""
    if win32print is None:
        messagebox.showerror("Error", "win32print no está instalado. Instala pywin32.")
        return
    encabezado, productos, totales = generar_factura_datos()

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
    lines.append(f"IVA:\t{totales['iva']}")
    lines.append(f"Subtotal:\t{totales['subtotal']}")
    lines.append(f"Exentas:\t{totales['exentas']}")
    lines.append(f"No sujetas:\t{totales['no_sujetas']}")
    lines.append(f"Descuentos:\t{totales['descuentos']}")
    lines.append(f"Total:\t{totales['total']}")
    factura = "\n".join(lines) + "\n"
    slip_cmd = SLIP4_MODE
    try:
        hprinter = win32print.OpenPrinter(printer_name)
        win32print.StartDocPrinter(hprinter, 1, ("Factura RAW Tabs", None, "RAW"))
        win32print.StartPagePrinter(hprinter)
        win32print.WritePrinter(hprinter, slip_cmd)
        win32print.WritePrinter(hprinter, factura.encode("utf-8"))
        win32print.EndPagePrinter(hprinter)
        win32print.EndDocPrinter(hprinter)
        win32print.ClosePrinter(hprinter)
        messagebox.showinfo("Éxito", "Factura enviada a la impresora.")
    except Exception as e:
        messagebox.showerror("Error de impresión", str(e))

def imprimir_factura_raw_crlf(printer_name):
    """Imprime una factura de prueba usando saltos de línea CRLF (\r\n)."""
    if win32print is None:
        messagebox.showerror("Error", "win32print no está instalado. Instala pywin32.")
        return
    encabezado, productos, totales = generar_factura_datos()

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
        lines.append(f"{cant: <5}{desc: <23}{prec: >7}    {ex:>4}     {ns:>4}   {grav:>4}")
    lines.append("")
    lines.append(totales["literal"])
    lines.append(f"Sumas: {totales['sumas']}")
    lines.append(f"IVA: {totales['iva']}")
    lines.append(f"Subtotal: {totales['subtotal']}")
    lines.append(f"Exentas: {totales['exentas']}")
    lines.append(f"No sujetas: {totales['no_sujetas']}")
    lines.append(f"Descuentos: {totales['descuentos']}")
    lines.append(f"Total: {totales['total']}")
    factura = "\r\n".join(lines) + "\r\n"
    slip_cmd = SLIP4_MODE
    try:
        hprinter = win32print.OpenPrinter(printer_name)
        win32print.StartDocPrinter(hprinter, 1, ("Factura RAW CRLF", None, "RAW"))
        win32print.StartPagePrinter(hprinter)
        win32print.WritePrinter(hprinter, slip_cmd)
        win32print.WritePrinter(hprinter, factura.encode("utf-8"))
        win32print.EndPagePrinter(hprinter)
        win32print.EndDocPrinter(hprinter)
        win32print.ClosePrinter(hprinter)
        messagebox.showinfo("Éxito", "Factura enviada a la impresora.")
    except Exception as e:
        messagebox.showerror("Error de impresión", str(e))

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
        if win32con:
            dc.SetMapMode(win32con.MM_TWIPS)
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

        # Productos
        y_base = 10.10
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
            totales["exentas"],
            totales["no_sujetas"],
            totales["descuentos"],
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

def imprimir_factura_os_startfile(printer_name):
    """Genera un txt temporal y lo imprime con os.startfile."""
    if not activar_modo_slip(printer_name):
        return
    factura = (
        "Factura de prueba\n"
        "Artículo 1\t1.00\n"
        "Artículo 2\t2.00\n"
    )
    with tempfile.NamedTemporaryFile(delete=False, suffix=".txt", mode="w", encoding="utf-8") as tmp:
        tmp.write(factura)
    try:
        os.startfile(tmp.name, "print")
        messagebox.showinfo("Éxito", "Factura enviada a la impresora predeterminada.")
    except Exception as e:
        messagebox.showerror("Error de impresión", str(e))

def imprimir_factura_pdf(printer_name):
    """Genera un PDF con reportlab y lo imprime."""
    if canvas is None:
        messagebox.showerror("Error", "reportlab no está instalado.")
        return
    if not activar_modo_slip(printer_name):
        return
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    c = canvas.Canvas(tmp.name, pagesize=LETTER)
    c.drawString(72, 720, "Factura de prueba")
    c.drawString(72, 700, "Artículo 1 - 1.00")
    c.showPage()
    c.save()
    try:
        os.startfile(tmp.name, "print")
        messagebox.showinfo("Éxito", "PDF enviado a la impresora.")
    except Exception as e:
        messagebox.showerror("Error de impresión", str(e))

def imprimir_factura_escpos(printer_name):
    """Imprime usando comandos ESC/POS si el driver está disponible."""
    if Serial is None:
        messagebox.showerror("Error", "Librería escpos no disponible.")
        return
    if not activar_modo_slip(printer_name):
        return
    try:
        p = Serial()  # Configuración por defecto /dev/ttyS0
        p.text("Factura de prueba\n")
        p.text("Articulo 1\t1.00\n")
        p.cut()
        messagebox.showinfo("Éxito", "Factura enviada por ESC/POS.")
    except Exception as e:
        messagebox.showerror("Error de impresión", str(e))

def vista_previa_factura():
    """Muestra una ventana con el texto de la factura generada (RAW coordenadas)."""
    encabezado, productos, totales = generar_factura_datos()

    factura_raw = ""
    prev_y_cm = 0.0

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
        line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, x, y, text)
        factura_raw += line

    y_base = 10.10
    row_height = 0.6
    for i, (cantidad, descripcion, precio, exentas, no_sujetas, gravadas) in enumerate(productos):
        y = y_base + i * row_height
        line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 2.22, y, cantidad)
        factura_raw += line
        line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 3.90, y, descripcion)
        factura_raw += line
        line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 9.21, y, precio)
        factura_raw += line
        line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 11.11, y, exentas)
        factura_raw += line
        line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 12.70, y, no_sujetas)
        factura_raw += line
        line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 14.10, y, gravadas)
        factura_raw += line
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
        totales["exentas"],
        totales["no_sujetas"],
        totales["descuentos"],
        totales["total"],
    ]
    for (x, y), text in zip(totals_pos, totals_vals):
        line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, x, y, text)
        factura_raw += line
    preview = tk.Toplevel()
    preview.title("Vista previa de factura (RAW coordenadas)")
    text = tk.Text(preview, wrap="none", font=("Courier New", 10))
    text.insert("1.0", factura_raw)
    text.pack(expand=True, fill=tk.BOTH)
    preview.geometry("800x600")

def imprimir_segun_metodo():
    metodo = metodo_var.get()
    printer_name = printer_var.get()
    if not printer_name:
        messagebox.showwarning("Advertencia", "Debes ingresar el nombre de la impresora.")
        return
    if metodo == "RAW (coordenadas y win32print)":
        imprimir_factura_raw(printer_name)
    elif metodo == "RAW simple (alineado por espacios)":
        imprimir_factura_raw_simple(printer_name)
    elif metodo == "RAW con tabulaciones (\t)":
        imprimir_factura_raw_tabs(printer_name)
    elif metodo == "RAW con CRLF (\r\n)":
        imprimir_factura_raw_crlf(printer_name)
    elif metodo == "win32ui (coordenadas absolutas)":
        imprimir_factura_win32ui(printer_name)
    elif metodo == "os.startfile (predeterminado)":
        imprimir_factura_os_startfile(printer_name)
    elif metodo == "ReportLab a PDF":
        imprimir_factura_pdf(printer_name)
    elif metodo == "ESC/POS":
        imprimir_factura_escpos(printer_name)
    elif metodo == "Vista previa (solo mostrar)":
        vista_previa_factura()
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
    metodo_var = tk.StringVar(value="RAW (coordenadas y win32print)")
    # Ordenados según la probabilidad de funcionar con EPSON TM-U950
    metodos = [
        "RAW (coordenadas y win32print)",
        "RAW con CRLF (\\r\\n)",
        "RAW con tabulaciones (\\t)",
        "RAW simple (alineado por espacios)",
        "win32ui (coordenadas absolutas)",
        "os.startfile (predeterminado)",
        "ReportLab a PDF",
        "ESC/POS",
        "Vista previa (solo mostrar)"
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
