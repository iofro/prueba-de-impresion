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
# SLIP_MODE habilita la bandeja de formularios y SLIP4_MODE indica que
# se utilizar\u00e1 la posici\u00f3n "slip 4" recomendada para facturas.
SLIP_MODE = b"\x1B\x69"
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
    return int(valor_cm * 567)

def imprimir_factura_raw(printer_name):
    """Imprime una factura de prueba directamente en la impresora RAW."""
    if win32print is None:
        messagebox.showerror("Error", "win32print no está instalado. Instala pywin32.")
        return
    factura_raw = ""
    prev_y_cm = 0.0
    # Encabezado (solo datos, sin etiquetas)
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 4.45, 4.80, "Francisco López")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 4.45, 5.40, "Col. Escalón, San Salvador")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 3.81, 6.40, "2025-06-12")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 3.81, 6.90, "Comercio")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 3.81, 7.50, "2025-06-10")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 4.45, 8.00, "30 DÍAS")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 4.45, 8.50, "María Pérez")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 7.62, 6.40, "123456-7")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 7.62, 7.50, "REM-00123")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 11.43, 6.40, "0614-250786-102-3")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 12.07, 7.50, "ORD-789")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 11.43, 8.00, "Distribuidora S.A.")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 11.75, 8.50, "2025-05-30")
    factura_raw += line
    # Tabla de productos
    productos = [
        ("2", "Paracetamol 500mg", "0.50", "0.00", "0.00", "1.00"),
        ("1", "Ibuprofeno 200mg", "0.75", "0.00", "0.00", "0.75"),
        ("3", "Vitamina C 1000mg", "0.60", "0.00", "0.00", "1.80"),
    ]
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
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 2.22, 22.23, "Cuatro dólares con cincuenta centavos")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 14.10, 21.59, "3.55")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 14.10, 22.23, "0.46")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 14.10, 22.86, "4.01")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 14.10, 23.45, "0.00")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 14.10, 24.00, "0.00")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 14.10, 24.60, "0.00")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 14.10, 25.08, "4.01")
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
    factura = """Francisco López\nCol. Escalón, San Salvador\n2025-06-12\nComercio\n2025-06-10\n30 DÍAS\nMaría Pérez\n123456-7\nREM-00123\n0614-250786-102-3\nORD-789\nDistribuidora S.A.\n2025-05-30\n\nCant  Descripción             Precio  Exentas  NoSuj  Gravadas\n2     Paracetamol 500mg       0.50    0.00     0.00   1.00\n1     Ibuprofeno 200mg        0.75    0.00     0.00   0.75\n3     Vitamina C 1000mg       0.60    0.00     0.00   1.80\n\nCuatro dólares con cincuenta centavos\nSumas: 3.55\nIVA: 0.46\nSubtotal: 4.01\nExentas: 0.00\nNo sujetas: 0.00\nDescuentos: 0.00\nTotal: 4.01\n"""
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
    factura = (
        "Francisco López\nCol. Escalón, San Salvador\n2025-06-12\nComercio\n2025-06-10\n30 DÍAS\nMaría Pérez\n123456-7\nREM-00123\n0614-250786-102-3\nORD-789\nDistribuidora S.A.\n2025-05-30\n\n"
        "Cant\tDescripción\t\t\tPrecio\tExentas\tNoSuj\tGravadas\n"
        "2\tParacetamol 500mg\t\t0.50\t0.00\t0.00\t1.00\n"
        "1\tIbuprofeno 200mg\t\t0.75\t0.00\t0.00\t0.75\n"
        "3\tVitamina C 1000mg\t\t0.60\t0.00\t0.00\t1.80\n\n"
        "Cuatro dólares con cincuenta centavos\nSumas:\t3.55\nIVA:\t0.46\nSubtotal:\t4.01\nExentas:\t0.00\nNo sujetas:\t0.00\nDescuentos:\t0.00\nTotal:\t4.01\n"
    )
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
    factura = (
        "Francisco López\r\nCol. Escalón, San Salvador\r\n2025-06-12\r\nComercio\r\n2025-06-10\r\n30 DÍAS\r\nMaría Pérez\r\n123456-7\r\nREM-00123\r\n0614-250786-102-3\r\nORD-789\r\nDistribuidora S.A.\r\n2025-05-30\r\n\r\n"
        "Cant  Descripción             Precio  Exentas  NoSuj  Gravadas\r\n"
        "2     Paracetamol 500mg       0.50    0.00     0.00   1.00\r\n"
        "1     Ibuprofeno 200mg        0.75    0.00     0.00   0.75\r\n"
        "3     Vitamina C 1000mg       0.60    0.00     0.00   1.80\r\n\r\n"
        "Cuatro dólares con cincuenta centavos\r\nSumas: 3.55\r\nIVA: 0.46\r\nSubtotal: 4.01\r\nExentas: 0.00\r\nNo sujetas: 0.00\r\nDescuentos: 0.00\r\nTotal: 4.01\r\n"
    )
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
        draw(4.45, 4.80, "Francisco L\u00f3pez")
        draw(4.45, 5.40, "Col. Escal\u00f3n, San Salvador")
        draw(3.81, 6.40, "2025-06-12")
        draw(3.81, 6.90, "Comercio")
        draw(3.81, 7.50, "2025-06-10")
        draw(4.45, 8.00, "30 D\u00cdAS")
        draw(4.45, 8.50, "Mar\u00eda P\u00e9rez")
        draw(7.62, 6.40, "123456-7")
        draw(7.62, 7.50, "REM-00123")
        draw(11.43, 6.40, "0614-250786-102-3")
        draw(12.07, 7.50, "ORD-789")
        draw(11.43, 8.00, "Distribuidora S.A.")
        draw(11.75, 8.50, "2025-05-30")

        # Productos
        productos = [
            ("2", "Paracetamol 500mg", "0.50", "0.00", "0.00", "1.00"),
            ("1", "Ibuprofeno 200mg", "0.75", "0.00", "0.00", "0.75"),
            ("3", "Vitamina C 1000mg", "0.60", "0.00", "0.00", "1.80"),
        ]
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
        draw(2.22, 22.23, "Cuatro d\u00f3lares con cincuenta centavos")
        draw(14.10, 21.59, "3.55")
        draw(14.10, 22.23, "0.46")
        draw(14.10, 22.86, "4.01")
        draw(14.10, 23.45, "0.00")
        draw(14.10, 24.00, "0.00")
        draw(14.10, 24.60, "0.00")
        draw(14.10, 25.08, "4.01")

        dc.EndPage()
        dc.EndDoc()
        dc.DeleteDC()
        win32print.ClosePrinter(hprinter)
        messagebox.showinfo("Éxito", "Factura enviada a la impresora.")
    except Exception as e:
        messagebox.showerror("Error de impresión", str(e))

def imprimir_factura_os_startfile(printer_name):
    """Genera un txt temporal y lo imprime con os.startfile."""
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
    factura_raw = ""
    prev_y_cm = 0.0
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 4.45, 4.80, "Francisco López")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 4.45, 5.40, "Col. Escalón, San Salvador")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 3.81, 6.40, "2025-06-12")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 3.81, 6.90, "Comercio")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 3.81, 7.50, "2025-06-10")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 4.45, 8.00, "30 DÍAS")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 4.45, 8.50, "María Pérez")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 7.62, 6.40, "123456-7")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 7.62, 7.50, "REM-00123")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 11.43, 6.40, "0614-250786-102-3")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 12.07, 7.50, "ORD-789")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 11.43, 8.00, "Distribuidora S.A.")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 11.75, 8.50, "2025-05-30")
    factura_raw += line
    productos = [
        ("2", "Paracetamol 500mg", "0.50", "0.00", "0.00", "1.00"),
        ("1", "Ibuprofeno 200mg", "0.75", "0.00", "0.00", "0.75"),
        ("3", "Vitamina C 1000mg", "0.60", "0.00", "0.00", "1.80"),
    ]
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
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 2.22, 22.23, "Cuatro dólares con cincuenta centavos")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 14.10, 21.59, "3.55")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 14.10, 22.23, "0.46")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 14.10, 22.86, "4.01")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 14.10, 23.45, "0.00")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 14.10, 24.00, "0.00")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 14.10, 24.60, "0.00")
    factura_raw += line
    line, prev_y_cm = coordenada_a_texto_raw(prev_y_cm, 14.10, 25.08, "4.01")
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
