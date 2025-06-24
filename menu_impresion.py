import tkinter as tk
from tkinter import ttk, messagebox

# Intenta importar win32print, si no está disponible muestra un error al intentar imprimir
try:
    import win32print
except ImportError:
    win32print = None

def coordenada_a_texto_raw(prev_y_cm, x_cm, y_cm, texto, ancho_char_cm=0.25, alto_linea_cm=0.40):
    """Convierte coordenadas en centimetros a texto usando saltos de linea.

    ``prev_y_cm`` es la última coordenada Y impresa. El valor se actualiza al
    devolver el texto formateado para que las siguientes llamadas puedan
    calcular la diferencia de líneas correctamente.
    """
    espacios = int(x_cm / ancho_char_cm)
    saltos = int((y_cm - prev_y_cm) / alto_linea_cm)
    return ("\n" * saltos) + (" " * espacios) + texto + "\n", y_cm

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
    SLIP_MODE = b"\x1B\x69"  # ESC i
    try:
        hprinter = win32print.OpenPrinter(printer_name)
        win32print.StartDocPrinter(hprinter, 1, ("Factura RAW", None, "RAW"))
        win32print.StartPagePrinter(hprinter)
        win32print.WritePrinter(hprinter, SLIP_MODE + factura_raw.encode("utf-8"))
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
    SLIP_MODE = b"\x1B\x69"
    try:
        hprinter = win32print.OpenPrinter(printer_name)
        win32print.StartDocPrinter(hprinter, 1, ("Factura RAW Simple", None, "RAW"))
        win32print.StartPagePrinter(hprinter)
        win32print.WritePrinter(hprinter, SLIP_MODE + factura.encode("utf-8"))
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
    SLIP_MODE = b"\x1B\x69"
    try:
        hprinter = win32print.OpenPrinter(printer_name)
        win32print.StartDocPrinter(hprinter, 1, ("Factura RAW Tabs", None, "RAW"))
        win32print.StartPagePrinter(hprinter)
        win32print.WritePrinter(hprinter, SLIP_MODE + factura.encode("utf-8"))
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
    SLIP_MODE = b"\x1B\x69"
    try:
        hprinter = win32print.OpenPrinter(printer_name)
        win32print.StartDocPrinter(hprinter, 1, ("Factura RAW CRLF", None, "RAW"))
        win32print.StartPagePrinter(hprinter)
        win32print.WritePrinter(hprinter, SLIP_MODE + factura.encode("utf-8"))
        win32print.EndPagePrinter(hprinter)
        win32print.EndDocPrinter(hprinter)
        win32print.ClosePrinter(hprinter)
        messagebox.showinfo("Éxito", "Factura enviada a la impresora.")
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
    printer_name = printer_entry.get()
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
    elif metodo == "Vista previa (solo mostrar)":
        vista_previa_factura()
    else:
        messagebox.showinfo("Info", f"Método '{metodo}' aún no implementado.")

root = tk.Tk()
root.title("Menú de Métodos de Impresión de Prueba")
root.geometry("420x220")

frame = ttk.Frame(root, padding=20)
frame.pack(fill=tk.BOTH, expand=True)

# Selección de método
ttk.Label(frame, text="Selecciona el método de impresión de prueba:").pack(anchor=tk.W)
metodo_var = tk.StringVar(value="RAW (coordenadas y win32print)")
metodos = [
    "RAW (coordenadas y win32print)",
    "RAW simple (alineado por espacios)",
    "RAW con tabulaciones (\t)",
    "RAW con CRLF (\r\n)",
    "Vista previa (solo mostrar)"
]
metodo_menu = ttk.Combobox(frame, textvariable=metodo_var, values=metodos, state="readonly")
metodo_menu.pack(fill=tk.X, pady=5)

# Entrada de nombre de impresora
ttk.Label(frame, text="Nombre de la impresora:").pack(anchor=tk.W, pady=(10,0))
printer_entry = ttk.Entry(frame)
printer_entry.pack(fill=tk.X, pady=5)
printer_entry.insert(0, "EPSON TM-U950")

# Botón de imprimir
imprimir_btn = ttk.Button(frame, text="Imprimir factura de prueba", command=imprimir_segun_metodo)
imprimir_btn.pack(pady=15)

root.mainloop()
