import tkinter as tk
from tkinter import ttk

# Colors
BACKGROUND = '#6a0dad'  # dark purple
FOREGROUND = '#ffffff'  # white text
BUTTON_BG = '#9932cc'   # medium purple
BUTTON_FG = '#ffffff'

LOVE_LETTER = (
    "Querida Luna,\n\n"
    "Desde el primer instante en que te vi supe que mi vida cambiar\u00eda para siempre. "
    "Tus ojos reflejan la ternura de las estrellas y tu sonrisa ilumina mis ma\u00f1anas. "
    "Espero que esta peque\u00f1a carta sea capaz de transmitir todo el cari\u00f1o que siento por ti.\n\n"
    "Con amor eterno,\n"
    "Tu Sol"
)

class CartaAmor(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Carta de Amor")
        self.configure(bg=BACKGROUND)
        self._build_widgets()

    def _build_widgets(self):
        # T\u00edtulo
        title = tk.Label(self, text="Mi Carta para Ti", bg=BACKGROUND, fg=FOREGROUND,
                         font=("Helvetica", 20, "bold"))
        title.pack(pady=(20, 10))

        # Marco para el texto de la carta
        frame = tk.Frame(self, bg=BACKGROUND)
        frame.pack(padx=20, pady=10)

        text = tk.Text(frame, width=60, height=10, wrap="word", bg=BACKGROUND,
                       fg=FOREGROUND, font=("Helvetica", 12), borderwidth=0)
        text.insert("1.0", LOVE_LETTER)
        text.config(state="disabled")
        text.pack()

        # Bot\u00f3n interactivo
        self.name_var = tk.StringVar()
        entry = ttk.Entry(self, textvariable=self.name_var)
        entry.pack(pady=(10, 5))
        entry.insert(0, "Ingresa tu nombre")

        button = tk.Button(self, text="Responder", command=self.show_message,
                           bg=BUTTON_BG, fg=BUTTON_FG, relief=tk.FLAT)
        button.pack(pady=(5, 20))

    def show_message(self):
        name = self.name_var.get().strip() or "An\u00f3nimo"
        message = f"Gracias por leer mi carta, {name}!" 
        tk.messagebox.showinfo("Respuesta", message)

if __name__ == "__main__":
    app = CartaAmor()
    app.mainloop()
