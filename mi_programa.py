import tkinter as tk
from tkinter import ttk, filedialog, messagebox  # Import filedialog explicitly
import pandas as pd
from tkcalendar import Calendar, DateEntry
import uuid # Para generar un id unico para cada campo de texto 
import re
import os  # Import os for file checks
from datetime import datetime  # Import datetime for timestamp

combo_width_limit = 100
entry_width_limit = 75  # Define the maximum width limit for the Entry widget

# ----------------------------------------------------------------------------
# ----------------------------------------------------------------------------
# Funciones para eventos de UI
# Función para mostrar la respuesta generada
def generar_respuesta():
    motivo_seleccionado = combo_motivo.get()
    if not motivo_seleccionado or motivo_seleccionado not in motivos:
        messagebox.showerror("Error", "Por favor selecciona un motivo válido.")
        return

    tema_seleccionado = combo_tema.get()
    if not tema_seleccionado or tema_seleccionado not in motivos[motivo_seleccionado]:
        messagebox.showerror("Error", "Por favor selecciona un tema válido.")
        return

    respuesta_label.config(text=f"Respuesta generada para {motivo_seleccionado} - {tema_seleccionado}")
    
    # Buscar la respuesta en el diccionario de acuerdo al motivo y tema seleccionados
    respuesta = motivos[motivo_seleccionado][tema_seleccionado][0]  # Assume the first response is used
    
    respuesta_text['state'] = 'normal'
    # Limpiar el campo de texto antes de mostrar la nueva respuesta
    respuesta_text.delete(1.0, tk.END)
    
    # Parse the respuesta string for placeholders (e.g., {name:type})
    #pattern = r"\{(\w+):(\w+)(?::([^}]+))?\}"  # Matches {name:type[:options]}
    pattern = r"\{(\w+):([^}:]+)(?::([^}]+))?\}"  # Matches {name:type[:options]}

    # Validate placeholders
    invalid_placeholders = re.findall(r"\{[^}]*\}", respuesta)  # Find all placeholders
    for placeholder in invalid_placeholders:
        if not re.match(pattern, placeholder):
            messagebox.showerror("Error", f"La respuesta en el archivo de motivos y respuestas está mal formateada: {placeholder}")
            return  # Stop processing if an invalid placeholder is found
    
    parts = re.split(pattern, respuesta)  # Split the string by the placeholder pattern
    
    i = 0
    while i < len(parts):
        if i % 4 == 0:
            # Insert plain text
            respuesta_text.insert(tk.END, parts[i])
            i += 1
        else:
            # Extract field name and type
            field_type = parts[i]
            #field_name = parts[i + 1].lower() + uuid.uuid4().hex  # Unique name for the widget
            field_name = uuid.uuid4().hex  # Unique name for the widget
            field_options = parts[i + 2] 
            if field_type == "text":
                # Insert a text Entry widget
                def adjust_entry_width(event, entry_widget):
                    # Adjust the width of the Entry widget based on the length of the text
                    new_width = min(entry_width_limit, max(20, len(entry_widget.get()) + 1))
                    entry_widget.config(width=new_width)

                entry = tk.Entry(respuesta_text, width=20, name=field_name, bg="#FFEE99")
                #entry = tk.Entry(respuesta_text, width=20, name=field_name)
                entry.bind("<KeyRelease>", lambda event, e=entry: adjust_entry_width(event, e))
                respuesta_text.window_create(tk.END, window=entry) 

            elif field_type == "date":
                # Insert a date picker button
                date_button = tk.Button(respuesta_text, text="Seleccionar Fecha", name=field_name,
                                        command=lambda b=field_name: open_date_picker(b))
                respuesta_text.window_create(tk.END, window=date_button)
            elif field_type == "month":
                # Insert a month picker button
                month_button = tk.Button(respuesta_text, text="Seleccionar Mes", name=field_name,
                                         command=lambda b=field_name: open_month_picker(b))
                respuesta_text.window_create(tk.END, window=month_button)
            elif field_type == "option":
                # Insert a combobox (dropdown)
                if field_options != None and field_options != "":
                    options = field_options.split("|")
                    max_option_length = min(combo_width_limit, max(len(option) for option in options))
                    combo = ttk.Combobox(respuesta_text, values=options, name=field_name, width=max_option_length, style="Custom.TCombobox")
                    # combo.state(["readonly"]) Se puede hacer que no se pueda editar el combobox
                    respuesta_text.window_create(tk.END, window=combo)
            else:
                messagebox.showerror("Error", f"Tipo de campo desconocido: {field_type}")
                return
            i += 3  # Skip the field name, type and options

    # Disable the text field to prevent editing
    respuesta_text['state'] = 'disabled'

def open_date_picker(button_name):
    # Create a top-level window for the calendar
    top = tk.Toplevel(root)
    top.title("Seleccionar Fecha")
    
    # Add a Calendar widget
    cal = Calendar(top, selectmode="day", date_pattern="dd/mm/yyyy")
    cal.pack(pady=10)
    
    # Add a button to confirm the date selection
    def select_date():
        selected_date = cal.get_date()
        #print("sale")
        #print(str(respuesta_text) + "." + button_name)
        button = root.nametowidget(str(respuesta_text) + "." + button_name)
        button.config(text=selected_date)  # Update the button text with the selected date
        top.destroy()  # Close the calendar window
    
    ttk.Button(top, text="Seleccionar", command=select_date).pack(pady=10)

def open_month_picker(button_name):
    # Create a top-level window for the calendar
    top = tk.Toplevel(root)
    top.title("Seleccionar Mes")
    
    # Add a Calendar widget with month selection mode
    #cal = Calendar(top, selectmode="month", date_pattern="mm/yyyy")
    cal = Calendar(top, selectmode="day", date_pattern="dd/mm/yyyy")
    cal.pack(pady=10)
    
    # Add a button to confirm the month selection
    def select_month():
        selected_month = cal.get_date()
        button = root.nametowidget(str(respuesta_text) + "." + button_name)
        button.config(text=selected_month[3:])  # Update the button text with the selected month
        top.destroy()  # Close the calendar window
    
    ttk.Button(top, text="Seleccionar", command=select_month).pack(pady=10)

# Función para actualizar los temas según el motivo seleccionado
def actualizar_temas(event):
    motivo_seleccionado = combo_motivo.get()
    combo_tema["values"] = list(motivos[motivo_seleccionado].keys())
    combo_tema.current(0)  # Seleccionar el primer tema por defecto

def update_status_bar(file_path):
    # Get the current time in hours and minutes
    current_time = datetime.now().strftime("%H:%M")
    # Update the status bar with the file path and time
    status_bar.config(text=f"Archivo cargado: {file_path} a las {current_time}")

def mostrar_acerca_de():
    messagebox.showinfo("Acerca de", "Este programa permite seleccionar motivos y temas para generar respuestas.")
    # Puedes agregar más información aquí, como la versión, autor, etc.

# ----------------------------------------------------------------------------
# ----------------------------------------------------------------------------
# Funciones para copiar la respuesta al portapapeles

def copiar_respuesta():
    # Retrieve all text from respuesta_text
    final_text = ""
    for index in respuesta_text.dump("1.0", tk.END):
        if index[0] == "text":
            # Append plain text
            final_text += index[1]
        elif index[0] == "window":
            # Retrieve the value from the embedded Entry widget
            widget = respuesta_text.nametowidget(index[1])
            if widget.winfo_class() == "Entry":
                # Append the value from the Entry widget
                final_text += widget.get()
            elif widget.winfo_class() == "Button":
                # Append the text from the Button widget (e.g., selected date)
                final_text += widget.cget("text")
            elif widget.winfo_class() == "TCombobox":
                # Append the selected value from the Combobox widget
                final_text += widget.get()
    

    # Copy the final text to the clipboard
    try:
        root.clipboard_clear()  # Clear the clipboard
        root.clipboard_append(final_text)  # Copy the final text to the clipboard
        root.update()  # Update the clipboard to make the content available
        messagebox.showinfo(message='Respuesta copiada al portapapeles')
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo copiar la respuesta al portapapeles: {e}")

# ----------------------------------------------------------------------------
# ----------------------------------------------------------------------------
# Funciones para generar el diccionario de motivos y respuestas desde el archivo de Excel

def fetch_temas(file_path):
    try:
        # Try to read the Excel file
        df = pd.read_excel(file_path)
        
        # Check if the required columns exist
        required_columns = ['TEMA', 'Submotivo', 'RESPUESTA ']
        for column in required_columns:
            if column not in df.columns:
                messagebox.showerror("Error", f"El archivo no contiene la columna requerida: {column}")
                return None
        
        # Check if the file is empty
        if df.empty:
            messagebox.showerror("Error", "El archivo está vacío.")
            return None

        # Create the dictionary of motivos, temas, and respuestas
        motivos = {}
        for index, row in df.iterrows():
            motivo = str(row['TEMA']).rstrip()
            tema = str(row['Submotivo']).rstrip()
            respuesta = row['RESPUESTA ']

            # Skip rows with missing data
            if pd.isna(motivo) or pd.isna(tema) or pd.isna(respuesta):
                continue

            if motivo not in motivos:
                motivos[motivo] = {}
            if tema not in motivos[motivo]:
                motivos[motivo][tema] = []
            motivos[motivo][tema].append(respuesta)

        # Sort motivos and submotivos alphabetically
        motivos = {motivo: dict(sorted(temas.items())) for motivo, temas in sorted(motivos.items())}

        # Check if the dictionary is empty
        if not motivos:
            messagebox.showerror("Error", "El archivo no contiene datos válidos.")
            return None
        
        set_stored_file_path(file_path)  # Guardar la ruta del archivo en el archivo JSON
        return motivos

    except FileNotFoundError:
        messagebox.showerror("Error", "El archivo no se encontró.")
        return None
    except ValueError as e:
        messagebox.showerror("Error", f"Error al leer el archivo: {e}")
        return None
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error inesperado: {e}")
        return None
    
    # Verificar si el archivo existe y se puede leer
    if not archivo or not os.path.exists(archivo):
        messagebox.showerror("Error", "El archivo seleccionado no existe.")
        return None
    if not os.access(archivo, os.R_OK):
        messagebox.showerror("Error", "No se puede leer el archivo seleccionado.")
        return None

    return archivo

def set_stored_file_path(file_path):
    # Store the file path in a JSON file
    with open("reclamos.json", "w") as f:
        f.write(file_path)  # Write the file path to the JSON file
    #print(f"Stored file path: {file_path}")  # Debugging line to check stored path

def get_stored_file_path():
    try:
        if os.path.exists("reclamos.json"):
            with open("reclamos.json", "r") as f:
                return f.read().strip()
        return None
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo leer el archivo de configuración: {e}")
        return None

# Function to handle loading a new Excel file
def cargar_nuevo_archivo():
    global archivo, motivos
    archivo = filedialog.askopenfilename(
        title="Selecciona un archivo .xlsx",
        filetypes=[("Archivos de MS Excel", "*.xlsx")]
    )
    if archivo:
        motivos = fetch_temas(archivo)
        if motivos and isinstance(motivos, dict):
            set_stored_file_path(archivo)
            update_status_bar(archivo)
            combo_motivo["values"] = list(motivos.keys())
            combo_motivo.set("")  # Clear the current selection
            combo_tema.set("")  # Clear the current selection
            respuesta_label.config(text=f"")
            respuesta_text['state'] = 'normal'
            # Limpiar el campo de texto antes de mostrar la nueva respuesta
            respuesta_text.delete(1.0, tk.END)
            respuesta_text['state'] = 'disabled'
            messagebox.showinfo("Éxito", "Archivo cargado correctamente.")
        else:
            messagebox.showerror("Error", "El archivo no tiene un formato válido.")

# funcion para chequear que todas las respuestas contengan placeholders válidos
def check_respuestas_placeholders():
    pattern = r"\{(\w+):([^}:]+)(?::([^}]+))?\}"  # Matches {name:type[:options]}
    valid_field_types = {"text", "date", "month", "option"}
    errores = []
    for motivo, temas in motivos.items():
        for tema, respuestas in temas.items():
            for respuesta in respuestas:
                placeholders = re.findall(r"\{[^}]*\}", respuesta)
                for placeholder in placeholders:
                    match = re.match(pattern, placeholder)
                    if not match:
                        errores.append(
                            f"Motivo: '{motivo}', Tema: '{tema}', Placeholder inválido: {placeholder}"
                        )
                    else:
                        field_type = match.group(1).strip()
                        if field_type not in valid_field_types:
                            errores.append(
                                f"Motivo: '{motivo}', Tema: '{tema}', Placeholder con tipo inválido: {placeholder} (tipo: '{field_type}')"
                            )
    if errores:
        messagebox.showerror("Error", "\n\n".join(errores))
        return False
    return True
# ----------------------------------------------------------------------------
# ----------------------------------------------------------------------------
# Aquí comienza la ejecución del programa
# ----------------------------------------------------------------------------
archivo = get_stored_file_path()

while True:
    if archivo and os.path.exists(archivo) and os.access(archivo, os.R_OK):
        motivos = fetch_temas(archivo)
        if motivos and isinstance(motivos, dict):
            # Successfully parsed the Excel file, update the stored file path
            set_stored_file_path(archivo) 
            break
        else:
            messagebox.showerror("Error", "El archivo no tiene un formato válido.")
    else:
        messagebox.showerror("Error", "El archivo de motivos y respuestas no existe o no se puede abrir. Por favor, seleccione su ubicación.")

    # Open file dialog to select a new Excel file
    archivo = filedialog.askopenfilename(
        title="Selecciona el archivo de motivos y respuestas",
        filetypes=[("Archivos de MS Excel", "*.xlsx")]
    )

    if not archivo:
        # Ask the user if they want to try again
        if not messagebox.askyesno("Error", "No se seleccionó un archivo válido. ¿Quieres intentar de nuevo?"):
            exit()

# ----------------------------------------------------------------------------
# ----------------------------------------------------------------------------
# Generar UI

# Crear ventana
root = tk.Tk()

style = ttk.Style()
style.configure(
    "Custom.TCombobox",  # Style name (must end with .TCombobox)
    fieldbackground="#FFEE99",  # Background color (light yellow)
)
root.title("Seleccionador de Respuestas")

# Create the menu bar
menu_bar = tk.Menu(root)

# Add "Archivo" menu
menu_archivo = tk.Menu(menu_bar, tearoff=0)
menu_archivo.add_command(label="Cargar nuevo archivo .xlsx ...", command=cargar_nuevo_archivo)
menu_archivo.add_separator()
menu_archivo.add_command(label="Verificar formato de respuestas en archivo .xlsx ...", command=check_respuestas_placeholders)
menu_archivo.add_separator()
menu_archivo.add_command(label="Salir", command=root.quit)
menu_bar.add_cascade(label="Archivo", menu=menu_archivo)

# Add "Acerca de ..." menu
menu_bar.add_command(label="Acerca de ...", command=mostrar_acerca_de)

# Configure the menu bar in the root window
root.config(menu=menu_bar)


mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(tk.N, tk.W, tk.E, tk.S))

# Etiqueta y lista desplegable para Motivo
ttk.Label(mainframe, text="Selecciona un Motivo:").grid(column=0, row=0, sticky=tk.W, pady=5)
combo_motivo = ttk.Combobox(mainframe, values=list(motivos.keys()))
combo_motivo.state(["readonly"])
combo_motivo.grid(column=1, row=0, sticky=(tk.W, tk.E))
combo_motivo.bind("<<ComboboxSelected>>", actualizar_temas)

# Etiqueta y lista desplegable para Tema
ttk.Label(mainframe, text="Selecciona un Tema:").grid(column=0, row=1, sticky=tk.W, pady=5)
combo_tema = ttk.Combobox(mainframe)
combo_tema.state(["readonly"])
combo_tema.grid(column=1, row=1, sticky=(tk.W, tk.E))

# Botón para generar respuesta
ttk.Button(mainframe, text="Generar Respuesta", command=generar_respuesta).grid(column=0, row=2, columnspan=2, pady=5)

# Etiqueta para mostrar la respuesta
respuesta_label = ttk.Label(mainframe, text="")
respuesta_label.grid(column=0, row=3, columnspan=2, pady=5)

# campo de texto para mostrar la respuesta
respuesta_text = tk.Text(mainframe, wrap=tk.WORD, height=30, width=150)
# Hacer el campo de texto de solo lectura
# respuesta_text['state'] = 'disabled'
respuesta_text.grid(column=0, row=4, columnspan=2, pady=5, sticky=(tk.N))

# Botón para copiar la respuesta al portapapeles
ttk.Button(mainframe, text="Copiar al portapapeles", command=copiar_respuesta).grid(column=0, row=5, columnspan=2, pady=5)


for child in mainframe.winfo_children(): 
    child.grid_configure(padx=10, pady=10)

# Add a status bar at the bottom of the window
status_bar = ttk.Label(mainframe, text="Archivo no cargado", relief=tk.SUNKEN, anchor=tk.W)
status_bar.grid(column=0, row=6, columnspan=2, sticky=(tk.W, tk.E, tk.S))
update_status_bar(archivo)

root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)
mainframe.columnconfigure(1, weight=1)
mainframe.rowconfigure(4, weight=1)

combo_motivo.focus()
root.bind("<Return>", generar_respuesta)

# Ejecutar la aplicación
try:
    root.mainloop()
except Exception as e:
    messagebox.showerror("Error", f"Ocurrió un error inesperado: {e}")
