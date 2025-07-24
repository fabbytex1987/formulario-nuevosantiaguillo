from flask import Flask, render_template, request, redirect, flash
import openpyxl
import os
import re  # Importamos para la validación de la fecha

app = Flask(__name__)
app.secret_key = 'clave_secreta_segura'

archivo = "base_socios.xlsx"

# Crear archivo si no existe
if not os.path.exists(archivo):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Socios"
    ws.append([
        "Nombres y Apellidos",
        "Cédula",
        "Parroquia",
        "Comunidad / Dirección",
        "Fecha de Nacimiento",
        "¿Tiene Discapacidad?",
        "¿Tiene Nichos?",
        "Celular",
        "Correo (Opcional)"
    ])
    wb.save(archivo)

# Verificar si cédula ya fue registrada
def cedula_existente(cedula):
    wb = openpyxl.load_workbook(archivo)
    ws = wb.active
    for fila in ws.iter_rows(min_row=2, values_only=True):
        if str(fila[1]).strip() == cedula.strip():
            return True
    return False

@app.route('/', methods=['GET', 'POST'])
def formulario():
    if request.method == 'POST':
        datos = [
            request.form.get('nombre').strip(),
            request.form.get('cedula').strip(),
            request.form.get('parroquia').strip(),
            request.form.get('direccion').strip(),
            request.form.get('fecha_nacimiento').strip(),
            request.form.get('discapacidad').strip(),
            request.form.get('nichos').strip(),
            request.form.get('celular').strip(),
            request.form.get('correo').strip() if request.form.get('correo') else ''
        ]

        # Validar campos obligatorios (excepto correo)
        if '' in datos[:8]:
            flash("Por favor completa todos los campos obligatorios (excepto correo).", "error")
            return redirect('/')

        # Validar formato de fecha DD/MM/AAAA
        patron_fecha = r"^(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[0-2])/[0-9]{4}$"
        if not re.match(patron_fecha, datos[4]):
            flash("La fecha de nacimiento debe tener formato DD/MM/AAAA.", "error")
            return redirect('/')

        # Verificar cédula duplicada
        if cedula_existente(datos[1]):
            flash(f"La cédula {datos[1]} ya fue registrada.", "error")
            return redirect('/')

        # Guardar datos en Excel
        wb = openpyxl.load_workbook(archivo)
        ws = wb.active
        ws.append(datos)
        wb.save(archivo)

        flash("✅ Datos guardados exitosamente.", "success")
        return redirect('/')

    return render_template('formulario.html')

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))  # Render asigna el puerto dinámicamente
    app.run(host='0.0.0.0', port=port, debug=True)
