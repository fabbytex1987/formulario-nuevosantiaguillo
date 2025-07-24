from flask import Flask, render_template, request, redirect, flash
import openpyxl
import os

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
            request.form.get('nombre'),
            request.form.get('cedula'),
            request.form.get('parroquia'),
            request.form.get('direccion'),
            request.form.get('fecha_nacimiento'),
            request.form.get('discapacidad'),
            request.form.get('nichos'),
            request.form.get('celular'),
            request.form.get('correo')
        ]

        if '' in datos[:8]:  # campos obligatorios
            flash("Por favor completa todos los campos obligatorios.", "error")
            return redirect('/')

        if cedula_existente(datos[1]):
            flash(f"La cédula {datos[1]} ya fue registrada.", "error")
            return redirect('/')

        wb = openpyxl.load_workbook(archivo)
        ws = wb.active
        ws.append(datos)
        wb.save(archivo)

        flash("✅ Datos guardados exitosamente.", "success")
        return redirect('/')

    return render_template('formulario.html')
