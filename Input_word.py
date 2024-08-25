import os
from docxtpl import DocxTemplate
from datetime import datetime
import pandas as pd

# documento word al que se le reemplazan los datos
letter1 = DocxTemplate("plantilla_letteresqueleto.docx")
linkedin = DocxTemplate("plantilla_linkedin.docx")
resum2 = DocxTemplate("plantilla_resume_2_tab.docx")
doc1 = DocxTemplate("plantilla_resume1_esqueleto.docx")



# Fecha Hoy para letter esqueleto
fecha_hoy = datetime.today().strftime("%d %b, %Y")

# Solicita al usuario la ruta y el nombre del archivo de Excel
ruta_excel = input("Introduce la ruta y el nombre del archivo de Excel: ")

# Recorre el excel (el encabezado del excel debe ser igual a lavariable del word)
df = pd.read_excel(ruta_excel, sheet_name='Consolidadov')

for index, fila in df.iterrows():

    context = {

        'nombre_1': fila['nombre_1'],
        'apellido_1': fila['apellido_1'],
        'telefono': fila['telefono'],
        'correo': fila['correo'],
        'linkedin': fila['linkedin'],
        'postula1': fila['postula1'],
        'postula2': fila['postula2'],
        'postula3': fila['postula3'],
        'universidad1': fila['universidad1'],
        'titulo1': fila['Titulo1'],
        'universidad2': fila['universidad2'],
        'Titulo2': fila['Titulo2'],
        'universidad3': fila['universidad3'],
        'Titulo3': fila['Titulo3'],
        'skill_tec1': fila['skill_tec1'],
        'skill_tec2': fila['skill_tec2'],
        'skill_tec3': fila['skill_tec3'],
        'skill_tec4': fila['skill_tec4'],
        'skill_tec5': fila['skill_tec5'],
        'skill_b1': fila['skill_b1'],
        'skill_b2': fila['skill_b2'],
        'skill_b3': fila['skill_b3'],
        'skill_b4': fila['skill_b4'],
        'skill_b5': fila['skill_b5'],
        'idioma1': fila['idioma1'],
        'idioma2': fila['idioma2'],
        'idioma3': fila['idioma3'],
        'certi1': fila['certi1'],
        'institu1': fila['institu1'],
        'año1': fila['año1'],
        'certi2': fila['certi2'],
        'institu2': fila['institu2'],
        'año2': fila['año2'],
        'certi3': fila['certi3'],
        'institu3': fila['institu3'],
        'año3': fila['año3'],
        'certi4': fila['certi4'],
        'institu4': fila['institu4'],
        'año4': fila['año4'],
        'certi5': fila['certi5'],
        'institu5': fila['institu5'],
        'año5': fila['año5'],
        'certi6': fila['certi6'],
        'institu6': fila['institu6'],
        'año6': fila['año6'],
        'certi7': fila['certi7'],
        'institu7': fila['institu7'],
        'año7': fila['año7'],
        'certi8': fila['certi8'],
        'institu8': fila['institu8'],
        'año8': fila['año8'],
        'certi9': fila['certi9'],
        'institu9': fila['institu9'],
        'año9': fila['año9'],
        'certi10': fila['certi10'],
        'institu10': fila['institu10'],
        'año10': fila['año10'],
        'certi11': fila['certi11'],
        'institu11': fila['institu11'],
        'año11': fila['año11'],
        'certi12': fila['certi12'],
        'institu12': fila['institu12'],
        'año12': fila['año12'],
        'volun1': fila['volun1'],
        'volun2': fila['volun2'],
        'volun3': fila['volun3'],
        'volun4': fila['volun4'],
        'volun5': fila['volun5'],
        'volun6': fila['volun6'],
        'cargo1': fila['cargo1'],
        'empresa1': fila['empresa1'],
        'ubicacion1': fila['ubicacion1'],
        'fecha_inicio1': fila['fecha_inicio1'],
        'fecha_fin1': fila['fecha_fin1'],
        'responsabilidades1': fila['responsabilidades1'],
        'logro1': fila['logro1'],
        'cargo2': fila['cargo2'],
        'empresa2': fila['empresa2'],
        'ubicacion2': fila['ubicacion2'],
        'fecha_inicio2': fila['fecha_inicio2'],
        'fecha_fin2': fila['fecha_fin2'],
        'responsabilidades2': fila['responsabilidades2'],
        'logro2': fila['logro2'],
        'cargo3': fila['cargo3'],
        'empresa3': fila['empresa3'],
        'ubicacion3': fila['ubicacion3'],
        'fecha_inicio3': fila['fecha_inicio3'],
        'fecha_fin3': fila['fecha_fin3'],
        'responsabilidades3': fila['responsabilidades3'],
        'logro3': fila['logro3'],
        'cargo4': fila['cargo4'],
        'empresa4': fila['empresa4'],
        'ubicacion4': fila['ubicacion4'],
        'fecha_inicio4': fila['fecha_inicio4'],
        'fecha_fin4': fila['fecha_fin4'],
        'responsabilidades4': fila['responsabilidades4'],
        'logro4': fila['logro4'],
        'cargo5': fila['cargo5'],
        'empresa5': fila['empresa5'],
        'ubicacion5': fila['ubicacion5'],
        'fecha_inicio5': fila['fecha_inicio5'],
        'fecha_fin5': fila['fecha_fin5'],
        'responsabilidades5': fila['responsabilidades5'],
        'logro5': fila['logro5'],
        'cargo6': fila['cargo6'],
        'empresa6': fila['empresa6'],
        'ubicacion6': fila['ubicacion6'],
        'fecha_inicio6': fila['fecha_inicio6'],
        'fecha_fin6': fila['fecha_fin6'],
        'responsabilidades6': fila['responsabilidades6'],
        'logro6': fila['logro6'],
    }

    # Crear carpeta con nombre y apellido
    nombre_apellido = fila['nombre_1'] + "_" + \
        fila['apellido_1'].replace(" ", "_")
    carpeta_path = os.path.join(os.getcwd(), nombre_apellido)
    os.mkdir(carpeta_path)

# _____________________________________________________________Resume2___
    my_context = {}
    for key, value in context.items():
        if not pd.isna(value) and value != '':
            my_context[key] = value
    # saca el nombre y el apellido del archivo
    keys = list(my_context.keys())
    first_key = keys[0]
    second_key = keys[1]
    first_value = my_context[first_key]
    second_value = my_context[second_key]
    the_name = first_value
    the_lastname = second_value

    # Generar archivos dentro de la carpeta
    doc_path = os.path.join(
        carpeta_path, f"{the_name}_{the_lastname}_Resume_2.docx")

    resum2.render(my_context)
    resum2.save(doc_path)
    # _____________________________________________________Resume1________
    my_context2 = {}
    for key, value in context.items():
        if not pd.isna(value) and value != '' and value != 'certificacion':
            my_context2[key] = value

    # saca el nombre y el apellido del archivo
    keys = list(my_context2.keys())
    first_key = keys[0]
    second_key = keys[1]
    first_value = my_context2[first_key]
    second_value = my_context2[second_key]
    the_name2 = first_value
    the_lastname2 = second_value

# Generar archivos dentro de la carpeta
    doc_path = os.path.join(
        carpeta_path, f"{the_name2}_{the_lastname2}_Resume_Master.docx")
    doc1.render(my_context2)
    doc1.save(doc_path)

    # _____________________________________________________Letter_esqueleto1________
    context3 = {
        'nombre_1': fila['nombre_1'],
        'apellido_1': fila['apellido_1'],
        'correo': fila['correo'],
        'linkedin': fila['linkedin'],
        'telefono': fila['telefono'],
        'fecha_hoy': fecha_hoy,
        'postula1': fila['postula1'],
    }
    my_context3 = {}
    for key, value in context3.items():
        if not pd.isna(value) and value != '':
            my_context3[key] = value

    # saca el nombre y el apellido del archivo
    keys = list(my_context3.keys())
    first_key = keys[0]
    second_key = keys[1]
    first_value = my_context3[first_key]
    second_value = my_context3[second_key]
    the_name3 = first_value
    the_lastname3 = second_value

    # Generar archivos dentro de la carpeta
    doc_path = os.path.join(
        carpeta_path, f"{the_name3}_{the_lastname3}_Letter.docx")
    letter1.render(my_context3)
    letter1.save(doc_path)

    # ___________________________________________________________Plantilla_linkedin

    my_context4 = {}
    for key, value in context.items():
        if not pd.isna(value) and value != '':
            my_context4[key] = value

    # saca el nombre y el apellido del archivo
    keys = list(my_context4.keys())
    first_key = keys[0]
    second_key = keys[1]
    first_value = my_context4[first_key]
    second_value = my_context4[second_key]
    the_name4 = first_value
    the_lastname4 = second_value

    # Generar archivos dentro de la carpeta
    doc_path = os.path.join(
        carpeta_path, f"{the_name4}_{the_lastname4}_Linkedin.docx")
    linkedin.render(my_context4)
    linkedin.save(doc_path)
