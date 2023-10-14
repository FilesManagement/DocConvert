from django.shortcuts import render, redirect
from django.contrib import messages
from docxtpl import DocxTemplate
from .forms import ArchivoForm
import pandas as pd
import os
from django.http import HttpResponse
from django.core.exceptions import ValidationError
from django.contrib.messages import get_messages


df = pd.DataFrame()
docx_template = None

def limpiar_columnas(df):
    df.columns = df.columns.str.replace(' ', '_').str.replace('/', '_').str.replace('(', '').str.replace(':', '').str.replace(')', '').str.strip()

def cargar_excel(excel_file, request=None):
    global df

    try:
        df = pd.read_excel(excel_file)
        limpiar_columnas(df)
        if excel_file:
            pass

    except ValidationError as e:
        if request:
            messages.error(request, f"Error al cargar el archivo Excel: {str(e)}")



def cargar_docx(docx_file, request=None):
    global docx_template

    try:
        docx_template = DocxTemplate(docx_file)
        if docx_file:
            pass
    except Exception as e:
        if request:
            messages.error(request, f"Error al cargar el archivo DOCX: {str(e)}")



def generar_documentos(request):
    global df, docx_template

    if request.method == 'POST':
        form = ArchivoForm(request.POST, request.FILES)

        if form.is_valid():
            excel_file = request.FILES.get('archivo_excel')
            docx_file = request.FILES.get('archivo_docx')

            cargar_excel(excel_file, request)
            cargar_docx(docx_file, request)


            if df.empty or docx_template is None:
                messages.error(request, "Por favor, cargue ambos archivos primero.")
            else:

                start_row = form.cleaned_data['start_row'] - 2
                end_row = form.cleaned_data['end_row'] - 1

                df_subset = df.iloc[start_row:end_row]
                print(df)



                try:
                    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")

                    for index, fila in df_subset.iterrows():
                        context = {column: fila[column] for column in df.columns}
                        Esta_privado = context.get('Se_Encuentra_Privado_de_la_Libertad_si_no', '').lower()

                        if Esta_privado == "si":
                            si = 'X'
                            no = ''
                        else:
                            si = ''
                            no = 'X'

                        context["Esta_privado_si"] = si
                        context["Esta_privado_no"] = no

                        del context['Se_Encuentra_Privado_de_la_Libertad_si_no']


                        doc = docx_template
                        doc.render(context)

                        save_path = os.path.join(downloads_folder, f"{context['Nombre_de_Usuario']}.docx")
                        doc.save(save_path)

                    messages.success(request, "Documentos generados correctamente.")

                except Exception as e:
                    messages.error(request, f"Error al generar documentos: {str(e)}")
        else:
            messages.error(request, "Formulario no válido. Por favor, revise los campos.")



    else:
        form = ArchivoForm()

    return render(request, 'cargar_archivo.html', {'form': form, 'df': df, 'docx_template': docx_template})
