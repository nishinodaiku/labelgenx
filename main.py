# Author: Marcos Suarez
# Date 18/11/2024
# This script resolves the functionality lost for the Archives personnel during the migration process to Windows 10
# as an old .doc file with a embedded script no longer works as expected. I write this to speed up the processes of
# the actual project and to bring a new, maintainable and efficient tool to the hands of the team that uses this.

import os
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.styles import Alignment
from datetime import datetime

def pages():
    # this function is in charge of getting the number of pages the user wants to print
    while True:
        pgs = input('Ingrese cantidad de planchas a generar: ')
        try:
            pgs = int(pgs)
        except:
            print('Entrada no valida. Intente nuevamente.\n')
            continue
        if pgs <= 0:
            print('No hay planchas que imprimir. Saliendo del programa....')
            quit()
        else:
            print(f'\nSe generaran {pgs*65} etiquetas del tipo "LEG" y {pgs*65} del tipo "PAG".')
            print(f'Seran un total de {pgs*65*2} etiquetas que utilizara {pgs*2} planchas.')
            ur = user_response(input('Desea continuar? (Y/n): '))
            if ur == 0:
                break
            elif ur < 0:
                continue
    return pgs

def label():
    # here the function asks the user to anwser with the initial label value
    while True:
        lbl = input('\nIngrese el numero inicial de la serie de etiquetas: ')
        try:
            lbl = int(lbl)
        except:
            print('Entrada no valida. Intente nuevamente.')
            continue
        if lbl < 0:
            print('Entrada no valida. Intente nuevamente.')
            continue
        else:
            print(f'\nEl numero de inicial sera {lbl}')
            ur = user_response(input('Desea continuar? (Y/n): '))
            if ur == 0:
                break
            elif ur < 0:
                continue
    return lbl

def user_response(yesno):
    # small function to get confirmation from user to go on
    yesno = yesno.upper()
    if yesno == '':
        return 0
    elif yesno in ['Y', 'YES']:
        return 0
    elif yesno in ['N', 'NO']:
        return -1
    else:
        print('Entrada no valida. Intente nuevamente.\n')
        return -2

def load_bp():
    # we try to load the file used as blueprint
    while True:
        try:
            wb = load_workbook('blueprint.xlsx')
            break
        except FileNotFoundError:
            print(f'El archivo "blueprint.xlsx" no se ha podido cargar.\nPor favor asegurese de que se encuentre en la carpeta donde esta el script.')
            quit()
        except Exception as e:
            print(f'Ocurrio un error. Codigo: {e}')
            quit()
    return wb

def save_path():
    user_profile = os.environ['USERPROFILE']
    current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    save_path = os.path.join(user_profile,'Desktop',f'etiquetas_{current_time}.xlsx')
    return save_path
    
def generator():
    # we start to gather all data to be able to create the file as is expected
    wb = load_bp()
    blueprint = wb["base"]
    master_sheet = wb.copy_worksheet(blueprint)
    master_sheet.title = "etiquetas"

    total_pages = pages()
    starting_label = label()
    # all this configuration tries to imitate the old script to ensure compatibility
    # of the label sheets used during printing
    columns = ['A', 'C', 'E', 'G', 'I']
    rows_per_page = 25
    font1 = Font(name='Code39HalfInch', size=20)
    font2 = Font(name='Arial', size=12)

    current_row = 1
    current_label = starting_label
    # after a few tries i realized i have no other option but to specify aligment
    for page in range(total_pages):
        for row_offset in range(0, rows_per_page, 2):
            for col in columns:
                cell1 = f'{col}{current_row}'
                cell2 = f'{col}{current_row + 1}'
                master_sheet[cell1] = f'*{current_label}LEG*'
                master_sheet[cell1].font = font1
                master_sheet[cell1].alignment = Alignment(horizontal='center',vertical='bottom')
                master_sheet[cell2] = f'{current_label}LEG'
                master_sheet[cell2].font = font2
                master_sheet[cell2].alignment = Alignment(horizontal='center',vertical='top')
                current_label += 1
            current_row += 2

    current_label = starting_label

    for page in range(total_pages):
        for row_offset in range(0, rows_per_page, 2):
            for col in columns:
                cell1 = f'{col}{current_row}'
                cell2 = f'{col}{current_row + 1}'
                master_sheet[cell1] = f'*{current_label}PAG*'
                master_sheet[cell1].font = font1
                master_sheet[cell1].alignment = Alignment(horizontal='center',vertical='bottom')
                master_sheet[cell2] = f'{current_label}PAG'
                master_sheet[cell2].font = font2
                master_sheet[cell2].alignment = Alignment(horizontal='center',vertical='top')
                current_label += 1
            current_row += 2

    # making sure all rows are the same height
    first_row_height = master_sheet.row_dimensions[1].height
    for row in range(1, current_row + 1):
        master_sheet.row_dimensions[row].height = first_row_height

    # removing the sheet used as blueprint from the new file before saving
    wb.remove(blueprint)

    print(f'\nIntentando generar archivo...')    
    try:
        wb.save(save_path())
    except Exception as e:
        print('Ocurrio un error. Codigo {e}')
        quit()

    print('\nArchivo generado exitosamente. Saliendo del programa...')

def welcome():
    # welcome message
    print('\nLabel Generator V1.0.0')
    print('Diseñado y creado por Marcos Suarez  para personal del área de Archivos.')
    print('Contacto: corruptedkeep@proton.me LinkedIn: https://linkedin.com/in/marcossuarezit')
    print('Repositorio oficial en https://github.com/nishinodaiku/labelgenx')
    print('\nEste programa está licenciado bajo la Licencia MIT.')
    print('Puedes usar, modificar y distribuir este software bajo los términos de esta licencia.\n')

# all that work for this
welcome()
generator()