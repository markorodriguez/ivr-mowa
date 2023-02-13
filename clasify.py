import pandas as pd
import xlsxwriter

high = ['si',  'hola si', 'si hola', 'ok si', 'si ok', 'yo si', 'claro si', 'si claro',
        'si vale', 'vale si', 'alo si', 'si alo']  # x3
low = ['buenas', 'buenos', 'dias', 'buenas', 'tardes', 'alo', 'ok']  # x2
to_confirm = ['hola soy']  # x1
wrong_person = ['no', 'equivocado', 'yo no']  # -1
machine = ['tono', 'buzon', 'mensaje', 'voz']  # +10


def normalize(s):
    replacements = (
        ("á", "a"),
        ("é", "e"),
        ("í", "i"),
        ("ó", "o"),
        ("ú", "u"),
    )
    for a, b in replacements:
        s = s.replace(a, b).replace(a.upper(), b.upper())
    return s


rows = []


def main():
    excel = pd.ExcelFile('./RPTS/RPTA_SULLANA.xlsx').parse('SULLANA')

    for i in range(len(excel)):
        json = {}
        json['audio'] = excel['Audio'][i]
        json['texto'] = excel['Texto'][i]
        txt = (normalize(excel['Texto'][i].lower()))
        json['dni'] = excel['DNI'][i]
        json['teléfono'] = excel['Teléfono'][i]
        json['campaña'] = excel['Campaña'][i]
        if (excel['Texto'][i]) == '-x-' or (excel['Texto'][i]) == '-':
            json['clasificacion'] = 'Sin respuesta'
            json['max_value'] = 0
        else:
            # high opportunity
            counter_high = 0
            for high_opportunity in high:
                for txt_word in txt.split(" "):
                    if high_opportunity.lower() == txt_word.lower():
                        counter_high += 4
                # if high_opportunity.lower() in normalize(excel['Texto'][i].lower()):
                # if normalize(excel['Texto'][i]).__contains__(high_opportunity):
                    # counter_high += 3
            counter_low = 0

            for low_opportunity in low:
                for txt_word in txt.split():
                    if low_opportunity.lower() == txt_word.lower():
                        counter_low += 2 
                # if low_opportunity.lower() in normalize(excel['Texto'][i].lower()):
                # if normalize(excel['Texto'][i]).__contains__(low_opportunity):
                    # counter_low += 2
            counter_to_confirm = 0
            for to_confirm_op in to_confirm:
                for txt_word in txt.split():
                    if to_confirm_op.lower() == txt_word.lower():
                        counter_to_confirm += 1
                # if to_confirm_op.lower() in normalize(excel['Texto'][i].lower()):
                # if normalize(excel['Texto'][i]).__contains__(to_confirm_op):
                    # counter_to_confirm += 1
            counter_wrong_person = 0
            for wrong_person_op in wrong_person:
                for txt_word in txt.split():
                    if wrong_person_op.lower() == txt_word.lower():
                        counter_wrong_person += 6
                # if wrong_person_op.lower() in normalize(excel['Texto'][i].lower()):
                # if normalize(excel['Texto'][i]).__contains__(wrong_person_op):
                    # counter_wrong_person += 6

            counter_machine = 0
            for machine_op in machine:
                if machine_op.lower() in normalize(excel['Texto'][i].lower()):
                    # if normalize(excel['Texto'][i]).__contains__(machine_op):
                    counter_machine += 20

            counter_some = 0
            for txt_word in txt.split():
                    if txt_word.lower() not in high and txt_word.lower() not in low and txt_word.lower() not in to_confirm and txt_word.lower() not in wrong_person and txt_word.lower() not in machine:
                        counter_some += 1

            values = [counter_high, counter_low, counter_to_confirm,
                      counter_wrong_person, counter_machine, counter_some]
            max_value = max(values)

            if max_value == 0:
                json['clasificacion'] = 'OTROS'
            elif max_value == counter_machine:
                json['clasificacion'] = 'Máquina'
            elif max_value == counter_high:
                json['clasificacion'] = 'Nivel alto de confirmación'
            elif max_value == counter_low:
                json['clasificacion'] = 'Nivel bajo de confirmación'
            elif max_value == counter_to_confirm:
                json['clasificacion'] = 'Por confirmar'
            elif max_value == counter_wrong_person:
                json['clasificacion'] = 'Persona equivocada'
            elif max_value == counter_some:
                json['clasificacion'] = 'OTROS'
            json['max_value'] = max_value
        rows.append(json)

    workbook = xlsxwriter.Workbook('./RPTS/RPTA_SULLANA_DONE.xlsx')
    worksheet = workbook.add_worksheet()

    worksheet.write(0, 0, 'Audio')
    worksheet.write(0, 1, 'Texto')
    worksheet.write(0, 2, 'DNI')
    worksheet.write(0, 3, 'Teléfono')
    worksheet.write(0, 4, 'Campaña')
    worksheet.write(0, 5, 'Clasificación')
    worksheet.write(0, 6, 'Max value')

    for i in range(len(rows)):
        worksheet.write(i+1, 0, rows[i]['audio'])
        worksheet.write(i+1, 1, rows[i]['texto'])
        worksheet.write(i+1, 2, rows[i]['dni'])
        worksheet.write(i+1, 3, rows[i]['teléfono'])
        worksheet.write(i+1, 4, rows[i]['campaña'])
        worksheet.write(i+1, 5, rows[i]['clasificacion'])
        worksheet.write(i+1, 6, rows[i]['max_value'])

    workbook.close()


main()
