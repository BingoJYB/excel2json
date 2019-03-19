import xlrd
import uuid
import json


def readExcel(path):
    workBook = xlrd.open_workbook(path)
    medication_sheet = workBook.sheet_by_name('Medication')
    medication_form_sheet = workBook.sheet_by_name('Medication Form PICASO')

    medication_keys = [medication_sheet.cell(0, col_index).value for col_index in range(medication_sheet.ncols)]
    medication_form_keys = [medication_form_sheet.cell(0, col_index).value for col_index in
                            range(medication_form_sheet.ncols)]

    medications = []
    for row_index in range(1, medication_sheet.nrows):
        medication = {medication_keys[col_index]: medication_sheet.cell(row_index, col_index).value
                      for col_index in range(medication_sheet.ncols)}
        medications.append(medication)

    medication_forms = []
    for row_index in range(1, medication_form_sheet.nrows):
        medication_form = {medication_form_keys[col_index]: medication_form_sheet.cell(row_index, col_index).value
                           for col_index in range(medication_form_sheet.ncols)}
        medication_forms.append(medication_form)

    return medications, medication_forms


def setFormJson(system='', code='', form=''):
    return {
        "system": system,
        "code": code,
        "display": form
    }


def setIngredientJson(medication, quantity, unit):
    return {
        "itemReference":
            {
                "display": medication
            },
        "isActive": "true",
        "amount":
            {
                "numerator":
                    {
                        "value": quantity,
                        "unit": unit
                    },
                "denominator":
                    {
                        "value": "1",
                        "unit": ""
                    }
            }
    }


def setMedicationJson(manufacturer):

    return {
        "id": '',
        "resourceType": "Medication",
        "manufacturer":
            {
                "display": manufacturer
            },
        "form":
            {
                "coding":
                    [
                    ]
            },
        "ingredient":
            [
            ]
    }


def setBundleJson(medications, medication_forms):
    index = 1
    bundle = {
        "id": str(uuid.uuid4()),
        "resourceType": "Bundle",
        "type": "collection",
        "entry":
            [
            ]
    }

    for medication in medications:
        hasDosage = False
        medicationJson = setMedicationJson(medication['Trade name'])

        for id, med in enumerate(medication['Active substance(s)'].split('/')):
            if medication['Dosage(s) of substance(s)'] != "":
                hasDosage = True
                quantity, unit = medication['Dosage(s) of substance(s)'].split(' ', 1)
                quantity_list = quantity.split('/')
                medicationJson['ingredient'].append(setIngredientJson(med, quantity_list[id], unit))

        if hasDosage:
            form = medication['Form']

            if form != '':
                try:
                    system = [f['System'] for f in medication_forms if form == f['Display']][0]
                    code = int([f['Code'] for f in medication_forms if form == f['Display']][0])
                except Exception:
                    code = [f['Code'] for f in medication_forms if form == f['Display']][0]

            medicationJson['id'] = 'it-' + stringUtils(index)
            medicationJson['form']['coding'].append(setFormJson(system, code, form))
            bundle['entry'].append(medicationJson)
            index = index + 1

    return bundle


def stringUtils(id):
    return "00" + str(id) if len(str(id)) == 1 else "0" + str(id) if len(str(id)) == 2 else str(id)


def writeJson(bundle, path):
    with open(path, 'w') as outfile:
        json.dump(bundle, outfile)


if __name__ == "__main__":
    medications, medication_forms = readExcel("Medication-List-IT.xlsx")
    writeJson(setBundleJson(medications, medication_forms), "Medication Object.json")
