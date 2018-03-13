
# 

import xlrd
import json
from collections import OrderedDict

wb = xlrd.open_workbook("FMB_Masterdata_Test.xls")

fmdi_sheet = wb.sheet_by_index(1)
typfall_sheet = wb.sheet_by_index(2)
#icd_array_filename = "json_fmdi_icd.json"
strings_filename = "json_fmdi_trimmed.json"
#typfall_icd_array_filename = "json_typfall_icd.json"
typfall_id_icd_filename = "json_typfall_id_icd.json"
fmdi_id_icd_filename = "json_fmdi_id_icd.json"

typfall_icd_array = []
fmdi_icd_array = []
fmdi_data_list = []
typfall_data_list = []
fmdi_icd_data_list = []


print("Working...", end="")

# Skapa JSON-filer för jämförelse av ICD-koder, för FMDI och Typfall ->

# FMDI

def text_deleter(text_to_delete, text_to_handle, *args, **kwargs):
  text_to_handle = text_to_handle.replace(text_to_delete,"")
  if kwargs is not None:
    for key, value in kwargs.items():
      text_to_handle = text_to_handle.replace(kwargs[key],"")
  return text_to_handle
  #print(text_to_handle)

# TODO för framtiden: Refaktorera denna loop för att hantera alla typer av koder istället för att återanvända koden som i icf_code_handler
for rownum in range(1, fmdi_sheet.nrows):
    print(".", end="")
    row_values = fmdi_sheet.row_values(rownum)
    remove_codesystem = "ICD10:"
    rows_icd_code = row_values[2].rstrip()
    rows_icd_code = text_deleter(remove_codesystem, rows_icd_code)
    #rows_icd_code = rows_icd_code.replace(remove_codesystem,"")
    rows_icd_code = rows_icd_code.split("\n")
    icd_temp_array = []
    final_counter = len(rows_icd_code)
    #print("Denna cell innehåller antal diagnoskoder: " + str(final_counter))
    for index, code in enumerate(rows_icd_code):
        #print("Index number: " + str(index))
        code = code.split(":")
        code.pop(1)
        temp_string = "".join(code)
        no_dots_string = text_deleter('.', temp_string)
        if (index+1) is not final_counter:
            no_dots_string = no_dots_string + ", "
        icd_temp_array1 = []
        icd_temp_array1.append(no_dots_string)
        icd_temp_array = icd_temp_array + icd_temp_array1
    #icd_code_string = '"' + "".join(map(str,icd_temp_array)) + '"'
    icd_code_string = "".join(icd_temp_array)
    #print(icd_code_string)
    fmdi_icd_array.append(icd_code_string)

for rownum in range(1, fmdi_sheet.nrows):
    print(".", end="")
    print(fmdi_icd_array[(rownum-1)])
    fmdi_icd = OrderedDict()
    row_values = fmdi_sheet.row_values(rownum)
    fmdi_icd['ID'] = int(row_values[0])
    fmdi_icd['diagnoskoder'] = fmdi_icd_array[(rownum-1)]
    fmdi_icd_data_list.append(fmdi_icd)
    
print(fmdi_icd_data_list)

j = json.dumps(fmdi_icd_data_list, ensure_ascii=False, sort_keys=True, indent=4)

with open(fmdi_id_icd_filename, 'wb') as f:
    f.write(j.encode())
    #rows_icd_code = rows_icd_code.split(':')
    #fmdi_icd_array.append()
f.close()

# Typfall
    

for rownum in range(1, typfall_sheet.nrows):
    print(".", end="")
    row_values = typfall_sheet.row_values(rownum)
    remove_codesystem = "ICD10:"
    rows_icd_code = row_values[3].rstrip()
    rows_icd_code = text_deleter(remove_codesystem, rows_icd_code)
    #rows_icd_code = rows_icd_code.replace(remove_codesystem,"")
    rows_icd_code = rows_icd_code.split("\n")
    icd_temp_array = []
    final_counter = len(rows_icd_code)
    #print("Denna cell innehåller antal diagnoskoder: " + str(final_counter))
    for index, code in enumerate(rows_icd_code):
        #print("Index number: " + str(index))
        code = code.split(":")
        code.pop(1)
        temp_string = "".join(code)
        no_dots_string = text_deleter('.', temp_string)
        if (index+1) is not final_counter:
            no_dots_string = no_dots_string + ", "
        icd_temp_array1 = []
        icd_temp_array1.append(no_dots_string)
        icd_temp_array = icd_temp_array + icd_temp_array1
    #icd_code_string = '"' + "".join(map(str,icd_temp_array)) + '"'
    icd_code_string = "".join(icd_temp_array)
    #print(icd_code_string)
    typfall_icd_array.append(icd_code_string)

for rownum in range(1, (typfall_sheet.nrows-1)):
    print(".", end="")
    #print(typfall_icd_array[rownum])
    typfall = OrderedDict()
    row_values = typfall_sheet.row_values(rownum)
    typfall['ID'] = int(row_values[2])
    typfall['diagnoskoder'] = typfall_icd_array[rownum]

    typfall_data_list.append(typfall)
    

j = json.dumps(typfall_data_list, ensure_ascii=False, sort_keys=True, indent=4)

with open(typfall_id_icd_filename, 'wb') as f:
    f.write(j.encode())

f.close()

# Fritext FMDI ->

def icf_code_handler(codesystem_prefix, stringified, **kwargs):
  """Handles ICF codes and returns formatted strings in an array"""
  if kwargs is not None:
      temp_holder = ""
      icf_code_array = []
      icf_string_array = []
      for key, value in kwargs.items():
        temp_holder = kwargs[key].replace(codesystem_prefix,"")
      temp_holder_array = temp_holder.split("\n")
      #print(str(temp_holder_array))
      final_counter = len(temp_holder_array)
      # TODO: Refaktorera till egen funktion
      for index, code in enumerate(temp_holder_array):
        code = code.split(":")
        code.pop(1)
        temp_string = "".join(code)
        if (index+1) is not final_counter:
          temp_string = temp_string + ", "
        temp_array = []
        temp_array.append(temp_string)
        icf_code_array = icf_code_array + temp_array
      icf_string = "".join(icf_code_array)
      icf_string_array.append(icf_string)
      #print("Centrala funk: " + str(icf_string_array))
      if stringified is True:
        icf_string_array = "".join(icf_string_array)
      return icf_string_array

for rownum in range(1, fmdi_sheet.nrows):
    print(".", end="")
    fmdi = OrderedDict()
    icf_codes_cent_funk = dict()
    icf_codes_cent_aktiv = dict()
    icf_codes_komp_funk = dict()
    icf_codes_komp_aktiv = dict()
    remove_codesystem = "ICF:"
    row_values = fmdi_sheet.row_values(rownum)
    fmdi['ID'] = int(row_values[0])
    # Ibland följer whitespace med i början eller slutet av strängen:
    rstripped_diagnosrubrik = row_values[1].rstrip()
    lstripped_diagnosrubrik = rstripped_diagnosrubrik.lstrip()
    #
    fmdi['diagnosrubrik'] = lstripped_diagnosrubrik
    fmdi['funktionsnedsattningsbeskrivning'] = row_values[3].rstrip()
    if row_values[4] is not "":
      icf_codes_cent_funk['centrala_funk'] = row_values[4].rstrip()
      #print("Centrala funk: " + str(icf_codes_cent_funk))
      fmdi['centralafunkkoder'] = icf_code_handler(remove_codesystem, True, **icf_codes_cent_funk)
    else:
      fmdi['centralafunkkoder'] = row_values[4].rstrip()
    fmdi['aktivitetsbegransningsbeskrivning'] = row_values[5].rstrip()
    if row_values[6] is not "":
      icf_codes_cent_aktiv['centrala_aktiv'] = row_values[6].rstrip()
      fmdi['centralaaktivitetskoder'] = icf_code_handler(remove_codesystem, True, **icf_codes_cent_aktiv)
    else:
      fmdi['centralaaktivitetskoder'] = row_values[6].rstrip()
    fmdi['rehabiliteringsinfo'] = row_values[7].rstrip()
    fmdi['forsakringsmedicinskinformation'] = row_values[8].rstrip()
    fmdi['symtomprognosbehandling'] = row_values[9].rstrip()
    if row_values[10] is not "":
      icf_codes_komp_funk['kompletterande_funk'] = row_values[10].rstrip()
      fmdi['kompletterandefunkkoder'] = icf_code_handler(remove_codesystem, True, **icf_codes_komp_funk)
    else:
      fmdi['kompletterandefunkkoder'] = row_values[10].rstrip()
    if row_values[11] is not "":
      icf_codes_komp_aktiv['kompletterande_aktiv'] = row_values[11].rstrip()
      fmdi['kompletterandeaktivitetskoder'] = icf_code_handler(remove_codesystem, True, **icf_codes_komp_aktiv)
    else:
      icf_codes_komp_aktiv['kompletterande_aktiv'] = row_values[11].rstrip()
    #print(fmdi['diagnosrubrik']+"#")

    fmdi_data_list.append(fmdi)

k = json.dumps(fmdi_data_list, ensure_ascii=False, sort_keys=True, indent=4)
#print(len(fmdi_data_list))

with open(strings_filename, 'wb') as f:
    f.write(k.encode())
f.close()

print("\n \n \n Done. Saved as {}, {} and {}.".format(fmdi_id_icd_filename, strings_filename, typfall_id_icd_filename))
