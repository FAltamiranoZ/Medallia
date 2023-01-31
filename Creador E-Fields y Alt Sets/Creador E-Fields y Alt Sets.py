import xlrd
import xlsxwriter as xlsx
import PySimpleGUI as sg
import unidecode

#Código creado por José Francisco Altamirano Zevallos
def generacion(direccionFinal, nombreHojaFinal):
        #Declaración Archivo de lectura
        direccionFinal = direccionFinal
        try:
                workbook = xlrd.open_workbook(direccionFinal)
                try:

                        worksheet = workbook.sheet_by_name(nombreHojaFinal)
                        companyName = str(worksheet.cell(1, 1))[6:-1]
                        programNameLabel = ((unidecode.unidecode(str(worksheet.cell(2, 1))[6:-1])).lower()).replace(" ", "_").replace("'", "").replace("\n", "").replace("\\n", "")
                        print(programNameLabel)

                        #Revisión de si hace falta un archivo de altset
                        num_rows = worksheet.nrows - 1
                        curr_row = 7
                        banderaAltSet = 0
                        while (curr_row < num_rows and banderaAltSet == 0):
                                dataTypeAltSetFlag = ((unidecode.unidecode(str(worksheet.cell(curr_row, 4))[6:-1])).lower()).replace(" ", "_").replace("'", "")
                                lengthAltSetFlag = len(str(worksheet.cell(curr_row, 8))[6:-1].split("\\n"))
                                if(worksheet.cell_type(curr_row, 7) in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK) and dataTypeAltSetFlag == "enumerado" and lengthAltSetFlag > 1):
                                        banderaAltSet = 1
                                curr_row += 1

                        #Creación Archivo de AltSet
                        if(banderaAltSet == 1):
                                aWorkbook = xlsx.Workbook('AltSetBulk' + programNameLabel + '.xlsx')
                                aWorksheet = aWorkbook.add_worksheet()
                                #Creación header AltSet
                                aWorksheet.write('A1', '%%AlternativeSet')
                                row = 1
                                column = 0
                                headers = ["# Key","Name","Company","ContentKind","FormKind","MagicId","StdRange","ForAskNow","Export value is numeric","uuid"]
                                for item in headers:
                                        aWorksheet.write(row, column, item)
                                        column += 1
                                #Creación AltSet
                                row = 2
                                column = 1
                                curr_row = 7
                                contadorPadre = 0
                                nombresAltSet = []
                                while curr_row < num_rows:
                                        dataTypeAltSet = ((unidecode.unidecode(str(worksheet.cell(curr_row, 4))[6:-1])).lower()).replace(" ", "_").replace("'", "")
                                        lengthAltSet = len(str(worksheet.cell(curr_row, 8))[6:-1].split("\\n"))
                                        if(lengthAltSet > 1 and dataTypeAltSet == "enumerado" and worksheet.cell_type(curr_row, 7) in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK)):
                                                altdbArray = str(worksheet.cell(curr_row, 8))[6:-1].split("\\n")
                                                if(len(altdbArray) > 1):
                                                        nombrePadreAltset = str(worksheet.cell(curr_row, 1))[6:-1] + "_" + programNameLabel
                                                        nombresAltSet.append(nombrePadreAltset)
                                                        elements = [nombrePadreAltset,companyName,"ENUMERATION","RADIO_BUTTON","","FALSE","FALSE","FALSE",""]
                                                        for item in elements:
                                                                aWorksheet.write(row, column, item)
                                                                column += 1
                                                        row += 1
                                                        contadorPadre +=1
                                        column = 1
                                        curr_row += 1
                                #Creación header AltDbs
                                row += 1
                                aWorksheet.write(row, 0, "%%AlternativeDb")
                                row += 1
                                column = 0
                                headers = ["# Key","Parent","In survey","In mobile survey","Employee Report","In report","Short form","Description","Visibility","SequenceNumber","NumericValue","Export value","PriorityRaw","RIColumn","RIColSpan","BoxColor","FontColor","Is Other Option","TranslationExplanation","uuid"]
                                for item in headers:
                                        aWorksheet.write(row, column, item)
                                        column += 1
                                #Creación AltDbs
                                row += 1
                                column = 1
                                curr_row = 7
                                while curr_row < num_rows:
                                        dataTypeAltDB = ((unidecode.unidecode(str(worksheet.cell(curr_row, 4))[6:-1])).lower()).replace(" ", "_").replace("'", "")
                                        altdbArray = str(worksheet.cell(curr_row, 8))[6:-1].split("\\n")
                                        if(len(altdbArray) > 1 and dataTypeAltDB == "enumerado" and worksheet.cell_type(curr_row, 7) in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK)):
                                                contadorHijo = 1
                                                nombrePadre = nombresAltSet.pop(0)
                                                banderaPipe = 0
                                                for altdb in altdbArray:
                                                        numericValue = ""
                                                        isOther = ""
                                                        if(worksheet.cell_type(curr_row, 7) not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK) and len(altdbArray) == contadorHijo):
                                                                isOther = "true"
                                                        altdbLabel = altdb.find("|")
                                                        if(altdbLabel != -1):
                                                                altdbPostPipe = str(altdb)[(altdbLabel+1):-1]
                                                                altdbPrePipe = str(altdb)[0:altdbLabel]
                                                                inSurveyAltdb = altdbPostPipe
                                                                inMobileSurveyAltdb = altdbPostPipe
                                                                descriptionAltdb = altdbPostPipe
                                                                inReport = altdbPrePipe
                                                                shortForm = altdbPrePipe
                                                                if(altdbPrePipe.isnumeric()):
                                                                        numericValue = altdbPrePipe
                                                                banderaPipe = 1
                                                        elif(banderaPipe == 1):
                                                                inSurveyAltdb = "[blank]"
                                                                inMobileSurveyAltdb = "[blank]"
                                                                descriptionAltdb = altdb
                                                                inReport = altdb
                                                                shortForm = altdb
                                                                if(altdb.isnumeric()):
                                                                        numericValue = altdb
                                                        else:
                                                                inSurveyAltdb = altdb
                                                                inMobileSurveyAltdb = altdb
                                                                descriptionAltdb = altdb
                                                                inReport = altdb
                                                                shortForm = altdb
                                                                if(altdb.isnumeric()):
                                                                        numericValue = altdb
                                                        elements = [nombrePadre,inSurveyAltdb,inMobileSurveyAltdb,inReport,inReport,shortForm,descriptionAltdb,"SURVEY_AND_REPORTING_REQUIRED",contadorHijo,numericValue,"","","","","","",isOther,"",""]
                                                        for item in elements:
                                                                aWorksheet.write(row, column, item)
                                                                column += 1
                                                        column = 1
                                                        contadorHijo += 1
                                                        row += 1
                                        column = 1
                                        curr_row += 1
                                aWorkbook.close()
                        #Creación Archivo de EField
                        eWorkbook = xlsx.Workbook('EFieldBulk' + programNameLabel + '.xlsx')
                        eWorksheet = eWorkbook.add_worksheet()
                        #Creación header EFields
                        eWorksheet.write('A1', '%%Efield')
                        row = 1
                        column = 0
                        headers = ["# Key","Name","Short","Keyname","Priority","Description","Required","Used for Duplicate Checking, Cohort Tracking, Sampling Priority, Episode Conditions, or Quarantine Rules","Used for ACE","Sticky","AlternativeSet","Company","Client identifier","Export label","Personally Identifying Data","Encrypted","TranslationExplanation","uuid"]
                        for item in headers:
                                eWorksheet.write(row, column, item)
                                column += 1
                        #Creación EFields
                        row = 2
                        curr_row = 7
                        priority = 10
                        while curr_row < (num_rows + 1):
                                if(worksheet.cell_type(curr_row, 7) in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK)):
                                        column = 0
                                        name = str(worksheet.cell(curr_row, 2))[6:-1]
                                        nameLabel = ((unidecode.unidecode(name)).lower()).replace(" ", "_").replace("'", "")
                                        dataType = ((unidecode.unidecode(str(worksheet.cell(curr_row, 4))[6:-1])).lower()).replace(" ", "_").replace("'", "")
                                        match dataType:
                                                case "autoindexado":
                                                        termination = "ai"
                                                        altSet = 58
                                                case "enumerado":
                                                        termination = "enum"
                                                case "texto":
                                                        termination = "txt"
                                                        altSet = 42
                                                case "fecha":
                                                        termination = "date"
                                                        altSet = 46
                                                case "fecha_y_hora":
                                                        termination = "datetime"
                                                        altSet = 47
                                                case "si/no":
                                                        termination = "yn"
                                                        altSet = 13
                                                case "entero":
                                                        termination = "int"
                                                        altSet = 44
                                                case "fraccional":
                                                        termination = "real"
                                                        altSet = 51
                                                case "unidad":
                                                        termination = "unit"
                                                        altSet = 3
                                                case "email":
                                                        termination = "email"
                                                        altSet = 7
                                                case _:
                                                        termination = ""
                                                        altSet = ""
                                        if (worksheet.cell_type(curr_row, 3) not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK) and str(worksheet.cell(curr_row, 3))[6:-1] == "S"):
                                                required = "true"
                                        else:
                                                required = "FALSE"
                                        if (worksheet.cell_type(curr_row, 10) not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK) and str(worksheet.cell(curr_row, 10))[6:-1] == "S"):
                                                duplicateChecking = "true"
                                        else:
                                                duplicateChecking = "FALSE"
                                        if (worksheet.cell_type(curr_row, 9) not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK) and str(worksheet.cell(curr_row, 9))[6:-1] == "S"):
                                                ACE = "true"
                                        else:
                                                ACE = "FALSE"
                                        if(worksheet.cell_type(curr_row, 8) not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK) and termination == "enum"):
                                                altdbArray = str(worksheet.cell(curr_row, 8))[6:-1].split("\\n")
                                                if(len(altdbArray) > 1):
                                                        altSet="Poner Aquí el altset generado al procesar el archivo de altset spec"
                                                else:
                                                        altSet=str(worksheet.cell(curr_row, 8))[7:-1].replace(".", "")
                                        efield = ["e_" + companyName + "_" + programNameLabel + "_" + nameLabel + "_" + termination,name,name,companyName + "_" + programNameLabel + "_" + nameLabel + "_" + termination,priority,"",required,duplicateChecking,ACE,"",altSet,companyName,"",name,"","","e_" + companyName + "_" + programNameLabel + "_" + nameLabel + "_" + termination,""]
                                        for efieldComponent in efield:
                                                eWorksheet.write(row, column, efieldComponent)
                                                column += 1
                                        row += 1
                                        priority += 10
                                curr_row += 1
                        eWorkbook.close()
                        sg.popup('Resultados: ','\nSe han generado tus archivos en la misma carpeta en la que se encuentra este programa, si no los ves, refresca la página o revisa que el spec esté llenado correctamente. \n\nEn caso de que hayas creado altsets, no olvides rellenar el espacio del altset en el excel de los E-Fields con los valores de los recien creados. \n')
                except:
                        sg.popup('Error: ','No se encuentra la hoja especificada')
        except:
                sg.popup('Error: ','No se encuentra el archivo en la dirección especificada')


        

#Interfaz Gráfica

sg.theme('Reddit') 
 
layout = [  [sg.Text('Porfavor introduce la dirección del archivo y el nombre de la hoja en los recuadros inferior y posteriormente haz clic en el botón de "Generar"')], 
            [sg.Text('Dirección del archivo con formato "C:\CarpetaProyecto\SpecMuestraAutoImporter.xls": '), sg.InputText(key='direccion')], 
            [sg.Text('Nombre de la hoja de excel: '), sg.InputText(key='nombreHoja')],
            [sg.Button('Generar')],
            [sg.Button('Salir')],
            [sg.Text(' ')],
            [sg.Text('Creado por: JFAZO')],
            [sg.Text('Versión 1.0')]
        ] 

window = sg.Window('Generador E-Fields y Alt Sets', layout) 

while True: 
    event, values = window.read() 
    if event == sg.WIN_CLOSED or event == 'Salir': 
        break 
    elif event == 'Generar':
        ValorDireccion = values['direccion']
        ValorNombreHoja = values['nombreHoja']
        generacion(ValorDireccion, ValorNombreHoja)

window.close()
