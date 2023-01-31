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
                        nombreCompañia = str(worksheet.cell(1, 1))[6:-1]
                        nombreEncuesta = str(worksheet.cell(2, 1))[6:-1]
                        nombreEncuestaLabel = ((unidecode.unidecode(nombreEncuesta)).lower()).replace(" ", "_").replace("'", "")

                        #Revisión de si hace falta un archivo de altset
                        num_rows = worksheet.nrows - 1
                        curr_row = 6
                        banderaAltSet = 0
                        while (curr_row < num_rows and banderaAltSet == 0):
                                if(worksheet.cell_type(curr_row, 8) not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK)):
                                        banderaAltSet = 1
                                curr_row += 1

                        #Creación Archivo de AltSet
                        if(banderaAltSet == 1):
                                aWorkbook = xlsx.Workbook('AltSetBulk' + nombreEncuesta + '.xlsx')
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
                                curr_row = 6
                                contadorPadre = 0
                                nombresAltSet = []
                                while curr_row < num_rows:
                                        if(worksheet.cell_type(curr_row, 8) in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK) and worksheet.cell_type(curr_row, 1) not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK)):
                                                nombrePadreAltset = str(worksheet.cell(curr_row, 1))[6:-1] + " " + nombreEncuesta
                                                nombresAltSet.append(nombrePadreAltset)
                                                elements = [nombrePadreAltset,nombreCompañia,"ENUMERATION","RADIO_BUTTON","","FALSE","FALSE","FALSE",""]
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
                                headers = ["# Key","Parent","In survey","In mobile survey","In report","Employee Report","Short form","Description","Visibility","SequenceNumber","NumericValue","Export value","PriorityRaw","RIColumn","RIColSpan","BoxColor","FontColor","Is Other Option","TranslationExplanation","uuid"]
                                for item in headers:
                                        aWorksheet.write(row, column, item)
                                        column += 1
                                #Creación AltDbs
                                row += 1
                                column = 1
                                curr_row = 6
                                while curr_row < num_rows:
                                        if(worksheet.cell_type(curr_row, 8) in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK) and worksheet.cell_type(curr_row, 1) not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK)):
                                                altdbArray = str(worksheet.cell(curr_row, 3))[6:-1].split("\\n")
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

                        #Creación Archivo de QField
                        qWorkbook = xlsx.Workbook('QFieldBulk' + nombreEncuesta + '.xlsx')
                        qWorksheet = qWorkbook.add_worksheet()
                        #Creación header QFields
                        qWorksheet.write('A1', '%%Question')
                        row = 1
                        column = 0
                        headers = ["# Key","Parent","Category","Keyname","Name","ShortName","Abbreviation","In survey","In mobile survey","With user text","Priority","Description","Used for Duplicate Checking, Cohort Tracking, Sampling Priority, Episode Conditions, or Quarantine Rules","Used for ACE","OtherQuestion","AlternativeSet","Kind","Formatting","MagicFlag","Client identifier","Export label","Personally Identifying Data","TranslationExplanation"]
                        for item in headers:
                                qWorksheet.write(row, column, item)
                                column += 1
                        #Creación parent QFields
                        row = 2
                        column = 0
                        headers = ["q_" + nombreCompañia + "_" + nombreEncuestaLabel,"",nombreEncuesta,nombreCompañia + "_" + nombreEncuestaLabel,nombreEncuesta,nombreEncuesta,nombreEncuesta,nombreEncuesta,nombreEncuesta,nombreEncuesta,10000,"","FALSE","FALSE","","","HEADING","REGULAR","","","q_" + nombreCompañia + "_" + nombreEncuestaLabel,"FALSE","q_" + nombreCompañia + "_" + nombreEncuestaLabel]
                        for item in headers:
                                qWorksheet.write(row, column, item)
                                column += 1
                        #Creación QFields
                        row = 3
                        column = 0
                        curr_row = 6
                        priority = 10
                        while curr_row < num_rows:
                                if(worksheet.cell_type(curr_row, 0) not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK) and worksheet.cell_type(curr_row, 1) not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK)):
                                        name = str(worksheet.cell(curr_row, 1))[6:-1]
                                        inSurvey = str(worksheet.cell(curr_row, 2))[6:-1]
                                        if worksheet.cell_type(curr_row, 6) not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                                                exportLabel = str(worksheet.cell(curr_row, 6))[6:-1]
                                        else:
                                                exportLabel = ""
                                        if worksheet.cell_type(curr_row, 8) not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                                                altSetQuestion = str(worksheet.cell(curr_row, 8))[7:-2]
                                        else:
                                                altSetQuestion = "Poner Aquí el altset generado al procesar el archivo de altset spec"
                                        termination = str(worksheet.cell(curr_row, 9))[6:-1]
                                        if worksheet.cell_type(curr_row, 10) not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                                                ACE = str(worksheet.cell(curr_row, 10))[6:-1]
                                        else:
                                                ACE = "FALSE"
                                        if worksheet.cell_type(curr_row, 11) not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                                                duplicateChecking = str(worksheet.cell(curr_row, 11))[6:-1]
                                        else:
                                                duplicateChecking = "FALSE"
                                        nameLabel = ((unidecode.unidecode(name)).lower()).replace(" ", "_").replace("'", "")
                                        if worksheet.cell_type(curr_row, 7) not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                                                isOtherQuestion = str(worksheet.cell(curr_row, 7))[6:-1]
                                                if(isOtherQuestion == "true"):
                                                        qfieldsOther = ["q_" + nombreCompañia + "_" + nombreEncuestaLabel + "_" + nameLabel + "_" + "otro_txt","q_" + nombreCompañia + "_" + nombreEncuestaLabel,nombreEncuesta,nombreCompañia + "_" + nombreEncuestaLabel + "_" + nameLabel + "_" + "otro_txt",name + " Otro",name + " Otro",name + " Otro",name + " Otro",name + " Otro",name + " Otro",priority,"","FALSE","FALSE","",42,"QUESTION","REGULAR","","",exportLabel+ " Otro","FALSE","q_" + nombreCompañia + "_" + nombreEncuestaLabel + "_" + nameLabel + "_" + "otro_txt"]
                                                        for qfieldOther in qfieldsOther:
                                                                qWorksheet.write(row, column, qfieldOther)
                                                                column += 1
                                                        column = 0
                                                        row += 1
                                                        priority += 10
                                                        isOtherQuestion = "q_" + nombreCompañia + "_" + nombreEncuestaLabel + "_" + nameLabel + "_" + "otro_txt"
                                                else:
                                                        isOtherQuestion = ""
                                        else:
                                                isOtherQuestion = ""
                                        qfields = ["q_" + nombreCompañia + "_" + nombreEncuestaLabel + "_" + nameLabel + "_" + termination,"q_" + nombreCompañia + "_" + nombreEncuestaLabel,nombreEncuesta,nombreCompañia + "_" + nombreEncuestaLabel + "_" + nameLabel + "_" + termination,name,name,name,inSurvey,inSurvey,name,priority,"",duplicateChecking,ACE,isOtherQuestion,altSetQuestion,"QUESTION","REGULAR","","",exportLabel,"FALSE","q_" + nombreCompañia + "_" + nombreEncuestaLabel + "_" + nameLabel + "_" + termination]
                                        for qfield in qfields:
                                                qWorksheet.write(row, column, qfield)
                                                column += 1
                                        column = 0
                                        row += 1
                                        priority += 10
                                curr_row += 1
                        qWorkbook.close()
                        sg.popup('Resultados: ','\nSe han generado tus archivos en la misma carpeta en la que se encuentra este programa, si no los ves, refresca la página o revisa que el spec esté llenado correctamente. \n\nEn caso de que hayas creado altsets, no olvides rellenar el espacio del altset en el excel de los Q-Fields con los valores de los recien creados. \n\nEn caso de que hayas creado una pregunta de Other, recuerda procesar el bulk sin el campo que tiene el other question y una vez procesado, procesar unicamente el de other question\n')
                except:
                        sg.popup('Error: ','No se encuentra la hoja especificada')
        except:
                sg.popup('Error: ','No se encuentra el archivo en la dirección especificada')


        

#Interfaz Gráfica

sg.theme('Reddit') 
 
layout = [  [sg.Text('Porfavor introduce la dirección del archivo y el nombre de la hoja en los recuadros inferior y posteriormente haz clic en el botón de "Generar"')], 
            [sg.Text('Dirección del archivo con formato "C:\CarpetaProyecto\SpecMuestraEncuesta.xls": '), sg.InputText(key='direccion')], 
            [sg.Text('Nombre de la hoja de excel: '), sg.InputText(key='nombreHoja')],
            [sg.Button('Generar')],
            [sg.Button('Salir')],
            [sg.Text(' ')],
            [sg.Text('Creado por: JFAZO')],
            [sg.Text('Versión 2.0')]
        ] 

window = sg.Window('Generador Q-Fields y Alt Sets', layout) 

while True: 
    event, values = window.read() 
    if event == sg.WIN_CLOSED or event == 'Salir': 
        break 
    elif event == 'Generar':
        ValorDireccion = values['direccion']
        ValorNombreHoja = values['nombreHoja']
        generacion(ValorDireccion, ValorNombreHoja)

window.close()
