import os
import psse3503
import psspy
import excelpy
from tkinter import filedialog , Tk
root = Tk() #CODIGO NECESARIO PARA OCULTAR VENTANA TK
root.withdraw() #CODIGO NECESARIO PARA OCULTAR VENTANA TK
userDir = os.path.expanduser("~")


#INDICO CUAL ES EL CASO DE ESTUDIO
file = filedialog.askopenfilename(title="SELECCIONA EL CASO DE ESTUDIO: ", filetypes = (("CASE PSSE","*.sav"),("All files","*.*")),\
     initialdir='%s/Documents/PTI/' % userDir) #OBTENGO LA RUTA DEL DIRECTORIO DEL CASO PSSE
nameCasePsse = os.path.basename(file) #RECORTO EL DIRECTORIO DEL ARCHIVO PARA OBTENER SOLO EL NOMBRE
psspy.psseinit(10000) #INICIO EL PSSE
psspy.case(file) 

#CREO LAS VARIABLES CON EL CONTENIDO DEL PSSE, PARA LUEGO PEGARLAS EN EL EXCEL

#VARIABLES DE BARRAS
ierr, busesNames = psspy.abuschar(-1, string="NAME")     #IMPORTA EL NOMBRE DE LA BARRA, YA QUE ANTERIORMENTE ESTAMOS PONIENDO EL NUMERO DE LA BARRA EN EL NOMBRE
busesNames = list(map(lambda x: x.strip(), busesNames[0]))
busesWithNames = False
for i in busesNames:
    if i:
        busesWithNames = True
ierr, busesVoltageBase = psspy.abusreal(-1, string="BASE")
busesVoltageBaseRounded = list(map(lambda x: round(x,1), busesVoltageBase[0]))
ierr, busesVoltage = psspy.abusreal(-1, string="KV")
busesVoltageRounded = list(map(lambda x: round(x,2), busesVoltage[0]))
ierr, busesVoltagePU = psspy.abusreal(-1, string="PU")
busesVoltagePURounded = list(map(lambda x: round(x,2), busesVoltagePU[0]))
ierr, busesNumbers = psspy.abusint(-1, string="NUMBER")  #LO SOLICITO PARA RELACIONAR EL NUMERO DE BARRA CON SU TENSION
ierr, busesTypes = psspy.abusint(-1, string="TYPE")
busesVoltageDiff = list(map(lambda x: x*100-100 ,busesVoltagePU[0]))
busesVoltageDiffRounded = list(map(lambda x: round(x,1), busesVoltageDiff))
busesVoltageDiffMax =list(map(lambda x: 8 if x<=33 else 7 if x<=66 else 5 if x<=220 else 3, busesVoltageBase[0]))
dicBusesRelation = dict(zip(busesNumbers[0], busesVoltageBase[0]))
#AVERIGUO CUAL ES LA BARRA SLACK Y GUARDO EL NUMERO DE BARRA EN slackBus
for i in range(len(busesTypes[0])):
    if busesTypes[0][i] == 3:
        slackBus = busesNumbers[0][i]


#VARIABLES DE LINEAS (ESTA CREADO DESPUES DE BARRAS, PORQUE NECESITO LA TENSION DE CADA BARRA PARA HALLAR LA TENSION BASE DE LAS LINEAS)
ierr, branchesNames = psspy.abrnchar(-1, string="BRANCHNAME")
branchesNames = list(map(lambda x: x.strip(), branchesNames[0]))
#NO HAY UNA API PARA OBTENER LA TENSION BASE DE LA LINEA, POR LO QUE PIDO EL NUMERO DE LA BARRA A LA QUE SE CONECTA, Y AVERIGUO CUAL ES LA TENSION DE ESA BARRA
ierr, branchesFromNodes = psspy.abrnint(-1, string="FROMNUMBER")
branchesVoltageBase = list(map(lambda x: dicBusesRelation[x], branchesFromNodes[0]))
branchesVoltageBase = list(map(lambda x: round(x,1), branchesVoltageBase))
ierr, branchesPowerFlow = psspy.abrnreal(-1, string="MVA")
branchesPowerFlowRounded= list(map(lambda x: round(x,2), branchesPowerFlow[0]))
ierr, branchesRate1MVA = psspy.abrnreal(-1, string="RATE1")
branchesRate1MVARounded= list(map(lambda x: round(x,2), branchesRate1MVA[0]))
ierr, branchesRate2MVA = psspy.abrnreal(-1, string="RATE2")
branchesRate2MVARounded= list(map(lambda x: round(x,2), branchesRate2MVA[0]))
ierr, branchesRate3MVA = psspy.abrnreal(-1, string="RATE3")
branchesRate3MVARounded= list(map(lambda x: round(x,2), branchesRate3MVA[0]))
ierr, branchesRate1Percent = psspy.abrnreal(-1, string="PCTMVARATE1")
branchesRate1PercentRounded = list(map(lambda x: round(x,1), branchesRate1Percent[0]))
ierr, branchesRate2Percent = psspy.abrnreal(-1, string="PCTMVARATE2")
branchesRate2PercentRounded = list(map(lambda x: round(x,1), branchesRate2Percent[0]))
ierr, branchesRate3Percent = psspy.abrnreal(-1, string="PCTMVARATE3")
branchesRate3PercentRounded = list(map(lambda x: round(x,1), branchesRate3Percent[0]))

#VARIABLES DE TRANSFORMADORES 2 devanados
ierr, tr2Names = psspy.atrnchar(-1, string="XFRNAME")
tr2NamesClean = list(map(lambda x: x.strip(), tr2Names[0]))
ierr, tr2Voltage1 = psspy.atrnreal(-1, string="NOMV1")
ierr, tr2Voltage2 = psspy.atrnreal(-1, string="NOMV2")
ierr, tr2MVA = psspy.atrnreal(-1, string="MVA")
tr2MVARounded = list(map(lambda x: round(x,2), tr2MVA[0]))
ierr, tr2RateMVA = psspy.atrnreal(-1, string="RATE1")
tr2RateMVARounded = list(map(lambda x: round(x,2), tr2RateMVA[0]))
ierr, tr2RatePercent = psspy.atrnreal(-1, string="PCTCRPRATE1")
tr2RatePercentRounded = list(map(lambda x: round(x,1), tr2RatePercent[0]))
tr2Voltage = list(map(lambda x,y: str(round(x,1)).replace('.0', '').replace(".",",")+ "/" + str(round(y,1)).replace('.0', '').replace(".",",") ,tr2Voltage1[0], tr2Voltage2[0]))
tr2Devanados = list(map(lambda x: "-" , tr2Names[0]))


#VARIABLES DE TRANSFORMADORES 3 devanados
ierr, tr3Names = psspy.awndchar(-1, string="XFRNAME", entry=2)
tr3NamesClean = list(map(lambda x: x.strip(), tr3Names[0]))
ierr, tr3Devanado = psspy.awndint(-1, string="WNDNUM", entry=2)
ierr, tr3Voltage = psspy.awndreal(-1, string="NOMV", entry=2)
ierr, tr3MVA = psspy.awndreal(-1, string="MVA", entry=2)
tr3MVARounded = list(map(lambda x: round(x,2), tr3MVA[0]))
ierr, tr3RateMVA = psspy.awndreal(-1, string="RATE1", entry=2)
tr3RateMVARounded = list(map(lambda x: round(x,2), tr3RateMVA[0]))
ierr, tr3RatePercent = psspy.awndreal(-1, string="PCTRATE1", entry=2)
tr3RatePercentRounded = list(map(lambda x: round(x,1), tr3RatePercent[0]))
#CREO EL LISTADO DE LOS DEVANADOS Y LA TENSION DE LOS TRANSFORMADORES DE TRES DEVANADOS
tr3DevanadoEditado = list(map(lambda x: "PRIMARIO" if x==1 else "SECUNDARIO" if x==2 else "TERCIARIO", tr3Devanado[0]))
tr3VoltageListado = list(map(lambda x: str(round(tr3Voltage[0][0+3*x],1)).replace('.0', '').replace(".",",")+ "/" + str(round(tr3Voltage[0][1+3*x],1)).replace('.0', '').replace(".",",") + "/" + \
    str(round(tr3Voltage[0][2+3*x],1)).replace('.0', '').replace(".",","), range(int(len(tr3Names[0])/3))))
tr3VoltageListado3 = []
for i in tr3VoltageListado:
    tr3VoltageListado3.append(i)
    tr3VoltageListado3.append(i)
    tr3VoltageListado3.append(i)


#VARIABLES GENERADORES y BARRA SWING
ierr, swingBus = psspy.busmsm(slackBus)
ierr, machineName = psspy.amachint(-1, string="NUMBER")
ierr, machineID = psspy.amachchar(-1, string="ID")
ierr, machineMW = psspy.amachreal(-1, string="PGEN")
machineMW= list(map(lambda x: round(x,2), machineMW[0]))
ierr, machineMVAR = psspy.amachreal(-1, string="QGEN")
machineMVAR= list(map(lambda x: round(x,2), machineMVAR[0]))
machineNameID = list(map(lambda x,y: str(x) + "-" + str(y), machineName[0], machineID[0]))

#CREO UN EXCEL LLAMADO x1 DONDE VOY A CARGAR LOS RESULTADOS DEL FLUJO DE CARGA.
x1 = excelpy.workbook(overwritesheet=True, mode='w') 

#CREO LA HOJA LINEAS
x1.worksheet_rename('LINEAS', 'Sheet1', overwritesheet=True) 
x1.set_active_sheet('LINEAS')
x1.set_cell("a1","NOMBRE")
x1.set_cell("b1","TENSION [kV]")
x1.set_cell("c1","CARGA [MVA]")
x1.set_cell("d1","RATE 1 [MVA]")
x1.set_cell("e1","RATE 1 [%]")
x1.set_cell("f1","RATE 2 [MVA]")
x1.set_cell("g1","RATE 2 [%]")
x1.set_cell("h1","RATE 3 [MVA]")
x1.set_cell("i1","RATE 3 [%]")

x1.set_range(2, 'a', branchesNames, transpose=True)
x1.set_range(2, 'b', branchesVoltageBase, transpose=True)
x1.set_range(2, 'c', branchesPowerFlowRounded, transpose=True)
x1.set_range(2, 'd', branchesRate1MVARounded, transpose=True)
x1.set_range(2, 'e', branchesRate1PercentRounded, transpose=True)
x1.set_range(2, 'f', branchesRate2MVARounded, transpose=True)
x1.set_range(2, 'g', branchesRate2PercentRounded, transpose=True)
x1.set_range(2, 'h', branchesRate3MVARounded, transpose=True)
x1.set_range(2, 'i', branchesRate3PercentRounded, transpose=True)



#CREO LA HOJA TRAFOS
if len(tr2NamesClean) != 0 or len(tr3NamesClean) != 0:
    x1.worksheet_add_end('TRAFOS') 
    x1.set_active_sheet('TRAFOS')
    x1.set_cell("a1","NOMBRE")
    x1.set_cell("b1","TENSION [kV]")
    x1.set_cell("c1","DEVANADO")
    x1.set_cell("d1","CARGA [MVA]")
    x1.set_cell("e1","RATE [MVA]")
    x1.set_cell("f1","RATE [%]")

    if len(tr2NamesClean) != 0:
        x1.set_range(2, 'a', tr2NamesClean, transpose=True)
        x1.set_range(2, 'b', tr2Voltage, transpose=True)
        x1.set_range(2, 'c', tr2Devanados, transpose=True)
        x1.set_range(2, 'd', tr2MVARounded, transpose=True)
        x1.set_range(2, 'e', tr2RateMVARounded, transpose=True)
        x1.set_range(2, 'f', tr2RatePercentRounded, transpose=True)

    if len(tr3NamesClean) != 0:
        x1.set_range(2+len(tr2Names[0]), 'a', tr3NamesClean, transpose=True)
        x1.set_range(2+len(tr2Names[0]), 'b', tr3VoltageListado3, transpose=True)
        x1.set_range(2+len(tr2Names[0]), 'c', tr3DevanadoEditado, transpose=True)
        x1.set_range(2+len(tr2Names[0]), 'd', tr3MVARounded, transpose=True)
        x1.set_range(2+len(tr2Names[0]), 'e', tr3RateMVARounded, transpose=True)
        x1.set_range(2+len(tr2Names[0]), 'f', tr3RatePercentRounded, transpose=True)



#CREO LA HOJA BARRAS
if len(busesNames) != 0:
    x1.worksheet_add_end('BARRAS') 
    x1.set_active_sheet('BARRAS')
    x1.set_cell("a1","NOMBRE")
    x1.set_cell("b1","TENSION BASE\n[kV]")
    x1.set_cell("c1","TENSION\n[kV]")
    x1.set_cell("d1","TENSION\n[PU]")
    x1.set_cell("e1","DESVIO DE TENSION [%]")
    x1.set_cell("f1","∆U MAX.\nPERMITIDO [±%]")

    if busesWithNames:  #VERIFICO SI LAS BARRAS TIENEN NOMBRE, SI NO TIENEN NOMBRE, PONGO LOS NUMEROS DE BARRA
        x1.set_range(2, 'a', busesNames, transpose=True)
    else:
        x1.set_range(2, 'a', busesNumbers, transpose=True)
    x1.set_range(2, 'b', busesVoltageBaseRounded, transpose=True)
    x1.set_range(2, 'c', busesVoltageRounded, transpose=True)
    x1.set_range(2, 'd', busesVoltagePURounded, transpose=True)
    x1.set_range(2, 'e', busesVoltageDiffRounded, transpose=True)
    x1.set_range(2, 'f', busesVoltageDiffMax, transpose=True)


#CREO LA HOJA GENERADORES
if len(machineNameID) != 0:
    x1.worksheet_add_end('GENERADORES') 
    x1.set_active_sheet('GENERADORES')
    x1.set_cell("a1","NOMBRE")
    x1.set_cell("b1","POTENCIA ACTIVA [MW]")
    x1.set_cell("c1","POTENCIA REACTIVA [MVAr]")
    x1.set_cell("a2","BARRA SWING")
    x1.set_cell("b2",round(swingBus.real,2))
    x1.set_cell("c2",round(swingBus.imag,2))
    x1.set_range(3, 'a', machineNameID, transpose=True)
    x1.set_range(3, 'b', machineMW, transpose=True)
    x1.set_range(3, 'c', machineMVAR, transpose=True)


x1.save(nameCasePsse.strip(".sav") + ".xlsx")
x1.close()
x1.close_app()
print('EJECUTADO CON EXITO')