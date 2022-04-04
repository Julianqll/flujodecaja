from flask import Flask, json, send_file, flash, jsonify, redirect, render_template, request, session
app = Flask(__name__)
import json
import xlwt
from xlwt import Workbook


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return render_template("index.html")
    else:
        financimiento = request.form.get("financiamiento")
        tiempo = request.form.get("tiempo")
        return redirect("/flujo/"+str(financimiento)+"/"+str(tiempo))


@app.route("/flujo/<financiamiento>/<tiempo>", methods=["GET", "POST"])
def flujo(financiamiento, tiempo):
    if request.method == "GET":
        return render_template("flujo.html", tiempo=tiempo, financiamiento=financiamiento)


@app.route("/flujocaja", methods=["GET","POST"])
def flujocaja():
    requerimientos = json.loads(request.form["reqs"])
    ingresos = json.loads(request.form["ingresos"])
    egresos = json.loads(request.form["egresos"])
    tiempo = request.form["tiempo"]
    financiamiento = request.form["financiamiento"]
    wb = Workbook()
    #requerimientos
    sheet1 = wb.add_sheet("Requerimientos")
    j=1
    reqs_resumen=[]
    reqs_total=[]
    for key, value in requerimientos.items():
        sheet1.write(1, j+1, 'Requerimientos - '+key)
        reqs_resumen.append(key) 
        sheet1.write(2,j+1, 'Item')
        sheet1.write(2,j+2, 'Cantidad')
        sheet1.write(2,j+3, 'Precio')
        sheet1.write(2,j+4, 'Total')
        i = 3
        sum=0
        for key1, value1 in value.items():
            sheet1.write(i,j+1,key1)
            sheet1.write(i,j+2, value1[0])
            sheet1.write(i,j+3, value1[1])
            sheet1.write(i,j+4, int(value1[0])*float(value1[1]))
            sum = sum + int(value1[0])*float(value1[1])
            i=i+1
        sheet1.write(i,j+1,'Total')
        sheet1.write(i,j+4,sum)
        reqs_total.append(sum)
        j=j+5


    #inversion inicial
    periodocero= 0
    sheet2 = wb.add_sheet("Inversion Inicial")
    sheet2.write(1, 1, 'Periodo 0')
    sheet2.write(2, 1, 'Concepto') 
    sheet2.write(2, 2, 'Precio total')
    j = 3
    for i in range(len(reqs_total)):
        j=j+i
        sheet2.write(j,1, reqs_resumen[i])
        sheet2.write(j,2, reqs_total[i])
        periodocero = periodocero + float(reqs_total[i])


    #ingresos
    sheet3 = wb.add_sheet("Ingresos")
    sheet3.write(1, 1, 'Ingresos totales')
    sheet3.write(2, 1, 'Concepto')
    for i in range(int(tiempo)):
        sheet3.write(2, i+2, str(i+1)+'° mes') 
    count = 0
    j = 3
    subtotales = [0]*int(tiempo)
    cantidades = []
    for key, value in ingresos.items():
        if '-' in key:
            sum = 0
            sheet3.write(j,1, "Precio item: "+key)
            for i in range(int(tiempo)):
                sheet3.write(j,i+2, float(value))
                subtotales[i]= subtotales[i] + cantidades[i]*float(value)
        else:
            cantidades = []
            sheet3.write(j,1, "Cant. item: "+key)
            for i in range(int(tiempo)):
                sheet3.write(j,i+2, int(value[i]))
                cantidades.append(int(value[i]))
        j=j+1
    sheet3.write(j,1, "Totales")
    for i in range(int(tiempo)):
        sheet3.write(j,i+2, subtotales[i])

    #egresos
    sheet4 = wb.add_sheet("Egresos")
    sheet4.write(1, 1, 'Egresos totales')
    sheet4.write(2, 1, 'Concepto')
    for i in range(int(tiempo)):
        sheet4.write(2, i+2, str(i+1)+'° mes')
    j = 3
    sumas = [0]*int(tiempo)
    for key, value in egresos.items():
        sheet4.write(j,1, key)
        for i in range(int(tiempo)):
            sheet4.write(j,i+2, float(value[i]))
            sumas[i] = sumas[i] + float(value[i])
        j=j+1         
    sheet4.write(j,1, "Totales")
    for i in range(int(tiempo)):
        sheet4.write(j,i+2, sumas[i])

    #FLUJO DE OPERACION
    sheet5 = wb.add_sheet("Flujo de operación")
    sheet5.write(1,1,"Flujo de caja de operación")
    for i in range(int(tiempo)):
        sheet5.write(2, i+2, str(i+1)+'° mes')
    sheet5.write(3,1,"Ingresos por ventas")
    sheet5.write(4,1,"Egresos totales")
    sheet5.write(5,1,"Saldos operativos mensuales")
    sheet5.write(6,1,"Saldos operativos acumulados")
    #ingresosporventas
    acumulados = []
    for j in range(int(tiempo)):
        sheet5.write(3,j+2, subtotales[j])
    for j in range(int(tiempo)):
        sheet5.write(4,j+2, sumas[j])
    for j in range(int(tiempo)):
        sheet5.write(5,j+2, subtotales[j]-sumas[j])
        acumulados.append(subtotales[j]-sumas[j])
    sumaacu = 0
    for i in range(int(tiempo)):
        sheet5.write(6,i+2, acumulados[i]+ sumaacu)
        sumaacu = sumaacu + acumulados[i]
    
    prestamo = 585.48
    prestamototal = 18144.81
    aportecap = 27217.22

    #adicional
    if int(financiamiento) == 1:
        print("ok")
        #cuadro de financiamiento
        sheet6 = wb.add_sheet("Inversion y financiamiento")
        sheet6.write(1,1,"Inversion y financiamiento")
        sheet6.write(2,1,"Concepto")
        sheet6.write(2,2,"Periodo 0")
        for i in range(int(tiempo)):
            sheet6.write(2, i+3, str(i+1)+'° mes')
        sheet6.write(3,1,"Inversion inicial")
        sheet6.write(4,1,"Aporte capital")
        sheet6.write(5,1,"Prestamo bancario")
        sheet6.write(6,1,"Pago a los socios")
        sheet6.write(7,1,"Pago del préstamo")
        sheet6.write(8,1,"Saldo de inversion y financiamiento")
        sheet6.write(3,2,-1*float(aportecap)-float(prestamototal))
        sheet6.write(4,2,float(aportecap))
        sheet6.write(5,2,float(prestamototal))
        totalprest = 0
        for j in range(int(tiempo)):
            sheet6.write(7,j+3, -1*float(prestamo))
            sheet6.write(8,j+3,-1*float(prestamo))
            totalprest = totalprest - float(prestamo)
        sheet6.write(2,int(tiempo)+3,"Total")
        sheet6.write(7,int(tiempo)+3,totalprest)


        #flujo de operación con financiamiento
        sheet7 = wb.add_sheet("Flujo de op financiamiento")
        sheet7.write(1,1,"Flujo de caja de operación con financiamiento")
        for i in range(int(tiempo)):
            sheet7.write(2, i+2, str(i+1)+'° mes')
        sheet7.write(3,1,"Saldos operativos mensuales")
        sheet7.write(4,1,"Saldos mensuales de inversion")
        sheet7.write(5,1,"Saldos totales de efectivo")
        sheet7.write(6,1,"Saldos totales acumulados")
        #ingresosporventas
        nuevoacumulados = []
        for j in range(int(tiempo)):
            sheet7.write(3,j+2, acumulados[j])
        for j in range(int(tiempo)):
            sheet7.write(4,j+2, float(prestamo)*-1)
        for j in range(int(tiempo)):
            sheet7.write(5,j+2, acumulados[j]-float(prestamo))
            nuevoacumulados.append(acumulados[j]-float(prestamo))
        sumaacu = 0
        for i in range(int(tiempo)):
            sheet7.write(6,i+2, nuevoacumulados[i]+ sumaacu)
            sumaacu = sumaacu + nuevoacumulados[i]

    #FLUJO DE OPE E INVERSION
    if int(financiamiento) == 0:
        acumuladosinver = 0
        sheet6 = wb.add_sheet("Flujo de operación e inversion")
        sheet6.write(1,1,"Flujo de caja de operación e inversion")
        sheet6.write(2,1,"Concepto")
        sheet6.write(2,2,"Periodo 0")
        for i in range(int(tiempo)):
            sheet6.write(2, i+3, str(i+1)+'° mes')
        sheet6.write(3,1,"Saldos mensuales de operaciones")
        sheet6.write(4,1,"Saldos mensaules de inversion y fin.")
        sheet6.write(5,1,"Saldos totales de efectivo por mes")
        sheet6.write(6,1,"Saldos totales acumulados")
        sheet6.write(4,2,-1*periodocero)
        sheet6.write(5,2,-1*periodocero)
        sheet6.write(6,2,-1*periodocero)
        acumuladosinver = acumuladosinver+ -1*periodocero
        for i in range(int(tiempo)):
            sheet6.write(3, i+3,acumulados[i])
            sheet6.write(5, i+3,acumulados[i])
            sheet6.write(6, i+3,acumulados[i]+acumuladosinver)
            acumuladosinver = acumuladosinver + acumulados[i]
    else:
        acumuladosinver = 0
        sheet8 = wb.add_sheet("Flujo de operación e inversion")
        sheet8.write(1,1,"Flujo de caja de operación e inversion")
        sheet8.write(2,1,"Concepto")
        sheet8.write(2,2,"Periodo 0")
        for i in range(int(tiempo)):
            sheet8.write(2, i+3, str(i+1)+'° mes')
        sheet8.write(3,1,"Saldos mensuales de operaciones")
        sheet8.write(4,1,"Pago prestamos")
        sheet8.write(5,1,"Inversion")
        sheet8.write(6,1,"Saldos de operación e inversion")
        sheet8.write(7,1,"Saldos totales acumulados")
        sheet8.write(5,2,-1*periodocero)
        sheet8.write(6,2,-1*periodocero)
        sheet8.write(7,2,-1*periodocero)
        acumuladosinver = acumuladosinver+ -1*periodocero
        for i in range(int(tiempo)):
            sheet8.write(3, i+3,acumulados[i])
            sheet8.write(4, i+3, -1*float(prestamo))
            sheet8.write(6, i+3,acumulados[i]-float(prestamo))
            acumulados[i] = acumulados[i] -float(prestamo)
            sheet8.write(7, i+3,acumulados[i]+acumuladosinver)
            acumuladosinver = acumuladosinver + acumulados[i]











    wb.save('flujodecaja.xls')
    path = "flujodecaja.xls"
    return send_file(path, as_attachment=True)