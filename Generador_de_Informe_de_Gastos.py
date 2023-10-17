import openpyxl as op
arhivo = "./informe_gastos.xlsx"
libro_excel = op.load_workbook(arhivo)
hoja = libro_excel.active
hoja.title = "Gastos"

Datos = []

while True:
    Gastos = float(input("Ingrese los detalles del gastos: "))
    Fecha = input("Ingrese la fecha DIA/MES/AÑO:")
    Descripcion = input("Ingrese descripcion:")
    Monto = float(input("Ingrese el monto de los gastos:"))
    
    Datos.append((Gastos,Fecha,Descripcion,Monto))
    
    continuar = input("¿Desea agregar otro gasto? (S/N): ").lower()
    if continuar != 's':
        break
hoja.append(["Gastos","Fecha","Descripcion","Monto"])
for Gastos,Fecha,Descripcion,Monto in Datos:
    hoja.append([Gastos,Fecha,Descripcion,Monto])

Monto_total=sum(Monto for _,_,_, Monto in Datos)
gasto_mas_caro=max(Datos,key=lambda x: x[0])
gasto_mas_barato=min(Datos,key=lambda x: x[0])

print(f"Número total de gastos: {len(Datos)}")
print(f"Gasto más caro: Fecha: {gasto_mas_caro[1]}, Descripción: {gasto_mas_caro[2]}, Monto: {gasto_mas_caro[3]}")
print(f"Gasto más barato: Fecha: {gasto_mas_barato[1]}, Descripción: {gasto_mas_barato[2]}, Monto: {gasto_mas_barato[3]}")
print(f"Monto total de gastos: {Monto_total}")

libro_excel.save(arhivo)
print("Informe de gastos guardado en 'informe_gastos.xlsx'")