import APP as app
from tkinter import ttk,Tk,Menu,messagebox
import openpyxl as op
import datetime



def cargar_todos_archivos():
        global seven, pure
        seven = app.Archivo_seven('/gastos_seven.xlsx')
        pure = app.Archivo_pure('/gastos_pure.xlsx') 

        if seven.wb != None and pure.wb != None and len(seven.proyectos.proyectos)>0:
            messagebox.showinfo('seven-->pure','Los archivos se cargaron correctamente')
            options = menu_archivo.index('end')
            for option in range(options-1):
                 name = menu_archivo.entrycget(option,'label')
                 if menu_archivo.entrycget(name,'state') == 'disabled':
                    menu_archivo.entryconfig(name,state='normal')
                 
        else:
            messagebox.showerror('seven-->pure',app.creacion) 

def salir():
     root.destroy()

def errores():
      if len(seven.registros_error)>0 :
        messagebox.showerror('seven-->pure','se presentaron '+str(len(seven.registros_error))+' errores')
        wb = op.Workbook()
        ws = wb.active
        encabezado = ['fila','id','cuenta','debito','error']
        ws.append(encabezado)
        for dato in seven.registros_error:
             ws.append(dato)
        fecha = str(datetime.datetime.now().strftime("%d%m%Y %H%M%S"))     
        wb.save('Reporte errores ' + fecha +'.xlsx')
          

def registrar():
    pure.verificar_registros(seven) 
    if app.num == 0 and len(seven.registros_error)==0  :
        messagebox.showinfo('seven-->pure','No existen registros nuevos')
    else:      
        messagebox.showinfo('seven-->pure',str(app.num) + ' gastos registrados')
    errores()         

def exportar():
    if pure.last_row > 2:
          pure.exportar()
          messagebox.showinfo('seven-->pure','Reporte pure fue generado con exito')
    else:
          messagebox.showinfo('seven-->pure','Reporte pure esta vacio')
             

root = Tk()
root.title('Migracion Seven a Pure')
root.minsize(400,400)

barra_menu = Menu()
root.config(menu=barra_menu)


menu_archivo = Menu(tearoff=0)
menu_archivo.add_command(label='Cargar',command = cargar_todos_archivos)
menu_archivo.add_command(label='Registrar',command = registrar ,state='disabled')
# menu_archivo.add_command(label='ver registros',command = cargar_todos_archivos,state='disabled')
menu_archivo.add_command(label='Exportar',command = exportar,state='disabled')

menu_archivo.add_separator()
menu_archivo.add_command(label='salir',command = salir)
barra_menu.add_cascade(label='Archivo', menu=menu_archivo)

root.mainloop()


