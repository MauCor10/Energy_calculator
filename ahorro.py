from tkinter import *
import os
import xlsxwriter

ventana = Tk()

# -----inicializar variables----#
tot_e = 0
tot_d = 0
tot_co2 = 0
tot_arboles = 0
pot = DoubleVar()
horas = DoubleVar()
cant = IntVar()
precio = DoubleVar()
fact_emision = DoubleVar()
days = IntVar()
production_annual = IntVar()

# -----calculos----#
def ahorro():
    ahorro_e = round(pot.get() / 1000 * horas.get() * days.get() * cant.get(), 2)
    preckwh = precio.get()
    ahorro_d = round(ahorro_e * preckwh, 2)
    facemco2 = fact_emision.get()
    ahorro_co2 = round(facemco2 * ahorro_e, 2)
    arboles = round(ahorro_co2 / 85, 2)
    kpi_a = round(ahorro_e*1000/production_annual.get())

    resultado_e = Label(frame, text=("The annual energy savings will be: " + str(ahorro_e)) + " kWh", fg="white", highlightthickness=1,
                        highlightbackground = "white", highlightcolor= "white", bg="darkblue",
                        font=("calibri", 12,'bold'))
    resultado_e.grid(row=4, column=0, sticky="e", padx=70)

    resultado_d = Label(frame, text=("The annual economic savings will be: $ " + str(ahorro_d)), fg="white", highlightthickness=1,
                        highlightbackground = "white", highlightcolor= "white", bg="darkblue",
                        font=("calibri", 12,'bold'))
    resultado_d.grid(row=5, column=0, sticky="e", padx=70)

    resultado_a = Label(frame, text=("Impact on annual KPI: -" + str(kpi_a)) + " MWh/car", fg="white", highlightthickness=1,
                        highlightbackground = "white", highlightcolor= "white", bg="darkblue", font=("calibri",12,'bold'))
    resultado_a.grid(row=6, column=0, sticky="e", padx=70)

    resultado_co2 = Label(frame, text=("The reduction of CO2 emissions will be: " + str(ahorro_co2)) + " kg", fg="white",
                          highlightthickness=1, highlightbackground = "white", highlightcolor= "white", bg="darkblue", font=("calibri", 12,'bold'))
    resultado_co2.grid(row=7, column=0, sticky="e", padx=70)

    resultado_arboles = Label(frame, text=("Same annual CO2 reduction (average): " + str(arboles)) + " trees", fg="white",
                              highlightthickness=1, highlightbackground = "white", highlightcolor= "white", bg="darkblue",
                              font=("calibri", 12,'bold'))
    resultado_arboles.grid(row=8, column=0, sticky="e", padx=70)

# ------para .exe-----#
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath("../Curso")

    return os.path.join(base_path, relative_path)

# ------exportar-----#
def export():
    ahorro_e = round(pot.get() / 1000 * horas.get() * days.get() * cant.get(), 2)
    preckwh = precio.get()
    ahorro_d = round(ahorro_e * preckwh, 2)
    facemco2 = fact_emision.get()
    ahorro_co2 = round(facemco2 * ahorro_e, 2)
    arboles = round(ahorro_co2 / 85, 2)
    kpi_a = round(ahorro_e * 1000 / production_annual.get())
    workbook = xlsxwriter.Workbook('Energy calculator.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.set_column('A:A', 20)
    worksheet.write('A7', 'Action')
    worksheet.write('B7', 'Annual saving energy (kWh)')
    worksheet.write('C7', 'Annual saving economic ($)')
    worksheet.write('D7', 'KPI Impact (MWh/car)')
    worksheet.write('E7', 'Reduction CO2 (kg)')
    worksheet.write('F7', 'Reduction in number of trees')
    worksheet.write('B8', ahorro_e)
    worksheet.write('C8', ahorro_d)
    worksheet.write('D8', -kpi_a)
    worksheet.write('E8', ahorro_co2)
    worksheet.write('F8', arboles)
    worksheet.set_column('A:A', 46)
    worksheet.set_column('B:B', 26)
    worksheet.set_column('C:C', 25)
    worksheet.set_column('D:D', 20)
    worksheet.set_column('E:E', 18)
    worksheet.set_column('F:F', 26)
    worksheet.insert_image('A1', 'stellantis.PNG', {'x_scale': 0.5, 'y_scale': 0.5})
    workbook.close()

# ------barra superior-----#
ventana.title("Energy Calculator")
ventana.config(borderwidth=1,relief="solid")

# ------logo------#
path = resource_path("logo.ico")
ventana.iconbitmap(path)

# -----frame-----#
frame = Frame()
frame.pack()
frame.config(width="1000", height="1000")

# --------fondo------#
bg = PhotoImage(file="fondo.PNG")
fondo = Label(frame, image = bg)
fondo.place(x=0,y=-500)

# --------imagen------#
path = resource_path("stellantis.PNG")
logostellantis = PhotoImage(file=path)
etiquetalogo = Label(frame, image=logostellantis, fg="lightblue", bg="darkblue").grid(row=2, column=0)

# -------titulo y subtitulo------#
titulo = Label(frame, text="Application for energy savings calculations", highlightthickness=1,
               highlightbackground = "darkblue", highlightcolor= "darkblue", bg="#FFFFFF", fg="darkblue", font=("calibri", 22))
titulo.grid(pady=5)
made_by = Label(frame, text="Made by:", fg="white", bg="darkblue", font=("calibri", 8))
titulo.grid(row=0, column=0)
made_by.grid(row=2, column=0, sticky="n")
intro = Label(frame,
              text="Software for calculating energy savings and their economic, productive and environmental impacts",
              highlightthickness=1, highlightbackground = "black", highlightcolor= "black", bg="#FFFFFF",
             fg="black", font=("calibri", 12))
intro.grid(row=3, column=0, pady=25)

# ------etiquetas y textos-----"
potencia_label = Label(frame, text="Set the power of the device you want to turn off (Watts):", highlightthickness=1,
                       highlightbackground = "darkblue", highlightcolor= "darkblue", bg="#FFFFFF", fg="darkblue",
                       font=("calibri", 12, "bold"))
potencia_label.grid(row=4, column=0, sticky="w", padx=45, pady=10)

horas_label = Label(frame, text="Set the daily shutdown hours of the device (hours):", highlightthickness=1,
                    highlightbackground = "darkblue", highlightcolor= "darkblue", bg="#FFFFFF", fg="darkblue",
                    font=("calibri", 12, "bold"))
horas_label.grid(row=5, column=0, sticky="w", padx=87, pady=10)

cantidad_label = Label(frame, text="If multiple devices are the same, enter the total amount:", highlightthickness=1,
                       highlightbackground = "darkblue", highlightcolor= "darkblue", bg="#FFFFFF", fg="darkblue",
                       font=("calibri", 12,"bold"))
cantidad_label.grid(row=6, column=0, sticky="w", padx=52, pady=10)

potencia_cuadro = Entry(frame, textvariable=pot, highlightthickness=1, highlightbackground = "darkblue",
                        highlightcolor= "darkblue", bg="#FFFFFF")
potencia_cuadro.grid(row=4, column=0, padx=450, pady=10)

horas_cuadro = Entry(frame, textvariable=horas, highlightthickness=1, highlightbackground = "darkblue",
                     highlightcolor= "darkblue", bg="#FFFFFF")
horas_cuadro.grid(row=5, column=0, pady=10)

cantidad_cuadro = Entry(frame, textvariable=cant, highlightthickness=1, highlightbackground = "darkblue",
                        highlightcolor= "darkblue", bg="#FFFFFF")
cantidad_cuadro.grid(row=6, column=0, pady=10)

precio_label = Label(frame, text="Set energy price ($/kWh):", highlightthickness=1, highlightbackground = "darkblue",
                     highlightcolor= "darkblue", bg="#FFFFFF", fg="darkblue",
                     font=("calibri", 12,"bold"))
precio_label.grid(row=7, column=0, sticky="w", padx=263, pady=10)

precio_cuadro = Entry(frame, textvariable=precio, highlightthickness=1, highlightbackground = "darkblue",
                      highlightcolor= "darkblue", bg="#FFFFFF")
precio_cuadro.grid(row=7, column=0, padx=63, pady=10)

planta_label = Label(frame, text="Annual production days:", highlightthickness=1, highlightbackground = "darkblue",
                     highlightcolor= "darkblue", bg="#FFFFFF", fg="darkblue", font=("calibri", 12,"bold"))
planta_label.grid(row=8, column=0, sticky="w", padx=270, pady=10)
planta_cuadro = Entry(frame, textvariable=days, highlightthickness=1, highlightbackground = "darkblue",
                      highlightcolor= "darkblue", bg="#FFFFFF")
planta_cuadro.grid(row=8, column=0, pady=10)

prodan_label = Label(frame, text="Set the annual production (quantity of cars budget):", highlightthickness=1,
                     highlightbackground = "darkblue", highlightcolor= "darkblue", bg="#FFFFFF", fg="darkblue", font=("calibri", 12,"bold"))
prodan_label.grid(row=9, column=0, sticky="w", padx=79, pady=10)
prodan_cuadro = Entry(frame, textvariable=production_annual, highlightthickness=1, highlightbackground = "darkblue",
                      highlightcolor= "darkblue", bg="#FFFFFF")
prodan_cuadro.grid(row=9, column=0,pady=10)

#----opciones----#
emision_label = Label(frame, text="Select your country:", highlightthickness=1, highlightbackground = "darkblue",
                      highlightcolor= "darkblue", bg="#FFFFFF", fg="darkblue",
                      font=("calibri",12,"bold"))
emision_label.grid(row=10,column=0, sticky="w", padx=300, pady=10)
argentina = Radiobutton(ventana, text="Argentina", variable=fact_emision, value=0.227, borderwidth=1, relief="solid",
                        bg="#FFFFFF").place(x=450, y=630)
brasil = Radiobutton(ventana, text="Brasil", variable=fact_emision, value=0.069, borderwidth=1, relief="solid",
                     bg="#FFFFFF").place(x=450, y=660)

# ----boton calcular-----#
boton_calcular = Button(frame, text="Calculate", command=ahorro, fg="white", bg="darkblue", font=("calibri",11,"bold"))
boton_calcular.grid(row=11, column=0, sticky="w",padx=210, pady=5)

# ----boton exportar-----#
boton_calcular = Button(frame, text="Export to Excel", command=export, fg="white", bg="darkblue", font=("calibri",11,"bold"))
boton_calcular.grid(row=11, column=0)

ventana.mainloop()
