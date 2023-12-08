import tkinter as tk
from tkinter import filedialog
import customtkinter
import threading
import requests
import xlsxwriter
import pandas as pd
import os.path
from PIL import Image


API_TOKEN = ''  #initialization of the token API string
SELECTED_FORMAT = 1 # 1=default , 2=IAE, 3=personalized
MAIL_CHECKBOX = False # boolean value for personal email checkbox
CREDIT_NECESSARY = 0 #initialization of the necesary credit amount
CREDIT_AMOUNT = 0 # initialization of credit-balance
PERCENTAGE_PER_PROFILE = 0 #initialization of value for linear increment of progress bar
PASSED = False #fix for progressbar when it resets
GIF_FRAMES = [] #Frames of loading circle gif
FRAME_DELAY = 0 #Frame delay

#File names
RESULT_FILE_NAME = 'Resultados' # name of the result file for the default and personalized formats
ACADEMIC_RESULT_FILE_NAME = 'Resultados_Historial_Academico' # name of the academic result file for the IAE format
EXPERIENCE_RESULT_FILE_NAME = 'Resultados_Datos_Laborales'   # name of the experience result file for the IAE format

#File paths
SELECTED_FILE_PATH = '' # initialization of the selected excel file path
SELECTED_FOLDER_PATH = '' # initialization of the selected folder path

#Iterations
ITERACIONES_EDUCACION = 5 # amount of iterations for educations
ITERACIONES_EXPERIENCIAS = 5 # amount of iterations for experiences
ITERACIONES_IDIOMA = 3 # amount of iterations for languages
ITERACIONES_MAILS = 3 # amount of iterations for personal emails

#     ----------------------------------------  DEFAULT ------------------------------------------

GENERAL_COLUMNS = ['Link','Nombre','Apellido','Puesto Actual','Pais','Descripcion','Resumen']
GENERAL_COLUMNS_PROXYCURL_EQUIVALENT = ['first_name', 'last_name', 'occupation', 'country_full_name', 'headline', 'summary']

EDUCATION_COLUMNS = ['Titulo', 'Universidad','Descripcion','Campo de estudio','Promedio','Fecha de Inicio','Fecha de Finalizacion']
EDUCATION_COLUMNS_PROXYCURL_EQUIVALENT = ['degree_name', 'school', 'description', 'field_of_study', 'grade', 'starts_at','ends_at']

EXPERIENCE_COLUMNS = ['Puesto', 'Empresa','Descripcion de Empresa','Fecha de Inicio','Fecha de Finalizacion']
EXPERIENCE_COLUMNS_PROXYCURL_EQUIVALENT = ['title', 'company', 'description', 'starts_at','ends_at']

#      ----------------------------------------  IAE   ------------------------------------------

IAE_ACADEMIC_COLUMNS = ['Preferente','Título Universitario','Titulo Univ. Manual','Título de Posgrado','Pais','Becado','Comentario','Historial academico inscripcion','Contacto','Universidad','Institucion Manual','Año de Comienzo','Nombre','Aplazos','Año de Finalización','Promedio','iaetitulo','Tipo Institucion Manual']
IAE_ACADEMIC_COLUMNS_PROXYCURL_EQUIVALENT = [None, None, 'degree_name', None, None, None,'description',None,'ContactID',None, 'school','start_year',None,None,'end_year','grade',None,None]

IAE_EXPERIENCE_COLUMNS = ['Empresa CRM','Empresa Manual','Año de Comienzo',	'Año de Finalización','Contacto','año','Cargo de Superior','Remuneracion Neta','Responsabilidades','Comentario','Nombre','Actual','Cargo','Cargo Codificado','Área','Área Codificada','Experiencia Laboral','Mes','Datos laborales inscripcion','Condicion Laboral','Rango de ingresos','Industria','Sub-Área','Salario Bruto Anual (U$D)','# Empleados a Cargo','Fecha de Ingreso']
IAE_EXPERIENCE_COLUMNS_PROXYCURL_EQUIVALENT = [None, 'company','start_year','end_year','ContactID','end_year',None,None,None,'description',None,'actual',None, 'title',None,None,None,'end_month',None,None,None,None,None,None,None,'starts_at']

#      ---------------------------------------------------------   PERSONALIZED   ------------------------------------

PERSONALIZED_COLUMNS = ['Link','Nombre','Apellido','Nombre Completo', 'Cantidad de Seguidores','Puesto Actual','Descripcion','Resumen','Pais','Pais completo','Ciudad','Provincia','Cantidad de Conexiones','Idiomas', 'Educacion_Titulo', 'Educacion_Universidad','Educacion_Descripcion','Educacion_Campo de estudio','Educacion_Promedio','Educacion_Fecha de Inicio','Educacion_Fecha de Fin','Experiencia_Puesto', 'Experiencia_Empresa','Experiencia_Descripcion','Experiencia_Fecha de Inicio','Experiencia_Fecha de Fin']
PERSONALIZED_COLUMNS_PROXYCURL_EQUIVALENT = ['Link', 'first_name','last_name','full_name','follower_count', 'occupation', 'headline', 'summary','country','country_full_name','city','state','connections','languages','degree_name', 'school', 'description', 'field_of_study', 'grade', 'starts_at','ends_at','title', 'company', 'description', 'starts_at','ends_at']
PERSONALIZED_COLUMNS_IS_CHECKED = [False] * len(PERSONALIZED_COLUMNS)


#This function reads the amount of rows that are in an excel file and stores the number in the CREDIT_NECESSARY variable
#also obtains the percentage_per_second
def amountOfNecessaryCredits():
    if os.path.isfile(SELECTED_FILE_PATH):
        df = pd.read_excel(SELECTED_FILE_PATH)
        global CREDIT_NECESSARY
        CREDIT_NECESSARY = len(df.index)
        global PERCENTAGE_PER_PROFILE
        PERCENTAGE_PER_PROFILE = 100 / CREDIT_NECESSARY
    else:
        CREDIT_NECESSARY = 0

#Function to configure amount of columns and rows for Tkinter grid
def configureColumnAndRow(frame, cols, rows):
    for i in range(cols):
        frame.columnconfigure(i, weight=1)
    for i in range(rows):
        frame.rowconfigure(i, weight=1)

# Wizard Page 1: Welcome Page
class WizardPage1:

    def __init__(self, master, button_command):
        # Initialize the page
        self.master = master
        self.frame = customtkinter.CTkFrame(self.master)
        self.frame.grid(row=0, column=0, sticky="nsew")

        configureColumnAndRow(self.master, 1, 1)
        configureColumnAndRow(self.frame, 1, 3)

        #create title
        self.title = customtkinter.CTkLabel(self.frame, text="Bienvenido!", anchor="center",font=('',30))
        self.title.grid(row=0, column=0, padx=10, pady=20)

        #create body
        self.description = customtkinter.CTkLabel(self.frame, text="El objetivo de este wizard es ayudarte a obtener los datos de LinkedIn usando tu cuenta de Proxycurl",font=('',20),wraplength=600)
        self.description.grid(row=1, column=0, padx=10, pady=20)

        #create next button
        self.button = customtkinter.CTkButton(master=self.frame, text="Siguiente", command=button_command)
        self.button.grid(row=2, column=0)

# Wizard Page 2: Enter API Token
class WizardPage2:
    
    def __init__(self, master, button_command_next, button_command_previous):
        # Initialize the page
        self.master = master
        self.frame = customtkinter.CTkFrame(self.master)
        self.frame.grid(row=0, column=0, sticky="nsew")
        configureColumnAndRow(self.master, 1, 1)
        configureColumnAndRow(self.frame, 2, 3)

        #initialize variables
        self.isTokenValid = False
        self.stop_gif = False
        self.frame_index = 0

        #create title
        self.title = customtkinter.CTkLabel(self.frame, text="Escribir la API token", anchor="center", font=('', 30))
        self.title.grid(row=0, column=0, columnspan=2, padx=10, pady=20) 


        #create text box entry
        sv = tk.StringVar()
        sv.trace("w", lambda name, index, mode, sv=sv: self.on_text_change(sv))
        self.text_box = customtkinter.CTkEntry(master=self.frame, textvariable = sv, font=('', 16), width=250)
        self.text_box.grid(row=1, column=0,columnspan=2) 

        #create validate button
        self.button_validate = customtkinter.CTkButton(master=self.frame, text="Validar", command=self.on_validate, width=90)
        self.button_validate.grid(row=1, column=0,columnspan=2, sticky='e', padx=120)

        #create previous button
        self.button_previous = customtkinter.CTkButton(master=self.frame, text="Atras", command=button_command_previous)
        self.button_previous.grid(row=2, column=0, padx=20) 
        #create next button
        self.button_next = customtkinter.CTkButton(master=self.frame, text="Siguiente", command=button_command_next, state='disabled')
        self.button_next.grid(row=2, column=1, padx=20) 

    #updates API_TOKEN variable when the text box changes
    def on_text_change(self, text):
        global API_TOKEN
        API_TOKEN = text.get().strip()
        #disable next button
        self.button_next.configure(state="disabled")
    
    #starts playing gif inside button
    def on_validate(self):
        self.button_validate.grid_forget()
        self.button_validate = customtkinter.CTkButton(master=self.frame, text="", width=70, state='disabled')
        self.button_validate.grid(row=1, column=1)
        self.stop_gif = False
        self.play_gif()
        self.worker_thread = threading.Thread(target=self.getCredits, daemon=True)
        self.worker_thread.start()
        self.check_for_completion()

    #returns the button to its original state
    def on_return_to_validate(self):
        self.button_validate.grid_forget()
        self.button_validate = customtkinter.CTkButton(master=self.frame, text="Validar", command=self.on_validate, width=90)
        self.button_validate.grid(row=1, column=0,columnspan=2, sticky='e', padx=120)

    #every 0.3s checks if the thread is still running
    def check_for_completion(self):
        if self.worker_thread.is_alive():
            #schedule another check
            self.master.after(300, self.check_for_completion)
        else:
            #finished operation, stop gif
            self.stop_gif = True 

    #starts playing gif
    def play_gif(self):
        if self.frame_index >= len(GIF_FRAMES):
            self.frame_index = 0
        else:
            current_frame = customtkinter.CTkImage(GIF_FRAMES[self.frame_index])
            self.frame_index += 1
            self.button_validate.configure(image=current_frame, text="")
        if not self.stop_gif:
            self.frame.after(FRAME_DELAY, self.play_gif)
        else:
            #stop gif
            if self.isTokenValid:
                #add check image to button
                button_image = customtkinter.CTkImage(Image.open("resources/check.png"))
                self.button_validate.grid_forget()
                self.button_validate = customtkinter.CTkButton(master=self.frame,image=button_image, text="", width=70,command=self.on_return_to_validate)
                self.button_validate.grid(row=1, column=1)

                self.button_next.configure(state="normal")
                
            else:
                #add cross image to button
                button_image = customtkinter.CTkImage(Image.open("resources/cross.webp"))
                self.button_validate.grid_forget()
                self.button_validate = customtkinter.CTkButton(master=self.frame,image=button_image, text="", width=70,command=self.on_return_to_validate)
                self.button_validate.grid(row=1, column=1)
        
    #Does an http request to the Proxycurl API and returns the credit-balance of the user
    def getCredits(self):
        headers = {'Authorization': 'Bearer ' + API_TOKEN}
        api_endpoint = 'https://nubela.co/proxycurl/api/credit-balance'
        response = requests.get(api_endpoint, headers=headers)
        
        if response.status_code == 200:
            data = response.json()  # Parse the JSON content
            amount = data['credit_balance']
            global CREDIT_AMOUNT
            CREDIT_AMOUNT = amount - 1

            self.isTokenValid = True
        else:
            self.isTokenValid = False
        
# Wizard Page 3: Select File
class WizardPage3:
    def __init__(self, master, button_command_next, button_command_previous):
        # Initialize the page
        self.master = master
        self.frame = customtkinter.CTkFrame(self.master)
        self.frame.grid(row=0, column=0, sticky="nsew")
        configureColumnAndRow(self.master, 1, 1)
        configureColumnAndRow(self.frame, 2, 3)

        #create title
        self.title = customtkinter.CTkLabel(self.frame, text="Elegir el archivo excel", anchor="center", font=('', 30))
        self.title.grid(row=0, column=0, columnspan=2, padx=10, pady=20) 

        #create browse button to select excel
        self.button_select_file = customtkinter.CTkButton(master=self.frame, text="Browse", command=self.select_file)
        self.button_select_file.grid(row=1, column=0, columnspan=2, padx=10, pady=20) 
        
        #create previous button
        self.button_previous = customtkinter.CTkButton(master=self.frame, text="Atras", command=button_command_previous)
        self.button_previous.grid(row=2, column=0, padx=20)  
        #create next button
        self.button_next = customtkinter.CTkButton(master=self.frame, text="Siguiente", command=button_command_next, state='disabled')
        self.button_next.grid(row=2, column=1, padx=20) 

    #opens file explorer so the user can choose an excel file, then stores the selected path
    def select_file(self):
        filetypes = (('Excel files', '*.xlsx'), ('All files', '*.*'))
        path = filedialog.askopenfilename(initialdir="/", title="Select file", filetypes=filetypes)
        self.button_select_file.configure(fg_color='#0C2D48')
        global SELECTED_FILE_PATH
        SELECTED_FILE_PATH = path
        #enable next button
        self.button_next.configure(state="normal")
        #call the function to obtain the necessary credits in a separate thread
        threading.Thread(target=amountOfNecessaryCredits, daemon=True).start()

# Wizard Page 4: Select Folder
class WizardPage4:
    def __init__(self, master, button_command_next, button_command_previous):
        # Initialize the page
        self.master = master
        self.frame = customtkinter.CTkFrame(self.master)
        self.frame.grid(row=0, column=0, sticky="nsew")
        configureColumnAndRow(self.master, 1, 1)
        configureColumnAndRow(self.frame, 2, 3)

        #create title
        self.title = customtkinter.CTkLabel(self.frame, text="Elegir la carpeta", anchor="center", font=('', 30))
        self.title.grid(row=0, column=0, columnspan=2, padx=10, pady=20) 

        #create browse button to select folder
        self.button_select_folder = customtkinter.CTkButton(master=self.frame, text="Browse", command=self.select_directory)
        self.button_select_folder.grid(row=1, column=0, columnspan=2, padx=10, pady=20) 
        
        #create previous button
        self.button_previous = customtkinter.CTkButton(master=self.frame, text="Atras", command=button_command_previous)
        self.button_previous.grid(row=2, column=0, padx=20)  
        #create next button
        self.button_next = customtkinter.CTkButton(master=self.frame, text="Siguiente", command=button_command_next, state='disabled')
        self.button_next.grid(row=2, column=1, padx=20)  

    #opens file explorer so the user can choose a folder, then stores the selected path
    def select_directory(self):
        path = filedialog.askdirectory(initialdir="/", title="Select a directory")
        self.button_select_folder.configure(fg_color='#0C2D48') 

        global SELECTED_FOLDER_PATH
        SELECTED_FOLDER_PATH = path
        #enable next button
        self.button_next.configure(state="normal")

# Wizard Page 5: Select Format
class WizardPage5:
    def __init__(self, master, button_command_next, button_command_previous):
        # Initialize the page
        self.master = master
        self.frame = customtkinter.CTkFrame(self.master)
        self.frame.grid(row=0, column=0, sticky="nsew")
        configureColumnAndRow(self.master, 1, 1)
        configureColumnAndRow(self.frame, 5, 4)

        #create title
        self.title = customtkinter.CTkLabel(self.frame, text="Elección del Formato", anchor="center", font=('', 30))
        self.title.grid(row=0, column=0,columnspan=4)  

        #create radio buttons to select a format
        self.radio_var = tk.IntVar(value=1)
       
        # Radiobutton 1
        self.radiobutton_1 = customtkinter.CTkRadioButton(self.frame, text="Predeterminado", command=self.radiobutton_event, variable=self.radio_var, value=1)
        self.radiobutton_1.grid(row=1, column=0,sticky="w",padx=30)  

        # Radiobutton 2
        self.radiobutton_2 = customtkinter.CTkRadioButton(self.frame, text="IAE", command=self.radiobutton_event, variable=self.radio_var, value=2)
        self.radiobutton_2.grid(row=1, column=1,sticky="w",padx=30) 

        # Radiobutton 3
        self.radiobutton_3 = customtkinter.CTkRadioButton(self.frame, text="Personalizado", command=self.radiobutton_event, variable=self.radio_var, value=3)
        self.radiobutton_3.grid(row=1, column=2,sticky="w",padx=20) 

        self.createColumns()

        #create personal email checkbox
        self.checkbox_var = customtkinter.BooleanVar(value=MAIL_CHECKBOX)
        self.checkbox = customtkinter.CTkCheckBox(self.frame, variable=self.checkbox_var, text="Obtener mails personales (requiere 1 credito extra por cada mail personal)", command=self.mail_checkbox_clicked)
        self.checkbox.grid(row=3, column=0,columnspan=4,sticky="w",padx=30)
        
        #create previous button
        self.button_previous = customtkinter.CTkButton(master=self.frame, text="Atras", command=button_command_previous)
        self.button_previous.grid(row=4, column=0,pady=40,padx=1,sticky="NE") 
        #create next button
        self.button_next = customtkinter.CTkButton(master=self.frame, text="Siguiente", command=button_command_next)
        self.button_next.grid(row=4, column=2,pady=40,padx=10,sticky="NW") 
    
    #manages click of a radiobutton
    def radiobutton_event(self):
        newFormat = self.radio_var.get()
        if newFormat == 3:
            self.showColumns()
            self.showMailCheckbox()
            if not any(PERSONALIZED_COLUMNS_IS_CHECKED):
                self.button_next.configure(state="disabled")
        elif newFormat == 2:
            self.hideMailCheckbox()
            self.hideColumns()
            self.button_next.configure(state="normal")
        elif newFormat == 1:
            self.showMailCheckbox()
            self.hideColumns()
            self.button_next.configure(state="normal")
        
        global SELECTED_FORMAT
        SELECTED_FORMAT = newFormat
        
    def showMailCheckbox(self):
        self.checkbox.grid(row=3, column=0,columnspan=4,sticky="w",padx=30)
    
    def hideMailCheckbox(self):
        self.checkbox.grid_forget()

    def showColumns(self):
        self.columns_frame.grid(row=2, column=0,columnspan=4)

    #creates the field selector widget
    def createColumns(self):
        #creates a scrollable frame
        self.columns_frame = customtkinter.CTkScrollableFrame(self.frame,orientation='horizontal', width=600, height=100, border_width=3)

        #adds all the fields
        for col, title in enumerate(PERSONALIZED_COLUMNS):
            square = customtkinter.CTkFrame(self.columns_frame, border_width=2, width=175, height=80)
            square.grid_propagate(0)
            square.grid(row=0, column=col, padx=5, pady=5)
            configureColumnAndRow(square, 1, 2)

            title_label = customtkinter.CTkLabel(square, text=title)
            title_label.grid(row=0, column=0, pady=5)

            checkbox_var = customtkinter.BooleanVar(value=PERSONALIZED_COLUMNS_IS_CHECKED[col])
            checkbox = customtkinter.CTkCheckBox(square, variable=checkbox_var, text="", command=lambda c=col: self.checkbox_clicked(c))
            checkbox.grid(row=1, column=0,pady=5, padx=75)

    def hideColumns(self):
        self.columns_frame.grid_forget()
    
    #handles when a checkbox is clicked
    def checkbox_clicked(self, col):
        global PERSONALIZED_COLUMNS_IS_CHECKED
        PERSONALIZED_COLUMNS_IS_CHECKED[col] = not PERSONALIZED_COLUMNS_IS_CHECKED[col]
        if not any(PERSONALIZED_COLUMNS_IS_CHECKED):
            #if no column is checked, disables the next button
            self.button_next.configure(state="disabled")
        else:
            #enables next button because there are checked columns
            self.button_next.configure(state="normal")
    
    #handles when the personal email checkbox is clicked
    def mail_checkbox_clicked(self):
        global MAIL_CHECKBOX
        MAIL_CHECKBOX = not MAIL_CHECKBOX
        
# Wizard Page 6: Validation page
class WizardPage6:
    def __init__(self, master, button_command_next, button_command_previous):
        # Initialize the page
        self.master = master
        self.frame = customtkinter.CTkFrame(self.master)
        self.frame.grid(row=0, column=0, sticky="nsew")
        self.button_command_next = button_command_next
        self.button_command_previous = button_command_previous

        #binds so the panelShown function is entered when the page becomes visible
        self.frame.bind("<Visibility>", self.panelShown)
        
    #shows the page
    def panelShown(self, event):
        configureColumnAndRow(self.master, 1, 1)
        configureColumnAndRow(self.frame, 2, 4)

        #create title
        self.title = customtkinter.CTkLabel(self.frame, text="Verificación", anchor="center", font=('', 30))
        self.title.grid(row=0, column=0, columnspan=2, padx=10, pady=20) 

        #create credit-balance label
        self.credit_amount_text = customtkinter.CTkLabel(self.frame, text="Creditos restantes: " + str(CREDIT_AMOUNT),font=('',20))
        self.credit_amount_text.grid(row=1, column=0, padx=10, pady=20)

        #create necesary-credits label
        self.credit_necesary_text = customtkinter.CTkLabel(self.frame, text="Creditos necesarios: " + str(CREDIT_NECESSARY),font=('',20))
        self.credit_necesary_text.grid(row=2, column=0, padx=10, pady=20)

        #create previous button
        self.button_previous = customtkinter.CTkButton(master=self.frame, text="Atras", command=self.button_command_previous)
        self.button_previous.grid(row=3, column=0, padx=20) 
        #create next button
        self.button_next = customtkinter.CTkButton(master=self.frame, text="Siguiente", command=self.button_command_next, state='disabled' if CREDIT_AMOUNT < CREDIT_NECESSARY else 'normal')
        self.button_next.grid(row=3, column=1, padx=20)

# Wizard Page 7: Data extraction and show results        
class WizardPage7:
    def __init__(self, master, button_finish, on_API_error, on_open_file_error):
        # Initialize the page
        self.master = master
        self.frame = customtkinter.CTkFrame(self.master)
        self.frame.grid(row=0, column=0, sticky="nsew")
        #initialize variables
        self.button_finish = button_finish
        self.on_API_error = on_API_error
        self.on_open_file_error = on_open_file_error

        #binds so the panelShown function is entered when the page becomes visible
        self.frame.bind("<Visibility>", self.panelShown)

    def panelShown(self, event):
        configureColumnAndRow(self.master, 1, 1)
        configureColumnAndRow(self.frame, 2, 4)

        #create title
        self.title = customtkinter.CTkLabel(self.frame, text="Consiguiendo datos ...", anchor="center",font=('',30))
        self.title.grid(row=0, column=0, padx=10, pady=20, columnspan=2)
        
        #create progressbar
        self.progressbar = customtkinter.CTkProgressBar(self.frame, orientation="horizontal", width=400,mode='determinate', determinate_speed=PERCENTAGE_PER_PROFILE/2)
        self.progressbar.set(0)
        self.progressbar.grid(row=1, column=0, padx=(30, 50), pady=20, sticky='e')
       
        #create percentage label
        self.value_label = customtkinter.CTkLabel(self.frame, text='0%')
        self.value_label.grid(row=1, column=1, padx=(0, 0), sticky='w')

        #starts the data extraction in a separate thread
        self.worker_thread = threading.Thread(target=self.getLinkedInDataAndShowResult, daemon=True)
        self.worker_thread.start()

    #updates the percentage label
    def progress(self, val):
        if val < 100:
            self.value_label.configure(text = str(int(val * 100) ) + '%')

    #stops the progressbar and updates the UI
    def finishOperation(self):
        #stop progressbar
        self.progressbar.set(1)
        self.progressbar.stop()
        self.value_label.configure(text = '100%')

        #update the UI
        self.showResult()   

    #this function calls a data extraction function depending on the format chosen. It also starts the progressbar
    #in case of an exeption, it alarms the user
    def getLinkedInDataAndShowResult(self):
        try:
            if SELECTED_FORMAT == 1:
                self.getLinkedInDataDefault()
            elif SELECTED_FORMAT == 2:
                self.getLinkedInDataIAE()
            elif SELECTED_FORMAT == 3:
                self.getLinkedInDataPersonalized()
        except Exception as e:
            self.on_API_error()
        
    #Updates the progressBar 
    def updateProgressBar(self):
        global PASSED
        val = self.progressbar.get()
        newVal = val
        if PASSED:
            #stop progressbar
            self.progressbar.stop()
        else: 
            #sets new progressbar value
            self.progressbar.step()
            newVal = self.progressbar.get()
            if newVal < val:
                #Avoid restart of progressbar bug
                PASSED = True
                newVal = 0.99
                self.progressbar.set(0.99)
        #updates percentage label
        self.progress(newVal)

    #show final results once the data fetching is finished
    def showResult(self):
        #create title
        self.title = customtkinter.CTkLabel(self.frame, text="Los datos se han guardado con éxito!", anchor="center", font=('', 30))
        self.title.grid(row=0, column=0, padx=10, pady=20,columnspan=2)

        #create button to view results
        self.open_file_button = customtkinter.CTkButton(master=self.frame, text="Ver resultados", command=self.openFiles)
        self.open_file_button.grid(row=2, column=0, columnspan=2)

        #create button to close program
        self.finish_button = customtkinter.CTkButton(master=self.frame, text="Cerrar programa", command=self.button_finish, width=180)
        self.finish_button.place(relx=0.5, rely=0.5, anchor=customtkinter.CENTER)
        self.finish_button.grid(row=3, column=0, columnspan=2)  
    
    #open result files
    #in case of error it notifies the user
    def openFiles(self):
        if SELECTED_FORMAT == 2:
            academic_path = SELECTED_FOLDER_PATH + "/" + ACADEMIC_RESULT_FILE_NAME + ".xlsx"
            experience_path = SELECTED_FOLDER_PATH + "/" + EXPERIENCE_RESULT_FILE_NAME + ".xlsx"
            if os.path.isfile(academic_path) and os.path.isfile(experience_path):
                os.system("start EXCEL.EXE " + academic_path)
                os.system("start EXCEL.EXE " + experience_path)
            else:
                self.on_open_file_error()
        else:
            full_path = SELECTED_FOLDER_PATH + "/" + RESULT_FILE_NAME + ".xlsx"
            if os.path.isfile(full_path):
                os.system("start EXCEL.EXE " + full_path)
            else:
                self.on_open_file_error()
        
    # ##-------------------------- GETTING DATA ---------------------------------------------------------------------------------------------------------

    #Inputs the url to the LinkedIn profile of a person
    #Function does an http requests to Proxycurl API
    #Returns the data returned from the API
    def sendRequest(self, url):
        headers = {'Authorization': 'Bearer ' + API_TOKEN}
        api_endpoint = 'https://nubela.co/proxycurl/api/v2/linkedin'
        params = {
        'linkedin_profile_url': url,
        }
        if MAIL_CHECKBOX and (SELECTED_FORMAT != 2):
            #Adds personal email
            params['personal_email'] = 'include'
        response = requests.get(api_endpoint, params=params, headers=headers)
        
        if response.status_code == 200:
            data = response.json()  # Parse the JSON content
            return data
        else:
            print(f"Request failed with status code {response.status_code}")
            return None 

    #Inputs a map with keys: day, month, year
    #Returns a string formatted dd/mm/yyyy
    def format_date(self, data):
        day = str(data["day"]).zfill(2)
        month = str(data["month"]).zfill(2)
        year = str(data["year"])

        return f"{day}/{month}/{year}"
    
    # ## --------------------------  IAE  ---------------------------------------------------------------------------------------------------------------

    #This function extracts all the data and saves it in an excel in the IAE format
    def getLinkedInDataIAE(self):
        #reads the urls and ids from the excel
        df = pd.read_excel(SELECTED_FILE_PATH)
        first_column_data = df.iloc[0:, 0].tolist()
        second_column_data = df.iloc[0:, 1].tolist()

        #creating academic excel
        academic_path = os.path.relpath(SELECTED_FOLDER_PATH + '\Resultados_Historial_Academico.xlsx')
        academic_workbook = xlsxwriter.Workbook(academic_path)
        academic_worksheet = academic_workbook.add_worksheet()

        #creating experience excel
        experience_path = os.path.relpath(SELECTED_FOLDER_PATH + '\Resultados_Datos_Laborales.xlsx')
        experience_workbook = xlsxwriter.Workbook(experience_path)
        experience_worksheet = experience_workbook.add_worksheet()
        
        #writing the columns in the academic worksheet
        index = 0
        for value in IAE_ACADEMIC_COLUMNS:
            academic_worksheet.write(0,index, str(value))
            index += 1
        
        #writing the columns in the experience worksheet
        index = 0
        for value in IAE_EXPERIENCE_COLUMNS:
            experience_worksheet.write(0,index, str(value))
            index += 1

        #main loop

        current_academic_row = 1
        current_experience_row = 1

        for i in range(len(first_column_data)):
            url = first_column_data[i]
            id = second_column_data[i]

            #feches data from API
            response = self.sendRequest(url)

            #updates the progressBar
            self.updateProgressBar()

            if(response == None): continue

            currentCol = 0

            #Writing academic worksheet

            educations = response['education']

            for education in educations:
                for val in IAE_ACADEMIC_COLUMNS_PROXYCURL_EQUIVALENT:
                    match val:
                        case None:
                            pass
                        case 'ContactID':
                            academic_worksheet.write(current_academic_row,currentCol, str(id))
                        case 'start_year':
                            if (education['starts_at'] != None): academic_worksheet.write(current_academic_row,currentCol, str(education['starts_at']['year']))
                        case 'end_year':
                            if (education['ends_at'] != None): academic_worksheet.write(current_academic_row,currentCol, str(education['ends_at']['year']))
                        case _:
                            if(education[val] == None):
                                pass
                            else:
                                academic_worksheet.write(current_academic_row,currentCol, str(education[val]))
                    currentCol += 1
                current_academic_row+=1
                currentCol = 0

            currentCol = 0

            #Writing experience worksheet
            
            experiences = response['experiences']

            for experience in experiences:
                for val in IAE_EXPERIENCE_COLUMNS_PROXYCURL_EQUIVALENT:
                    match val:
                        case None:
                            pass
                        case 'ContactID':
                            experience_worksheet.write(current_experience_row,currentCol, str(id))
                        case 'start_year':
                            if (experience['starts_at'] != None): experience_worksheet.write(current_experience_row,currentCol, str(experience['starts_at']['year']))
                        case 'end_year':
                            if (experience['ends_at'] != None): experience_worksheet.write(current_experience_row,currentCol, str(experience['ends_at']['year']))
                        case 'actual':
                            experience_worksheet.write(current_experience_row,currentCol, 'Sí' if experience['ends_at'] == None else 'No')
                        case 'end_month':
                            if (experience['ends_at'] != None): experience_worksheet.write(current_experience_row,currentCol, str(experience['ends_at']['month']))
                        case _:
                            data = experience[val]
                            if(data == None):
                                pass
                            elif(type(data) == dict):
                                experience_worksheet.write(current_experience_row,currentCol, self.format_date(data))
                            else:
                                experience_worksheet.write(current_experience_row,currentCol, str(data))
                    currentCol += 1
                current_experience_row+=1
                currentCol = 0
                
        #autofits the text inside cells
        academic_worksheet.autofit()
        experience_worksheet.autofit()

        #closes workbooks
        academic_workbook.close()
        experience_workbook.close()

        #ends progressBar
        self.finishOperation()

    # ## --------------------------  DEFAULT  ---------------------------------------------------------------------------------------------------------------

    #Creates an array of the column names in the default format
    def getDefaultFormat(self):
        format = []
        format += GENERAL_COLUMNS

        #languages
        for i in range(1, ITERACIONES_IDIOMA+1):
            format.append('Idioma_' + str(i))

        #educations
        for i in range(1, ITERACIONES_EDUCACION+1):
            for col in EDUCATION_COLUMNS:
                format.append('Educacion_' + str(i) + '_' + col)
        
        #experiences
        for i in range(1, ITERACIONES_EXPERIENCIAS+1):
            for col in EXPERIENCE_COLUMNS:
                format.append('Experiencia_' + str(i) + '_' + col)
        
        #personal emails
        if MAIL_CHECKBOX:
            for i in range(1, ITERACIONES_MAILS+1):
                format.append('Mail Personal_' + str(i))

        return format
    
    #Writes the data of iterative fields, such as education and experience. For default format
    def iterateColumns(self, map, iterations, currentCol, currentRow, equivalent_array, worksheet):
        length = len(map)
        index=0
        for i in range(iterations):
            if(index < length):
                inner_map = map[index]
                for val in equivalent_array:
                    data = inner_map[val]
                    if(data == None):
                        currentCol += 1
                        continue
                    elif (type(data) == dict):
                        worksheet.write(currentRow,currentCol, self.format_date(data))
                    else:
                        worksheet.write(currentRow,currentCol, str(data))  
                    currentCol += 1
                index+=1
            else:
                currentCol += len(equivalent_array)
        return currentCol

    #This function extracts all the data and saves it in an excel in the default format
    def getLinkedInDataDefault(self):
        #reads the column of the excel file
        df = pd.read_excel(SELECTED_FILE_PATH)
        first_column_data = df.iloc[0:, 0].tolist()

        #creating excel
        path = os.path.relpath(SELECTED_FOLDER_PATH + '\\' + RESULT_FILE_NAME + '.xlsx')
        workbook = xlsxwriter.Workbook(path)
        worksheet = workbook.add_worksheet()

        format = self.getDefaultFormat()

        #writing the column names in the worksheet
        index = 0
        for value in format:
            worksheet.write(0,index, str(value))
            index += 1

        #main loop

        currentRow = 1

        for url in first_column_data:
            #fetches data
            response = self.sendRequest(url)

            #updates the progressBar
            self.updateProgressBar()

            if(response == None): continue

            currentCol = 0

            #Writing url on first column
            worksheet.write(currentRow,currentCol, url)
            currentCol += 1

            #Writing general columns

            for val in GENERAL_COLUMNS_PROXYCURL_EQUIVALENT:
                worksheet.write(currentRow,currentCol, response[val])
                currentCol += 1
                    
            #Writing language columns

            languages = response['languages']
            languages_length = len(languages)
            languages_index=0
            for i in range(ITERACIONES_IDIOMA):
                if(languages_index < languages_length):
                    worksheet.write(currentRow,currentCol, languages[languages_index])
                    languages_index+=1
                currentCol += 1
                
            #Writing education columns
            
            currentCol = self.iterateColumns(response['education'], ITERACIONES_EDUCACION, currentCol, currentRow, EDUCATION_COLUMNS_PROXYCURL_EQUIVALENT, worksheet)

            #Writing experience columns

            currentCol = self.iterateColumns(response['experiences'], ITERACIONES_EXPERIENCIAS, currentCol, currentRow, EXPERIENCE_COLUMNS_PROXYCURL_EQUIVALENT, worksheet)

            #Writing mail columns
            if MAIL_CHECKBOX:
                mails = response['personal_emails']
                mails_length = len(mails)
                mails_index=0
                for i in range(ITERACIONES_MAILS):
                    if(mails_index < mails_length):
                        worksheet.write(currentRow,currentCol, mails[mails_index])
                        mails_index+=1
                    currentCol += 1

            currentRow +=1

        worksheet.autofit()

        workbook.close()

        #ends progressBar
        self.finishOperation()


         # ## --------------------------  PERSONALIZED  ---------------------------------------------------------------------------------------------------------------

# ## --------------------------  PERSONALIZED  ---------------------------------------------------------------------------------------------------------------

    #Creates an array of the column names in the personalized format
    def getPersonalizedFormat(self):
        format = []
        
        #add general columns
        j = 0
        while PERSONALIZED_COLUMNS[j] != 'Idiomas':
            if PERSONALIZED_COLUMNS_IS_CHECKED[j]:
                format.append(PERSONALIZED_COLUMNS[j])
            j += 1

        #add language columns
        if PERSONALIZED_COLUMNS_IS_CHECKED[j]:
            for i in range(1, ITERACIONES_IDIOMA+1):
                format.append('Idioma_' + str(i))
        
        j += 1

        #add education columns
        education_cols = []
        while PERSONALIZED_COLUMNS[j].split('_')[0] == 'Educacion':
            if PERSONALIZED_COLUMNS_IS_CHECKED[j]:
                education_cols.append(PERSONALIZED_COLUMNS[j])
            j+=1
        
        for i in range(1, ITERACIONES_EDUCACION+1):
            for col in education_cols:
                format.append(col + '_' + str(i))
        
        #add experience columns
        experience_cols = []
        while j < len(PERSONALIZED_COLUMNS) and PERSONALIZED_COLUMNS[j].split('_')[0] == 'Experiencia':
            if PERSONALIZED_COLUMNS_IS_CHECKED[j]:
                experience_cols.append(PERSONALIZED_COLUMNS[j])
            j+=1
            
        for i in range(1, ITERACIONES_EXPERIENCIAS+1):
            for col in experience_cols:
                format.append(col + '_' + str(i))
        
        #add personal email columns
        if MAIL_CHECKBOX:
            for i in range(1, ITERACIONES_MAILS+1):
                format.append('Mail Personal_' + str(i))

        return format

    #Writes the data of iterative fields, such as education and experience. For personalized format
    def iterateColumnsPersonalized(self, map, split):
        index = PERSONALIZED_COLUMNS.index(split[0] + '_' + split[1])
        n = int(split[2]) - 1
        if n < len(map):
            innerMap = map[n]
            data = innerMap[PERSONALIZED_COLUMNS_PROXYCURL_EQUIVALENT[index]]
            if(data == None):
                return ''
            elif (type(data) == dict):
               return self.format_date(data)
            else:
                return str(data)
            
    #This function extracts all the data and saves it in an excel in the personalized format
    def getLinkedInDataPersonalized(self):
            #reads the column of the excel file
            df = pd.read_excel(SELECTED_FILE_PATH)
            first_column_data = df.iloc[0:, 0].tolist()

            #creating excel
            path = os.path.relpath(SELECTED_FOLDER_PATH + '\\' + RESULT_FILE_NAME + '.xlsx')
            workbook = xlsxwriter.Workbook(path)
            worksheet = workbook.add_worksheet()

            format = self.getPersonalizedFormat()
            
            #writing the column names in the worksheet
            index = 0
            for value in format:
                worksheet.write(0,index, str(value))
                index += 1
            
            #main loop

            currentRow = 1

            for url in first_column_data:
                #fetches data
                response = self.sendRequest(url)

                #updates the progressBar
                self.updateProgressBar()

                if(response == None): continue

                currentCol = 0

                for val in format:
                    toWrite = ''
                    split = val.split('_')
                    if val == 'Link':
                        toWrite = url
                    elif len(split) > 1:
                        if split[0] == 'Idioma':
                            languages = response['languages']
                            n = int(split[1]) - 1
                            if n < len(languages):
                                toWrite = languages[n]
                        elif split[0] == 'Educacion':
                            toWrite = self.iterateColumnsPersonalized(response['education'], split)
                        elif split[0] == 'Experiencia':
                            toWrite = self.iterateColumnsPersonalized(response['experiences'], split)
                        elif split[0] == 'Mail Personal':
                            mails = response['personal_emails']
                            n = int(split[1]) - 1
                            if n < len(mails):
                                toWrite = mails[n]
                    else:
                        index = PERSONALIZED_COLUMNS.index(val)
                        toWrite = response[PERSONALIZED_COLUMNS_PROXYCURL_EQUIVALENT[index]]
                    worksheet.write(currentRow,currentCol, toWrite)
                    currentCol += 1

                currentRow += 1

            worksheet.autofit()
            
            workbook.close()

            #ends progressBar
            self.finishOperation()

# Wizard Error Page: Shows error description and solution
class WizardErrorPage:
    def __init__(self, master, errorType, button_command):
        # Initialize the page
        self.master = master
        self.frame = customtkinter.CTkFrame(self.master)
        self.frame.grid(row=0, column=0, sticky="nsew")

        configureColumnAndRow(self.master, 1, 1)
        configureColumnAndRow(self.frame, 1, 4)
        
        #define description and button text depending in error type
        if errorType == 1:
            #Invalid excel format error
            self.description_text = "Los datos en el archivo excel con los links no estan en el formato indicado. Ante cualquier duda acerca del formato, revise el manual de usuario"
            self.button_text = "Volver"
        elif errorType == 2:
            #Error while fetching data 
            self.description_text = "No se pudo extraer los datos debido a un error con la API de Proxycurl. Revise el manual de usuario o vuelva a intentar mas tarde"
            self.button_text = "Cerrar programa"
        elif errorType == 3:
            #Invalid folder path error 
            self.description_text = "La carpeta seleccionada es invalida. No tenes los permisos para editar esta carpeta, por favor seleccion otra carpeta"
            self.button_text = "Volver"
        elif errorType == 4:
            #Error while opening file
            self.description_text = "No se pudo abrir el archivo debido a que este nunca fue creado. Revise el manual de usuario o vuelva a intentarlo de nuevo seleccionando otra carpeta"
            self.button_text = "Cerrar programa"

        #create warning image
        button_image = customtkinter.CTkImage(Image.open("resources/aviso.png"), size=(100, 100))
        self.image_button = customtkinter.CTkLabel(self.frame, text="",image=button_image)
        self.image_button.grid(row=0, column=0)

        #create title
        self.title = customtkinter.CTkLabel(self.frame, text="Ha ocurrido un error", anchor="center",font=('',35))
        self.title.grid(row=1, column=0, padx=10, pady=20)

        #create body
        self.description = customtkinter.CTkLabel(self.frame, text=self.description_text,font=('',18),wraplength=600)
        self.description.grid(row=2, column=0, padx=10, pady=20)

        #create button
        self.button = customtkinter.CTkButton(master=self.frame, text=self.button_text, command=button_command)
        self.button.grid(row=3, column=0)

#Wizard class: Connects all the wizard pages
class MyWizard:
    def __init__(self, root):
        self.root = root
        self.root.title("LinkedIn Data Wizard")
        #ready gif
        threading.Thread(target=self.ready_gif, daemon=True).start()

        self.pages = []

        self.page1 = WizardPage1(self.root, self.show_page2)
        self.pages.append(self.page1)

        self.page2 = WizardPage2(self.root,self.show_page3, self.show_page1)
        self.page2.frame.grid_forget()
        self.pages.append(self.page2)

        self.page3 = WizardPage3(self.root,self.show_page4, self.show_page2)
        self.page3.frame.grid_forget()
        self.pages.append(self.page3)

        self.page4 = WizardPage4(self.root,self.show_page5,self.show_page3)
        self.page4.frame.grid_forget()
        self.pages.append(self.page4)

        self.page5 = WizardPage5(self.root,self.show_page6,self.show_page4)
        self.page5.frame.grid_forget()
        self.pages.append(self.page5)

        self.page6 = WizardPage6(self.root,self.show_page7,self.show_page5)
        self.page6.frame.grid_forget()
        self.pages.append(self.page6)

        self.page7 = WizardPage7(self.root,self.close_wizard, self.show_errorWhenFetchingDataPage, self.show_errorWhenOpeningFilePage)
        self.page7.frame.grid_forget()
        self.pages.append(self.page7)

        #ERROR PAGES

        #Invalid excel format error page
        self.invalidExcelFormatErrorPage = WizardErrorPage(self.root, 1, self.show_page3)
        self.invalidExcelFormatErrorPage.frame.grid_forget()
        self.pages.append(self.invalidExcelFormatErrorPage)

        #Error while fetching data page
        self.ErrorWhenFetchingDataPage = WizardErrorPage(self.root, 2, self.close_wizard)
        self.ErrorWhenFetchingDataPage.frame.grid_forget()
        self.pages.append(self.ErrorWhenFetchingDataPage)

        #Invalid folder path error page
        self.InvalidFolderErrorPage = WizardErrorPage(self.root, 3, self.show_page4)
        self.InvalidFolderErrorPage.frame.grid_forget()
        self.pages.append(self.InvalidFolderErrorPage)

        #Error while opening file page
        self.ErrorWhenOpeningFilePage = WizardErrorPage(self.root, 4, self.close_wizard)
        self.ErrorWhenOpeningFilePage.frame.grid_forget()
        self.pages.append(self.ErrorWhenOpeningFilePage)

        self.current_page = 0
        self.show_current_page()

    def show_page1(self):
        self.pages[self.current_page].frame.grid_forget()
        self.page1.frame.grid(row=0, column=0, sticky="nsew")
        self.current_page = 0

    def show_page2(self):
        self.pages[self.current_page].frame.grid_forget()
        self.page2.frame.grid(row=0, column=0, sticky="nsew")
        self.current_page = 1

    def show_page3(self):
        self.pages[self.current_page].frame.grid_forget()
        self.page3.frame.grid(row=0, column=0, sticky="nsew")
        self.current_page = 2

    def show_page4(self):
        self.pages[self.current_page].frame.grid_forget()
        self.page4.frame.grid(row=0, column=0, sticky="nsew")
        self.current_page = 3

    def show_page5(self):
        if self.isFolderPathValid():
            self.pages[self.current_page].frame.grid_forget()
            self.page5.frame.grid(row=0, column=0, sticky="nsew")
            self.current_page = 4
        else:
            self.show_invalidFolderErrorPage()

    def show_page6(self):
        self.pages[self.current_page].frame.grid_forget()
        self.page6.frame.grid(row=0, column=0, sticky="nsew")
        self.current_page = 5
    
    def show_page7(self):
        if self.isFormatValid():
            self.pages[self.current_page].frame.grid_forget()
            self.page7.frame.grid(row=0, column=0, sticky="nsew")
            self.current_page = 6
        else:
            self.show_invalidExcelFormatErrorPage()
    
    def show_invalidExcelFormatErrorPage(self):
        self.pages[self.current_page].frame.grid_forget()
        self.invalidExcelFormatErrorPage.frame.grid(row=0, column=0, sticky="nsew")
        self.current_page = 7
    
    def show_errorWhenFetchingDataPage(self):
        self.pages[self.current_page].frame.grid_forget()
        self.ErrorWhenFetchingDataPage.frame.grid(row=0, column=0, sticky="nsew")
        self.current_page = 8
    
    def show_invalidFolderErrorPage(self):
        self.pages[self.current_page].frame.grid_forget()
        self.InvalidFolderErrorPage.frame.grid(row=0, column=0, sticky="nsew")
        self.current_page = 9
    
    def show_errorWhenOpeningFilePage(self):
        self.pages[self.current_page].frame.grid_forget()
        self.ErrorWhenOpeningFilePage.frame.grid(row=0, column=0, sticky="nsew")
        self.current_page = 10

    def close_wizard(self):
        self.root.destroy()

    def show_current_page(self):
        self.pages[self.current_page].frame.grid(row=0, column=0, sticky="nsew")
    
    #checks if the folder path is valid
    def isFolderPathValid(self):
        permissions = os.stat(SELECTED_FOLDER_PATH).st_mode
        if (permissions & 0o400) and (permissions & 0o200) and (permissions & 0o100):
            return True
        return False
        
    #Checks if the provided excel format is valid
    def isFormatValid(self):
        #reads the urls and ids from the excel
        df = pd.read_excel(SELECTED_FILE_PATH)
        first_column_data = df.iloc[0:, 0].tolist()
        second_column_data = []
        if len(df.columns) > 1:
            second_column_data = df.iloc[0:, 1].tolist()
        if SELECTED_FORMAT == 1 or SELECTED_FORMAT == 3:
            if len(first_column_data) == 0: 
                return False
        elif SELECTED_FORMAT == 2:
            if len(first_column_data) == 0 or (len(first_column_data) != len(second_column_data)):
                return False
        return True
    
    #stores gif in GIF_FRAMES array
    def ready_gif(self):
        global FRAME_DELAY, GIF_FRAMES
        gif_file = Image.open("resources/loading-gif.gif")

        for r in range(0, gif_file.n_frames):
            gif_file.seek(r)
            GIF_FRAMES.append(gif_file.copy())
        FRAME_DELAY = gif_file.info['duration']

#Creates wizard and runs program
def main():
    customtkinter.set_appearance_mode("dark")  # Modes: system (default), light, dark
    customtkinter.set_default_color_theme("dark-blue")  # Themes: blue (default), dark-blue, green
    root = customtkinter.CTk()
    MyWizard(root)
    

    width = 700 # Width 
    height = 500 # Height

    # set minimum window size value
    root.minsize(width, height)
    
    # set maximum window size value
    root.maxsize(width, height)
    
    centerWindow(root, width, height)
    root.mainloop()

#Centers window
def centerWindow(root, width, height):
    screen_width = root.winfo_screenwidth()  # Width of the screen
    screen_height = root.winfo_screenheight() # Height of the screen
    
    # Calculate Starting X and Y coordinates for Window
    x = (screen_width/2) - (width/2)
    y = (screen_height/2) - (height/2)
    
    root.geometry('%dx%d+%d+%d' % (width, height, x, y))

if __name__ == "__main__":
    main()
