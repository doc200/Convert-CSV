from tkinter import *
import tkinter as tk
from PIL import Image, ImageTk
from tkinter import filedialog
import pandas as pd
import os
import sys

def resource_path(relative_path):
      """ Get absolute path to resource, works for dev and for PyInstaller """
      base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
      return os.path.join(base_path, relative_path)
    
img_path_icon = resource_path("img/icon.ico")
img_path_image = resource_path("img/image.jpg")
  
class App:

    def __init__(self, window):
        self.test1 = False
        self.test2 = False
        #setting title
        window.title("Tools convert CSV to Excel")
        window.iconbitmap(img_path_icon)
        #setting window size
        window.geometry("700x700")#"1280x700"
        window.columnconfigure(0, weight=1)
        window.columnconfigure(1, weight=1)
        window.columnconfigure(2, weight=1)
        window.columnconfigure(3, weight=1)
        window.columnconfigure(4, weight=1)
        window.columnconfigure(5, weight=1)
        window.columnconfigure(6, weight=1)
        

        window.rowconfigure(0, weight=1)
        window.rowconfigure(1, weight=1)
        window.rowconfigure(2, weight=1)
        window.rowconfigure(3, weight=1)
        window.rowconfigure(4, weight=1)
        window.rowconfigure(5, weight=1)
        window.rowconfigure(6, weight=1)
        
        
        window.config(bg="#38374D")
        window.resizable(width=False, height=False)
        # Image
        fond = Image.open(img_path_image)
        photo = ImageTk.PhotoImage(fond)
        label = Label(window, image=photo)
        label.image = photo
        label.grid(row=6, column=0,  columnspan=6, sticky="s")
        # ajouter un second text
        version = Label(window, text="VER 2.01", font=("Courrier", 12), bg="#38374D", fg="white")
        version.grid(row=6, column=0,  sticky="ws", padx=1)
        # Ajouter un premier text
        label_title = Label(window, text="Convertisseur de CSV en XLSX (EXCEL)", font=("Courrier", 20), bg="#38374D",
                            fg="white")
        label_title.grid(row=0, column=0, columnspan=6, sticky="n", padx=15)

        # Source bouton*********************************************************************************************************************************
        bouton_getdir = Button(window, text="Sélectionner des fichiers .csv", font=("Courrier", 13), command=(self.GetDir))
        bouton_getdir.grid(row=1, column=0, sticky="W", padx=5)

        # Target bouton
        bouton_targetdir = Button(window, text="Sélectionner un dossier pour enregistrer", font=("Courrier", 13), command=(self.TargetDir))
        bouton_targetdir.grid(row=2, column=0, sticky="W", padx=5)

        # Bouton Start
        self.bouton_start = Button(window, text="Start", font=("Courrier", 15), repeatdelay=150, activeforeground="blue", command=(self.Controle))
        self.bouton_start.grid(row=4, column=5, sticky="W")

        # bouton de sortie
        bouton = Button(window, text="Fermer", font=("Courrier", 15), activeforeground="red", command=window.quit)
        bouton.grid(row=4, column=4, sticky="W", padx=5)
            
    
    def GetDir(self):  # sourcery skip: assign-if-exp, boolean-if-exp-identity
          self.dirname = window.filename =  filedialog.askopenfilenames(initialdir = "/",title='Veuillez selectioner des fichiers .csv',
                                filetypes = (("csv", ("*.csv")), ("all files","*.*")))
          self.liste_images = list(self.dirname)
          self.nb_lst = len(self.liste_images)
          if not self.liste_images:
            self.test1= False
          else:
            self.test1 = True
          label_subtitle = Label(
              window,
              text=f"{self.nb_lst}  fichier(s)",
              font=("Courrier", 10),
              bg="#38374D",
              fg="white")
          label_subtitle.grid(row=1, column=2, columnspan=2,  sticky="W") 

    
    def TargetDir(self):
          # Chemin dossier Target
          self.tarname = filedialog.askdirectory()
          label_tar = Label(window, text=self.tarname, font=("Courrier", 10), bg="#38374D", fg="white")
          label_tar.grid(row=2, column=2, columnspan=5, sticky="W", padx=1)
          self.test2 = bool(self.tarname)

    def Controle(self):

      if self.test1 == False:
        controle1 = Label(window, text=('Selectioner des fichiers .csv'), font=("Courrier", 10), bg="#38374D", fg="white")
        controle1.grid(row=1, column=2, columnspan=3,  sticky="W")
      
        
      if self.test2 == False:
        controle2 = Label(window, text=('Selectioner un dossier'), font=("Courrier", 10), bg="#38374D", fg="white")
        controle2.grid(row=2, column=2, columnspan=3,  sticky="W")

      if (self.test1 == True and self.test2 == True):
        self.Start()


    def Start(self):
        i = self.nb_lst

        for i, _ in enumerate(self.liste_images):
          self.repertoire_complet = self.liste_images[i]
          self.Nom_image_complet = (self.liste_images[i].split('/')[-1])
          self.extension = (self.liste_images[i].split('.')[-1])
          self.Nom_sans_extension = (self.Nom_image_complet.split(".")[-2])

          read_file = pd.read_csv (self.repertoire_complet)
          read_file.to_excel (self.tarname + '/' + self.Nom_sans_extension + '.xlsx', index = False, header=True, encoding='utf8')

        fin = Label(window, text="Convertion terminé", font=("Courrier", 19), bg="#38374D", fg="white")
        fin.grid(row=4, column=0, columnspan=2, sticky="w", padx=20)



if __name__ == "__main__":
    window = tk.Tk()
    app = App(window)
    window.mainloop()