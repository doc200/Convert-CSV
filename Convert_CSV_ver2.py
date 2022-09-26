from tkinter import *
import tkinter as tk
import os
import sys
from tkinter import filedialog
import pandas as pd

def resource_path(relative_path):
      """ Get absolute path to resource, works for dev and for PyInstaller """
      base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
      return os.path.join(base_path, relative_path)
    
img_path_icon = resource_path("img/icon.ico")
img_path_background = resource_path("img/background.png")
img_path_img0 = resource_path("img/img0.png")
img_path_img1 = resource_path("img/img1.png")
img_path_img2 = resource_path("img/img2.png")
img_path_img3 = resource_path("img/img3.png")

class App:

    def __init__(self, window):
        self.test1 = False
        self.test2 = False
        window.geometry("1000x700")
        window.resizable(width=False, height=False)
        window.configure(bg = "#ffffff")
        window.title("Tools convert CSV to Excel")
        window.iconbitmap(img_path_icon)
        self.canvas = Canvas(
            window,
            bg = "#ffffff",
            height = 700,
            width = 1000,
            bd = 0,
            highlightthickness = 0,
            relief = "ridge")
        self.canvas.place(x = 0, y = 0)

        self.background_img = PhotoImage(file = img_path_background)
        self.background = self.canvas.create_image(500.0, 350.0,image=self.background_img)

        #***********VER PROGRAMME ***********************************
        self.canvas.create_text(
            73.5, 657.5,
            text = "VER 2.03",
            fill = "#ffffff",
            font = ("Inter-Medium", int(16.0)))

        #***********Bouton START ***********************************************
        self.img0 = PhotoImage(file = img_path_img0)
        self.bouton_start = Button(
            image = self.img0,
            borderwidth = 0,
            highlightthickness = 0,
            command = self.Controle,
            relief = "flat")

        self.bouton_start.place(
            x = 768, y = 383,
            width = 166,
            height = 79)

        #***********Bouton FERMER ***********************************************
        self.img1 = PhotoImage(file = img_path_img1)
        self.bouton_fermer = Button(
            image = self.img1,
            borderwidth = 0,
            highlightthickness = 0,
            command = window.quit,
            relief = "flat")

        self.bouton_fermer.place(
            x = 602, y = 386,
            width = 171,
            height = 76)

        #***********Bouton SELECTIONNER DOSSIER ***********************************
        self.img2 = PhotoImage(file = img_path_img2)
        self.bouton_getdir = Button(
            image = self.img2,
            borderwidth = 0,
            highlightthickness = 0,
            command = self.TargetDir,
            relief = "flat")

        self.bouton_getdir.place(
            x = 171, y = 190,
            width = 434,
            height = 89)

        #***********Bouton SELECTIONNER FICHIER CSV ***********************************
        self.img3 = PhotoImage(file = img_path_img3)
        self.bouton_targetdir = Button(
            image = self.img3,
            borderwidth = 0,
            highlightthickness = 0,
            command = self.GetDir,
            relief = "flat")

        self.bouton_targetdir.place(
            x = 173, y = 83,
            width = 336,
            height = 85)

    def GetDir(self):
          # Chemin dossier Source
          self.dirname = window.filename =  filedialog.askopenfilenames(initialdir = "/",title='Veuillez selectioner des fichiers .csv',
                                filetypes = (("csv", ("*.csv")), ("all files","*.*")))
          self.liste_images = list(self.dirname)
          self.nb_lst = len(self.liste_images)
          if not self.liste_images:
                self.test1= False
          else:
                self.test1 = True

                self.Dir = self.canvas.create_text(
                    710,
                    126,
                    text=f"{self.nb_lst}  fichier(s)",
                    fill="#38374D",
                    font=("Inter-Medium", int(16.0)))

    def TargetDir(self):
          # Chemin dossier Target
          self.tarname = filedialog.askdirectory()
          self.Tar = self.canvas.create_text(
                730, 231,
                text = self.tarname,
                fill = "#38374D",
                font = ("Inter-Medium", int(12.0)))
          self.test2 = bool(self.tarname)

    def Controle(self):

          if self.test1 == False:
            self.text1= Label(self.canvas, text="Selectioner des fichiers .csv", foreground='white',  font= ("Inter-Medium", int(17.0)), bg='red')
            self.text1.place(x=660, y=120)
            self.text1.after(3500, self.text1.destroy)    

          if self.test2 == False:
            self.text2= Label(self.canvas, text="Selectioner un dossier", foreground='white',  font= ("Inter-Medium", int(17.0)), bg='red')
            self.text2.place(x=660, y=220)
            self.text2.after(3500, self.text2.destroy)    

          if (self.test1 == True and self.test2 == True):
            self.Start()

          self.test2 = bool(self.tarname)


    def Start(self):
        i = self.nb_lst

        for i, _ in enumerate(self.liste_images):
          self.repertoire_complet = self.liste_images[i]
          self.Nom_image_complet = (self.liste_images[i].split('/')[-1])
          self.extension = (self.liste_images[i].split('.')[-1])
          self.Nom_sans_extension = (self.Nom_image_complet.split(".")[-2])

          read_file = pd.read_csv (self.repertoire_complet)
          read_file.to_excel (self.tarname + '/' + self.Nom_sans_extension + '.xlsx', index = False, header=True, encoding='utf8')

        self.fin = self.canvas.create_text(
            320, 418,
            text = 'Convertion terminé',
            fill = "#38374D",
            font = ("Inter-Medium", int(16.0)))




if __name__ == "__main__":
    window = tk.Tk()
    app = App(window)
    window.mainloop()