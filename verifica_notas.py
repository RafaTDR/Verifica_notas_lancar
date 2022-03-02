from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import xml.etree.ElementTree as ET
import pandas as pd
import os
import shutil

tkWindow = Tk()
tkWindow.geometry('400x150')
tkWindow.title('Verifica Notas')
tkWindow.eval('tk::PlaceWindow . center')
tkWindow.iconbitmap("xml-ico.ico")

def verifica():
    messagebox.showinfo('Processo', 'Selecionar a pasta com os arquivos XML.')
    path = filedialog.askdirectory(title="Arquivos XML")
    messagebox.showinfo('Processo', 'Selecionar o relatório em Excel.')
    pathexcel = filedialog.askopenfilename(title="Relatório Excel")
    messagebox.showinfo('Processo', 'Selecionar a pasta onde deseja salvar os XML para lançar.')
    destino = filedialog.askdirectory(title="Salvar XML")

    if str(path) == "" or str(pathexcel) == "" or str(destino) == "":
        messagebox.showinfo('Processo', 'Selecionar um caminho valido.')
        pass
    else:

        tkWindow.config(cursor="circle")
        tkWindow.update()

        for filename in os.listdir(path):
            if not filename.endswith('.xml'): continue
            fullname = os.path.join(path, filename)
            tree = ET.parse(fullname)

            doc = tree.getroot()
            nodefind = doc.find(
                '{http://www.portalfiscal.inf.br/nfe}NFe/{http://www.portalfiscal.inf.br/nfe}infNFe/{http://www.portalfiscal.inf.br/nfe}det')

            for ide in doc.iter('{http://www.portalfiscal.inf.br/nfe}nNF'):
                numeronfe = ide.text

            for emit in doc.iter('{http://www.portalfiscal.inf.br/nfe}emit'):

                for CNPJ in emit.iter('{http://www.portalfiscal.inf.br/nfe}CNPJ'):
                    cnpjemit = CNPJ.text

            localizador = str(numeronfe+cnpjemit)
            plan_lancar = pd.read_excel(str(pathexcel))
            for localizador_plan in plan_lancar["LOCALIZADOR"]:
                if localizador == localizador_plan:
                    shutil.move(path+"/"+filename,destino)
                else:
                    print(localizador_plan)
                    print(localizador)

        messagebox.showinfo('Processo', 'Processo concluido')
        tkWindow.config(cursor="")
        tkWindow.update()




button1 = Button(tkWindow,
                 text='VERIFICA NFE',
                 command=verifica)
button1.grid(row=1, column=0, columnspan=5, padx=100, ipadx=80)

tkWindow.mainloop()

