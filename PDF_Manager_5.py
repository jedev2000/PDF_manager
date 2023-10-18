import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog
import PyPDF2
import os
import tabula
from pdf2docx import Converter
from PIL import Image, ImageTk
from tkinter import messagebox
import pandas as pd

# global panelC_visible, panelC
# panelC_visible = False

def merge_2_files():
    fichier = filedialog.askopenfilename(initialdir="/", title="Select PDF file 1",
                    filetypes=(("pdf files", "*.pdf"),("all files", "*.*")))
    if fichier:
        filename_1 = fichier
    else:
        return

    fichier = filedialog.askopenfilename(initialdir="/", title="Select PDF file 2",
                    filetypes=(("pdf files", "*.pdf"),("all files", "*.*")))
    if fichier:
        filename_2 = fichier
    else:
        return
    
    filelist = [filename_1,filename_2]
    Output_name = ".\PDF" + "\PDF_Merge_2.pdf"

    #create object
    merger = PyPDF2.PdfMerger()

    #loop through all PDF files in alphabetic order and merge 1 by 1
    for file in filelist:
        if file.endswith(".pdf"):
            merger.append(file)
    merger.write(Output_name)
    
    messagebox.showinfo("Information", "Operation successful : 2 files merged !")


def PDF2Word():
    fichier = filedialog.askopenfilename(initialdir="/", title="Select PDF file to convert",
                    filetypes=(("pdf files", "*.pdf"),("all files", "*.*")))
    if fichier:
        filename = fichier
    else:
        return

    # Spécifiez le chemin du fichier Word de sortie
    Output_name = ".\PDF" + "\PDF_Convert.docx"

    # Utilisez la classe Converter pour effectuer la conversion
    cv = Converter(filename)
    cv.convert(Output_name, start=0, end=None)

    # Fermez le convertisseur 
    cv.close()

    # print(f"Conversion de '{filename}' en '{Output_name}' terminée avec succès.")
    messagebox.showinfo("Information", "Operation successful : PDF converted in WORD in file : " + Output_name)


def PDF2Excel():
    fichier = filedialog.askopenfilename(initialdir="/", title="Select PDF file to convert",
                    filetypes=(("pdf files", "*.pdf"),("all files", "*.*")))
    if fichier:
        filename = fichier
    else:
        return


    df_list = tabula.read_pdf(filename, pages='all')

    # Convertir les données en un seul DataFrame Pandas (en concaténant)
    df = pd.concat(df_list)

    # Si vous souhaitez afficher les données
    # print(df)

    # Spécifiez le chemin du fichier Word de sortie
    Output_name_csv = ".\PDF" + "\PDF_Convert.csv"
    Output_name_xlsx = ".\PDF" + "\PDF_Convert.xlsx"

    df.to_csv(Output_name_csv, index=False)
    df.to_excel(Output_name_xlsx, index=False)

    # print(f"Conversion de '{filename}' en '{Output_name}' terminée avec succès.")
    messagebox.showinfo("Information", "Operation successful : PDF converted in .CSV in file : " + Output_name_xlsx + ". You can also check : https://www.adobe.com/acrobat/online/pdf-to-excel.html")

def merge_files():
    #create object
    merger = PyPDF2.PdfMerger()
    nb_file = 0

    repertoire = filedialog.askdirectory(title="Select a directory")
    output_name = repertoire+"/"+"PDF_Merge .pdf"
    if os.path.exists(output_name):
        os.remove(output_name)

    if repertoire:
        for file in os.listdir(repertoire):
            if file.endswith(".pdf"):
                nb_file += 1
                input_name = repertoire+"/"+file
                merger.append(input_name)
        merger.write(output_name)
        merger.close

    else:
        print("Nothing selected")

    messagebox.showinfo("Information", "Operation successful : " + str(nb_file) + " PDF files merged in file : " + output_name)


def split_file():
    # fichier à traiter
    filename = filedialog.askopenfilename(initialdir="/", title="Select file",
                        filetypes=(("pdf files", "*.pdf"),("all files", "*.*")))

    fichier = open(filename, 'rb')

    # objet PdfFileReader pour le fichier PDF
    lecteur = PyPDF2.PdfReader(fichier)

    # objet PdfFileWriter pour le fichier de sortie
    ecrivain = PyPDF2.PdfWriter()

    # Page de début, fin et verifications
    page_debut = int(simpledialog.askstring("Extract", "First page :"))
    if isinstance(page_debut, int):
        test = 1
    else:
        return
    
    page_fin_str = simpledialog.askstring("Extract", "Last page (E=end) :")

    if page_fin_str == "E" or page_fin_str == "e" :
        page_fin = len(lecteur.pages)

    elif isinstance(int(page_fin_str), int):
            page_fin = int(page_fin_str)
    else:
        return
    
    if page_debut > len(lecteur.pages):
        page_debut = 1

    if page_fin > len(lecteur.pages):
        page_fin = len(lecteur.pages)

    # Ajout des pages
    for page_num in range(page_debut-1, page_fin):
        #print(page_debut, page_num, page_fin)
        page = lecteur.pages[page_num]
        ecrivain.add_page(page)

    # nouveau fichier PDF pour la sortie : appli / PDF
    chemin_du_fichier =  '.\PDF' + '\PDF_extract.pdf'
    fichier_sortie = open(chemin_du_fichier, "wb")

    # Écrivez le contenu extrait dans le fichier de sortie
    ecrivain.write(fichier_sortie)

    # Fermez les fichiers
    fichier.close()
    fichier_sortie.close()
    messagebox.showinfo("Information", "Operation successful : file extracted from page " + str(page_debut) + " to page " + str(page_fin) + " in file : " + chemin_du_fichier)

def split_all():
    # fichier à traiter
    filename = filedialog.askopenfilename(initialdir="/", title="Select file",
                        filetypes=(("pdf files", "*.pdf"),("all files", "*.*")))

    fichier = open(filename, 'rb')

    # objet PdfFileReader pour le fichier PDF
    lecteur = PyPDF2.PdfReader(fichier)

    # objet PdfFileWriter pour le fichier de sortie
    ecrivain = PyPDF2.PdfWriter()

    # Page de début, fin et verifications
    page_debut = 1
    page_fin = len(lecteur.pages)

    # Ajout des pages
    for page_num in range(page_debut-1, page_fin):
        #print(page_debut, page_num, page_fin)
        ecrivain = PyPDF2.PdfWriter()
        page = lecteur.pages[page_num]
        ecrivain.add_page(page)
        # nouveau fichier PDF pour la sortie : appli / PDF
        chemin_du_fichier =  '.\PDF' + '\PDF_extract_' + str(page_num) + '.pdf'
        # print("file = ",chemin_du_fichier)
        fichier_sortie = open(chemin_du_fichier, "wb")
        # Écrivez le contenu extrait dans le fichier de sortie
        ecrivain.write(fichier_sortie)
        # Ferme les fichiers
        fichier_sortie.close()
        ecrivain=""

    # Ferme les fichiers
    fichier.close()
    messagebox.showinfo("Information", "Operation successful : PDF split in " + str(page_fin) + " pages in file : " + chemin_du_fichier)

def Compress_PDF():
    fichier = filedialog.askopenfilename(initialdir="/", title="Select PDF file to compress",
                    filetypes=(("pdf files", "*.pdf"),("all files", "*.*")))
    if fichier:
        filename = fichier
    else:
        return

    # Spécifiez le chemin du fichier Word de sortie
    Output_name = ".\PDF" + "\PDF_Compress.pdf"


    # Ouvrez le fichier PDF en mode lecture
    with open(filename, 'rb') as fichier_entree:
        # Ouvrez le fichier de sortie en mode écriture
        with open(Output_name, 'wb') as fichier_sortie:
            pdf = PyPDF2.PdfReader(fichier_entree)
            pdf_writer = PyPDF2.PdfWriter()

            # Copiez chaque page du PDF dans le PDFWriter (cela réduira la taille)
            for page_num in range(len(pdf.pages)):
                page = pdf.pages[page_num]
                pdf_writer.add_page(page)

            # Écrivez le PDF compressé dans le fichier de sortie
            pdf_writer.write(fichier_sortie)

    messagebox.showinfo("Information", "Operation successful : PDF compressed in file : " + Output_name)


def Remove_Pages():
    # fichier à traiter
    filename = filedialog.askopenfilename(initialdir="/", title="Select initial PDF file",
                        filetypes=(("pdf files", "*.pdf"),("all files", "*.*")))

    fichier = open(filename, 'rb')

    # objet PdfFileReader pour le fichier PDF
    lecteur = PyPDF2.PdfReader(fichier)
    page_fin = len(lecteur.pages)
    if page_fin == 0:
        return

    # objet PdfFileWriter pour le fichier de sortie
    ecrivain = PyPDF2.PdfWriter()

    # Page de début, fin et verifications
    Pages_in = simpledialog.askstring("PAGES", "Pages to remove, separated by ','")
    Pages_list=Pages_in.split(',',)
    Pages_int = [int(x) for x in Pages_list]
    Pages_max = [nombre for nombre in Pages_int if nombre <= page_fin]
    Pages_sorted = sorted(Pages_max)
    Pages_minus_1 = [x - 1 for x in Pages_sorted]
    Pages_final = list(set(Pages_minus_1))

    # Ajout des pages
    for page_num in range(0, page_fin):
        page = lecteur.pages[page_num]
            
        if page_num  not in Pages_final:
            ecrivain.add_page(page)

    # nouveau fichier PDF pour la sortie : appli / PDF
    chemin_du_fichier =  '.\PDF' + '\PDF_Removed_Pages.pdf'
    fichier_sortie = open(chemin_du_fichier, "wb")

    # Écrivez le contenu extrait dans le fichier de sortie
    ecrivain.write(fichier_sortie)

    # Fermez les fichiers
    fichier.close()
    fichier_sortie.close()
    messagebox.showinfo("Information", "Operation successful : " + str(len(Pages_final)) + " pages removed in file : " + chemin_du_fichier)


def Replace_1_Page():
    # fichier à traiter
    filename = filedialog.askopenfilename(initialdir="/", title="Select file to modify",
                        filetypes=(("pdf files", "*.pdf"),("all files", "*.*")))

    fichier = open(filename, 'rb')

    # objet PdfFileReader pour le fichier PDF
    lecteur = PyPDF2.PdfReader(fichier)
    page_fin = len(lecteur.pages)
    if page_fin == 0:
        return

    # objet PdfFileWriter pour le fichier de sortie
    ecrivain = PyPDF2.PdfWriter()

    # Page de début, fin et verifications
    Pages_in = simpledialog.askstring("PAGES", "Page to replace ? ")
    Pages_list=Pages_in.split(',',)
    Pages_replace = int(Pages_list[0])

    if Pages_replace > page_fin :
        return

    # page à ajouter
    filename_2 = filedialog.askopenfilename(initialdir="/", title="Select PDF with page 1 for replacement",
                        filetypes=(("pdf files", "*.pdf"),("all files", "*.*")))

    fichier_2 = open(filename_2, 'rb')

    # objet PdfFileReader pour le fichier PDF
    lecteur = PyPDF2.PdfReader(fichier)
    lecteur_2 = PyPDF2.PdfReader(fichier_2)
    page_fin = len(lecteur.pages)
    if page_fin == 0:
        return


    # Ajout des pages
    for page_num in range(0, page_fin):
        #page = lecteur.pages[page_num]
            
        if page_num  == Pages_replace - 1:
            ecrivain.add_page(lecteur_2.pages[0])
        else :
            page = lecteur.pages[page_num]
            ecrivain.add_page(page)

    # nouveau fichier PDF pour la sortie : appli / PDF
    chemin_du_fichier =  '.\PDF' + '\PDF_Replace_1_Page.pdf'
    fichier_sortie = open(chemin_du_fichier, "wb")

    # Écrivez le contenu extrait dans le fichier de sortie
    ecrivain.write(fichier_sortie)

    # Fermez les fichiers
    fichier.close()
    fichier_sortie.close()
    messagebox.showinfo("Information", "Operation successful : 1 page replaced in file : " + chemin_du_fichier)

def PDF_data():
    fichier = filedialog.askopenfilename(initialdir="/", title="Select PDF file",
                    filetypes=(("pdf files", "*.pdf"),("all files", "*.*")))
    if fichier:
        filename = fichier
    else:
        return

    # Ouvrez le fichier PDF en mode lecture
    with open(filename, 'rb') as fichier_entree:
            pdf = PyPDF2.PdfReader(fichier_entree)
            file_data = pdf.metadata

            if file_data.author:
                author = str(file_data.author)
            else:
                author = "None"

            if file_data.creator:
                creator = str(file_data.creator)
            else:
                creator = "None"

            if file_data.subject:
                subject = str(file_data.subject)
            else:
                subject = "None"           

            if file_data.producer:
                producer = str(file_data.producer)
            else:
                producer = "None"   

            if file_data.title:
                title = str(file_data.title)
            else:
                title = "None"  

            Infos =  " Title : " + title + "\n Number of pages : " + str(len(pdf.pages)) + "\n Author : " + author + "\n Producer : " + producer + "\n Creator : " + creator + "\n Subject : " + subject


    messagebox.showinfo("Information", Infos)

# Fonction pour afficher ou masquer le panneau
def toggle_panel():
    global panelC_visible
    global panelC
    if panelC_visible:
        panelC.grid_remove()  # Pour masquer le panneau
    else:
        panelC.grid(row=1, column=0)  # Pour afficher le panneau
    panelC_visible = not panelC_visible  # Inverser l'état

# Fonction pour créer une nouvelle fenêtre avec du texte
def ouvrir_fenetre_texte():
    nouvelle_fenetre = tk.Toplevel()  # Crée une nouvelle fenêtre
    nouvelle_fenetre.title("Information")
    
    # Définir la position de la nouvelle fenêtre
    nouvelle_fenetre.geometry("400x605+192+100")  # Largeur x Hauteur + X + Y

    message1 = "\n\nHELP AND INFORMATION\n\n\n\n MERGE 2 PDF = Merge two PDF in one\n\n\nMERGE ALL PDF = Merge all the PDF in a specified directory"
    message2 ="\n\n\nSPLIT ALL PAGES = Split a PDF in all its pages\n\n\n EXTRACT N PAGES = Extract a specified part of a PDF\n\n\n REMOVE PAGES = Remove specified pages from a PDF"
    message3 = "\n\n\nREPLACE 1 PAGE = Replace one page of a PDF \n\n\nPDF TO WORD = Convert a PDF to WORD \n\n\nPDF TO EXCEL = Convert a PDF to CSV and EXCEL files"
    message4 = "\n\n\nCOMPRESS A PDF = Compress a PDF \n\n\nPDF DATA = Display PDF data"
    message5 = "\n\n\nOutput files are store in sub directory /PDF/"
    message6 = "\n\n\nCo Jerome C. - 2023"
    message = message1 + message2 + message3 + message4 + message5 + message6
    # Ajoutez un widget Label pour afficher du texte
    texte_label = tk.Label(nouvelle_fenetre, text = message)
    texte_label.pack()

def main():
    # Create Directory for output PDF files
    output_directory = "./PDF"
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    # FRAMES
    # Main window
    fenetre = tk.Tk()
    fenetre.title("File selection")
    fenetre.geometry("190x605+0+100")
    fenetre.title("JC Apps - 2023")
    my_font = ("Arial", 10)
    my_font_large = ("Arial", 12, 'bold')
    my_font_italic = ("Arial", 10, 'italic')
    my_font_title = ("Arial", 15, 'bold')

    # Frame 0 : Title
    frame_0 = tk.Frame(fenetre)
    frame_0.grid(row=0, column=0, columnspan=1, sticky='nsew')
    label= tk.Label(frame_0,font = my_font_title, text="PDF MANAGER", fg='#7961f2', justify="center")
    label.grid(padx=(15,0), pady=(25,0), sticky="nsew")

    # Frame 1 : to merge 2 PDF
    frame_1 = tk.Frame(fenetre)
    frame_1.grid(row=1, column=0)


    # Frame 2 : to merge all PDF contained in a directory
    frame_2 = tk.Frame(fenetre)
    frame_2.grid(row=2, column=0)


    # Frame 3 : to split a PDF in all its pages
    frame_3 = tk.Frame(fenetre)
    frame_3.grid(row=3, column=0)


    # Frame 4 : Extract pages from a PDF
    frame_4 = tk.Frame(fenetre)
    frame_4.grid(row=4, column=0)


    # Frame 5 : Remove pages
    frame_5 = tk.Frame(fenetre)
    frame_5.grid(row=5, column=0)


    # Frame 6 : Replace 1 Page
    frame_6 = tk.Frame(fenetre)
    frame_6.grid(row=6, column=0)


    # Frame 7 : PDF 2 Word
    frame_7 = tk.Frame(fenetre)
    frame_7.grid(row=7, column=0)

    # Frame 8 : PDF2Excel
    frame_8 = tk.Frame(fenetre)
    frame_8.grid(row=8, column=0)

    # Frame 9 : Compress PDF
    frame_9 = tk.Frame(fenetre)
    frame_9.grid(row=9, column=0)

    # Frame 11 : PDF Data
    frame_11 = tk.Frame(fenetre)
    frame_11.grid(row=10, column=0)

    # Créer un panneau avec des widgets à l'intérieur
    global panelC, panelC_visible
    panelC = tk.Frame(fenetre)
    labelC = tk.Label(panelC, text="HELP & INFORMATION")
    labelC.grid(row=3, column=0)

    # Variable pour suivre l'état du panneau
    panelC_visible = tk.BooleanVar()
    panelC_visible.set(False)  # Initialement masqué

    #Toggle Panel
    # Créer un bouton pour afficher/masquer le panneau
    # toggle_button = tk.Button(fenetre, text="Help & information", command=toggle_panel)
    toggle_button = tk.Button(fenetre, text="Help & information", command=ouvrir_fenetre_texte)
    toggle_button.grid(row=11, column=0, pady=(10))



    # BUTTONS & COMMANDS
    # 1 - Merge 2 PDF
    # bouton_ouvrir = tk.Button(frame_1, text="Merge 2 PDF", image=photo, width=20, height=2, bg='#0268af', fg='white', command=merge_2_files)
    image1 = Image.open("./Images/Logo_Merge_2.png")
    photo1 = ImageTk.PhotoImage(image1)
    bouton_ouvrir = tk.Button(frame_1, image=photo1, text="   MERGE 2 PDF", compound=tk.LEFT, width=160, height=40, bg='#7961f2', fg='white', anchor="w", command=merge_2_files)
    bouton_ouvrir.grid(padx=10,pady=(25,0))
 
    # 2 - Merge all PDF
    image2 = Image.open("./Images/Logo_Merge.png")
    photo2 = ImageTk.PhotoImage(image2)
    bouton_ouvrir = tk.Button(frame_2, text="  MERGE ALL PDF", image=photo2, compound=tk.LEFT, width=160, height=40, bg='#7961f2', fg='white', anchor="w", command=merge_files)
    bouton_ouvrir.grid()

    # 3 - PDF Extract
    image3 = Image.open("./Images/Logo_Extract_6.png")
    photo3 = ImageTk.PhotoImage(image3)
    bouton_ouvrir = tk.Button(frame_4, text="  EXTRACT N PAGES", image=photo3, compound=tk.LEFT, width=160, height=40, bg='#7961f2', fg='white', anchor="w", command=split_file)
    bouton_ouvrir.grid()

    # 3 - Split PDF in all its pages
    image4 = Image.open("./Images/Logo_ciseaux.png")
    photo4 = ImageTk.PhotoImage(image4)
    bouton_ouvrir = tk.Button(frame_3, text="  SPLIT ALL PAGES", image=photo4, compound=tk.LEFT, width=160, height=40, bg='#7961f2', fg='white', anchor="w", command=split_all)
    bouton_ouvrir.grid()

    # 5 - PDF to Word
    image5 = Image.open("./Images/Logo_Word_2.png")
    photo5 = ImageTk.PhotoImage(image5)
    bouton_ouvrir = tk.Button(frame_7, text="  PDF TO WORD", image=photo5, compound=tk.LEFT, width=160, height=40, bg='#7961f2', fg='white', anchor="w", command=PDF2Word)
    bouton_ouvrir.grid()

    # 7 - Compress PDF
    image6 = Image.open("./Images/Logo_Compress.png")
    photo6 = ImageTk.PhotoImage(image6)
    bouton_ouvrir = tk.Button(frame_9, text="  COMPRESS A PDF", image=photo6, compound=tk.LEFT, width=160, height=40, bg='#7961f2', fg='white', anchor="w", command=Compress_PDF)
    bouton_ouvrir.grid()

    # 8 - Remove pages
    image7 = Image.open("./Images/Logo_Remove.png")
    photo7 = ImageTk.PhotoImage(image7)
    bouton_ouvrir = tk.Button(frame_5, text="  REMOVES PAGES", image=photo7, compound=tk.LEFT, width=160, height=40, bg='#7961f2', fg='white', anchor="w", command=Remove_Pages)
    bouton_ouvrir.grid()

    # 9 - Convert 2 Excel
    image8 = Image.open("./Images/Logo_Excel.png")
    photo8 = ImageTk.PhotoImage(image8)
    bouton_ouvrir = tk.Button(frame_8, text=" PDF TO EXCEL", image=photo8, compound=tk.LEFT, width=160, height=40, bg='#7961f2', fg='white', anchor="w", command=PDF2Excel)
    bouton_ouvrir.grid()

    # 10 - Replace 1 page
    image9 = Image.open("./Images/Logo_Replace_1_Page.png")
    photo9 = ImageTk.PhotoImage(image9)
    bouton_ouvrir = tk.Button(frame_6, text="  REPLACE 1 PAGE", image=photo9, compound=tk.LEFT, width=160, height=40, bg='#7961f2', fg='white', anchor="w", command=Replace_1_Page)
    bouton_ouvrir.grid()

    # 10 - PDF Data
    image10 = Image.open("./Images/Logo_Data.png")
    photo10 = ImageTk.PhotoImage(image10)
    bouton_ouvrir = tk.Button(frame_11, text="   PDF DATA", image=photo10, compound=tk.LEFT, width=160, height=40, bg='#7961f2', fg='white', anchor="w", command=PDF_data)
    bouton_ouvrir.grid()

    fenetre.mainloop()

if __name__ == "__main__":
    main()
