from pathlib import Path
import sys
import io
import math
import glob
import os
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
from ttkbootstrap import Style, Colors, Entry, Button
import tkinter as tk
from tkinter import *
# Explicit imports to satisfy Flake8
from tkinter import Tk, ttk, Entry, Text, Button, PhotoImage, scrolledtext
# Importing tkinter Canvas as tkcanvas instead of canvas => reportlab canvas
from tkinter import Canvas as tkcanvas
import tkinter.messagebox as tk1
import tkinter.filedialog
from datetime import datetime
import threading
from threading import Timer
import argparse
from pdf2image import convert_from_path
from reportlab.pdfgen import canvas
from reportlab.lib import pagesizes
from PIL import Image, ImageSequence, ImageTk, ImageDraw
from pypdf import PdfWriter, PdfReader
from azure.core.credentials import AzureKeyCredential
from azure.ai.formrecognizer import DocumentAnalysisClient
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from arabic_reshaper import reshape
from bidi.algorithm import get_display
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import Color, black
from reportlab.graphics.shapes import Rect
import webbrowser



if getattr(sys, "frozen", False):
    fonts = os.path.join(
        sys._MEIPASS, "fonts"
    )  
    poppler_bin = os.path.join(sys._MEIPASS, "poppler-bin")
    images = os.path.join(sys._MEIPASS, "images")
else:
    fonts = os.path.join(
        os.path.dirname(__file__), "fonts"
    )  
    poppler_bin = os.path.join(os.path.dirname(__file__), "poppler-bin")
    images = os.path.join(os.path.dirname(__file__), "images")

#Path to font
ttfFile = os.path.join(fonts, "Al-Nile-Regular.ttf")

#Registering the font
pdfmetrics.registerFont(TTFont("Al Nile", ttfFile))

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

path = getattr(sys, "_MEIPASS", os.getcwd())
os.chdir(path)


# Azure Key
global key
key = ""

# Azure Endpoint
global endpoint
endpoint = ""

# Input Path
global input_path
input_path = ""

# Output Path
global output_path
output_path = ""

# Counter
global counter
counter = 0

# Total
global total
total = 0

# Interruption Switch
global is_on
is_on = True

# Unlocking the go_back function
global unlock
unlock = False

# Toggle: Activating/Deactivating the confidence index
global confidence
confidence = 0

# Corrupt files counter
global corrupt
corrupt = 0

# Corrupt files names
global corrupt_names
corrupt_names = []


class App:
    def __init__(self, root=None):
        #Custom entry style
        ttkb.Style("new").configure(
            "custom.TEntry",
            background="#F1F5FF",
            selectbackground="#e8e846",
            foreground="#a2a6b0",
            selectforeground="#292928",
            bordercolor="",
            borderwidth=5,
            relief="ridge",
            selectbordercolor="#F1F5FF",
            fieldbackground="#F1F5FF",
            selectborderwidth=1,
            font=("Helvetica", 24),
        )

        ttkb.Style("new").map(
            "custom.TEntry",
            foreground=[
                ("disabled", "#a2a6b0"),
                ("focus !disabled", "#272729"),
                ("hover !disabled", "#a2a6b0"),
            ],
        )

        ttkb.Style("new").map(
            "custom.TEntry",
            bordercolor=[
                ("disabled", "#F1F5FF"),
                ("focus !disabled", "#F1F5FF"),
                ("hover !disabled", "#F1F5FF"),
            ],
        )

        self.root = root
        self.frame = ttkb.Frame(self.root, height=594, width=872, bootstyle="primary")
        page00 = self.frame
        self.tasks = Tasks(master=self.root, app=self)
        self.frame.pack()

        canvas = tkcanvas(
            page00,
            bg="#3A7FF6",
            height=594,
            width=872,
            bd=0,
            highlightthickness=0,
            relief="ridge",
        )
        canvas.place(x=0, y=0)

        canvas.create_rectangle(436.0, 0.0, 872.0, 594.0, fill="#FCFCFC", outline="")
        canvas.create_rectangle(0.0, 0.0, 436.0, 594.0, fill="#3A7FF6", outline="")

        self.image_image_1 = PhotoImage(file=os.path.join(images, "image_1.png"))
        self.image_1 = canvas.create_image(653.0, 397.0, image=self.image_image_1)

        self.entry_image_1 = PhotoImage(file=os.path.join(images, "entry_1.png"))
        self.entry_bg_1 = canvas.create_image(635.0, 403.5, image=self.entry_image_1)
        self.entry_1 = ttkb.Entry(page00, style="custom.TEntry")
        self.entry_1.place(x=498.0, y=390.0, width=274.0, height=25.0)

        self.image_image_2 = PhotoImage(file=os.path.join(images, "image_2.png"))
        self.image_2 = canvas.create_image(653.0, 315.0, image=self.image_image_2)

        self.entry_image_2 = PhotoImage(file=os.path.join(images, "entry_2.png"))
        self.entry_bg_2 = canvas.create_image(635.0, 323.5, image=self.entry_image_2)
        self.entry_2 = ttkb.Entry(page00, style="custom.TEntry")
        self.entry_2.place(x=498.0, y=310.0, width=274.0, height=25.0)

        self.image_image_3 = PhotoImage(file=os.path.join(images, "image_3.png"))
        self.image_3 = canvas.create_image(653.0, 233.0, image=self.image_image_3)

        self.entry_image_3 = PhotoImage(file=os.path.join(images, "entry_3.png"))
        self.entry_bg_3 = canvas.create_image(635.0, 241.5, image=self.entry_image_3)
        self.entry_3 = ttkb.Entry(page00, style="custom.TEntry")
        self.entry_3.place(x=498.0, y=228.0, width=274.0, height=25.0)

        canvas.create_text(40.0,
            127.0,
            anchor="nw",
            text="Outil OCR Azure",
            fill="#FCFCFC",
            font=("Roboto Bold", 24 * -1),
        )

        canvas.create_text(
            481.0,
            46.0,
            anchor="nw",
            text="Informations.",
            fill="#505485",
            font=("Roboto Bold", 24 * -1),
        )

        canvas.create_rectangle(40.0, 160.0, 100.0, 165.0, fill="#FCFCFC", outline="")

        canvas.create_rectangle(109.0, 396.0, 130.0, 401.0, fill="#FCFCFC", outline="")

        canvas.create_text(
            40.0,
            191.0,
            anchor="nw",
            text="Ce programme utilise l’outil",
            fill="#FCFCFC",
            font=("Inter", 20 * -1),
        )

        canvas.create_text(
            40.0,
            215.0,
            anchor="nw",
            text="de reconnaissance optique",
            fill="#FCFCFC",
            font=("Inter", 20 * -1),
        )

        canvas.create_text(
            40.0,
            239.0,
            anchor="nw",
            text="des caractères Azure",
            fill="#FCFCFC",
            font=("Inter", 20 * -1),
        )

        canvas.create_text(
            40.0,
            262.0,
            anchor="nw",
            text="Document Intelligence AI",
            fill="#FCFCFC",
            font=("Inter", 20 * -1),
        )

        canvas.create_text(
            40.0,
            283.0,
            anchor="nw",
            text="pour transformer les PDF",
            fill="#FCFCFC",
            font=("Inter", 20 * -1),
        )

        canvas.create_text(
            40.0,
            305.0,
            anchor="nw",
            text="scannés en documents lisibles.",
            fill="#FCFCFC",
            font=("Inter", 20 * -1),
        )

        canvas.create_text(
            40.0,
            349.0,
            anchor="nw",
            text="Pour plus d’informations,",
            fill="#FCFCFC",
            font=("Inter", 20 * -1),
        )

        def know_more_clicked(event):
            instructions = (
        "https://redazaireg.com/ocr.html"
        )
            webbrowser.open_new_tab(instructions)
            
        know_more = ttk.Label(
            page00,
            text="cliquer ici.",
            foreground="#FCFCFC",
            background="#3A7FF6",
            font=("Inter", 20 * -1),
        )
        know_more.place(x = 40.0, y = 371.0)
        know_more.bind('<Button-1>', know_more_clicked)


        self.button_image_1 = PhotoImage(file=os.path.join(images, "button_1.png"))
        self.button_1 = Button(
            page00,
            image=self.button_image_1,
            borderwidth=0,
            highlightthickness=0,
            command= lambda: self.make_page_2(),
            relief="flat",
        )
        self.button_1.place(x=560.0, y=500.0, width=180.0, height=55.0)
        self.page2 = Tasks(master=self.root, app=self)

        canvas.create_text(
            493.0,
            372.0,
            anchor="nw",
            text="Output folder",
            fill="#3A7FF6",
            font=("Roboto Bold", 14 * -1),
        )

        self.button_image_2 = PhotoImage(file=os.path.join(images, "button_2.png"))
        self.button_2 = Button(
            page00,
            image=self.button_image_2,
            borderwidth=0,
            highlightthickness=0,
            command= lambda: select_output_path(),
            relief="flat",
        )
        self.button_2.place(x=786.0, y=393.0, width=24.0, height=22.0)

        canvas.create_text(
            493.0,
            292.0,
            anchor="nw",
            text="Input folder",
            fill="#3A7FF6",
            font=("Roboto Bold", 14 * -1),
        )

        self.button_image_3 = PhotoImage(file=os.path.join(images, "button_3.png"))
        self.button_3 = Button(
            page00,
            image=self.button_image_3,
            borderwidth=0,
            highlightthickness=0,
            command= lambda: select_input_path(),
            relief="flat",
        )
        self.button_3.place(x=786.0, y=313.0, width=24.0, height=22.0)

        canvas.create_text(
            493.0,
            210.0,
            anchor="nw",
            text="Endpoint",
            fill="#3A7FF6",
            font=("Roboto Bold", 14 * -1),
        )

        self.image_image_4 = PhotoImage(file=os.path.join(images, "image_4.png"))
        self.image_4 = canvas.create_image(653.0, 151.0, image=self.image_image_4)

        self.image_image_5 = PhotoImage(file=os.path.join(images, "image_5.png"))
        self.image_5 = canvas.create_image(653.0, 151.0, image=self.image_image_5)

        canvas.create_text(
            493.0,
            127.0,
            anchor="nw",
            text="Clé Azure",
            fill="#3A7FF6",
            font=("Roboto Bold", 14 * -1),
        )

        entry_image_4 = PhotoImage(file=os.path.join(images, "entry_4.png"))
        self.entry_bg_4 = canvas.create_image(635.0, 159.5, image=entry_image_4)
        self.entry_4 = ttkb.Entry(page00, style="custom.TEntry")
        self.entry_4.place(x=498.0, y=146.0, width=274.0, height=25.0)

        def select_input_path():
            global input_path
            input_path = tk.filedialog.askdirectory()
            self.entry_2.delete(0, tk.END)
            self.entry_2.insert(0, input_path)
            return

        def select_output_path():
            global output_path
            output_path = tk.filedialog.askdirectory()
            self.entry_1.delete(0, tk.END)
            self.entry_1.insert(0, output_path)
            return

        #Confidence index toggle command
        var1 = IntVar()
        def checker():
            global confidence
            confidence = 1 if var1.get() == 1 else 0

        #Confidence index toggle
        confidence_index = ttkb.Checkbutton(page00, bootstyle="square-toggle", variable = var1, onvalue = 1, offvalue = 0, command = checker)
        confidence_index.place(x=786.0, y=445.0, width=24.0, height=22.0)
        canvas.create_text(
            650.0,
            445.0,
            anchor="nw",
            text="Index de confiance",
            fill="#3A7FF6",
            font=("Roboto Bold", 14 * -1),
        )

        #Packing the main page
        page00.pack()

    #Def: for going back to the main page
    def main_page(self):
        self.frame.pack()


    #Generating the credentials, input and output path
    def generate(self):
        global key
        key = self.entry_4.get()
        global endpoint
        endpoint = self.entry_3.get()
        global input_path
        input_path = self.entry_2.get()
        input_path = input_path.strip()
        global output_path
        output_path = self.entry_1.get()
        output_path = output_path.strip()

        if not key:
            tk.messagebox.showerror(
                title="Champ vide", message="Entrez une clé Azure."
            )
            return
        if not endpoint:
            tk.messagebox.showerror(
                title="Champ vide", message="Entrez un endpoint."
            )
            return
        if not input_path:
            tk.messagebox.showerror(
                title="Chemin invalide", message="Entrez un input path."
            )
            return
        if not output_path:
            tk.messagebox.showerror(
                title="Chemin invalide", message="Entrez un output path."
            )
            return

        key = key.strip()
        endpoint = endpoint.strip()
        input =  Path(f"{input_path}/build").expanduser().resolve()
        output = Path(f"{output_path}/build").expanduser().resolve()
            
        if input.exists() and not input.is_dir():
                tk1.showerror(
            "Erreur",
            f"{input} existe déjà et n'est pas un dossier.\n"
            "Entrez un input path valide.")
                
                
        if output.exists() and not output.is_dir():
            tk1.showerror(
                "Erreur",
                f"{output} existe déjà et n'est pas un dossier.\n"
                "Entrez un output path valide.")
                    
        elif output.exists() and output.is_dir() and tuple(output.glob('*')):
            response = tk1.askyesno(
                "Continuer ?",
                f"Le dossier {output} n'est pas vide.\n"
                "Souhaitez-vous continuer ?")
            if not response:
                return
                    
    #Page 2 
    def make_page_2(self):
        self.__init__
        self.frame.pack_forget()
        self.tasks.start_page()
        
        def make():
            App.generate(self)
            (threading.Thread(target=batch(input_path, output_path)).start)
        t = Timer(2, make)
        t.start()


class Console(Text):

  def __init__(self, *args, **kwargs):
    super().__init__(*args, **kwargs)
    self.config(state="disabled")
    self.bind("<Destroy>", self.reset)

    self.old_stdout = sys.stdout
    self.old_stderr = sys.stderr
    sys.stdout = self
    sys.stderr = self

  def delete(self, *args, **kwargs):
    self.config(state="normal")
    super().delete(*args, **kwargs)
    self.config(state="disabled")

  def write(self, content):
    self.config(state="normal")
    self.insert("end", content)
    self.config(state="disabled")

  def reset(self):
    sys.stdout = sys.__stdout__
    sys.stderr = sys.__stderr__

  def flush(self):
    pass

class Tasks:
    def __init__(self, master=None, app=None):
        self.master = master
        self.app = app
        self.root = root

        self.frame = ttkb.Frame(self.master, bootstyle="primary", width=872, height=594)
        subframe = self.frame

        #Interruption button style
        ttkb.Style("new").configure(
            "custom.TButton",
            font=("Roboto Bold", 14),
            width=20,
            height=10,
        )
        ttkb.Style("new").map(
            "custom.TButton",
            background=[
                ("disabled", "white"),
                ("active", "#d9534f"),
                ("pressed", "#d9534f"),
            ],
        )

        #Left canvas
        canvas = tkcanvas(
            subframe,
            bg="#3A7FF6",
            height=594,
            width=436,
            bd=0,
            highlightthickness=0,
            relief="ridge",
        )

        canvas.place(x=0, y=0)
        canvas.create_rectangle(0.0, 0.0, 872.0, 594.0, fill="#3A7FF6", outline="")

        canvas.create_text(
            35.0,
            32.0,
            anchor="nw",
            text="Tâches",
            fill="#FCFCFC",
            font=("Roboto Bold", 24 * -1),
        )

        canvas.create_rectangle(35.0, 65.0, 95.0, 70.0, fill="#FCFCFC", outline="")

        self.button_return_image = PhotoImage(file=os.path.join(images,"return_on.png"))
        self.button_return = Button(
            subframe,
            image=self.button_return_image,
            borderwidth=0,
            highlightthickness=0,
            command= self.go_back,
            relief="flat",
        )    

        #Right frame & right canvas
        self.frame2 = ttkb.Frame(self.frame, bootstyle="light", width=436, height=594)
        subframe2 = self.frame2
        canvas1 = tkcanvas(
            subframe2,
            bg="#ffffff",
            height=594,
            width=436,
            bd=0,
            highlightthickness=0,
            relief="ridge",
        )
        canvas1.place(x=0, y=0)

        canvas1.create_text(
            35.0,
            32.0,
            anchor="nw",
            text="Progression",
            fill="#3A7FF6",
            font=("Roboto Bold", 24 * -1),
        )
        canvas1.create_rectangle(35.0, 65.0, 95.0, 70.0, fill="#3A7FF6", outline="")

        #Meter
        self.meter = ttkb.Meter(
            master=subframe2,
            metersize=400,
            padding=5,
            metertype="full",
            textright="%",
            interactive=FALSE,
        )
        self.meter.place(relx=0.15, rely=0.5, anchor="w")


        #Interruption button
        self.button_off = ttkb.Button(
            subframe2,
            style="custom.TButton",
            text="Interrompre l'opération",
            command=self.switch,
        )
        self.button_off.place(relx=0.27, rely=0.9, anchor="sw")

        #Meter subtext
        self.progress = ttkb.Label(
            subframe2,
            font=("Roboto Bold", 10 * -1),
            bootstyle = "secondary"
        )
        self.progress.place(relx=0.36, rely=0.57, anchor="sw")

        #Console export
        self.button_out = ttkb.Button(
            subframe,
            text="Exporter la console",
            command= lambda: select_text_out(),
        )
        self.button_out.place(relx=0.32, rely=0.9, anchor="se")

        def select_text_out():
            textout_path = tk.filedialog.askdirectory()
            textout_path = textout_path.strip()
            txt_output = os.path.join(textout_path, f"ocr-log-{datetime.now():%d-%m-%Y_%Hh%M}.txt")
            stdout = self.terminal.get("1.0",'end-1c')
            with open(txt_output, "w") as f:
                f.write(stdout)

            f.close()
            



    #Widget loop (1s refresh)
    def counter_meter(self):
         if unlock:
             self.button_return.place(x=10, y=35, width=14.0, height=24.0) 
         else:
             self.button_return.place_forget()
         if is_on:
              self.button_off['state'] = 'enabled'
              if counter == 0 or total == 0:
                  self.set_meter(0, 'warning')
              elif counter < total:
                  self.set_meter(int(counter/total*100), 'primary')
              else:
                  self.set_meter(100, 'success')
                  self.button_off['state'] = 'disabled'
                  return
         else:
             self.button_off['state'] = 'disabled'
             if counter == 0 or total == 0: 
                 self.set_meter(0, 'danger')
             else:
                 self.set_meter(int(counter/total*100), 'danger')
                 if unlock:
                     return
                    
        

         self.schedule_next_update()
        

    def set_meter(self, amount, style):
        self.meter.configure(amountused=amount)
        self.meter['bootstyle'] = style
        self.progress.configure(text=f"Fichiers traités: {counter} sur {total}\n Fichiers corrompus: {corrupt}")     

    def schedule_next_update(self):
        self.button_off.after(1000, self.counter_meter)

    def switch(self):
        global is_on
        # Interrupting the process
        if is_on:
            is_on = False
            print(f"✕ {datetime.now():%d/%m/%Y %H:%M:%S}: L'opération sera interrompue à la fin du processus en cours\n")
        else:
            is_on = True


    def start_page(self):
        self.frame.pack(expand=TRUE, fill=BOTH, side=LEFT)
        self.frame2.pack(expand=FALSE, fill=BOTH, side=RIGHT)
        (threading.Thread(target=self.counter_meter()).start)
        terminal_config = {
            "bd": 0,
            "bg": "#3A7FF6",
            "fg": "#000716",
            "highlightthickness": 0, 
            "width": 40,
            "height": 10,
            "font": ("Roboto Bold", 14),
            "autostyle": False,
            "foreground": "#FFFFFF"
            }

        self.terminal = Console(root, **terminal_config)
        self.terminal.place(x=55, y=147, width=360, height=330)
        
        

    def go_back(self):
        self.reset_globals()
        self.reset_ui()
        self.terminal.reset()
        self.app.main_page()

    def reset_globals(self):
        global counter, total, is_on, corrupt, corrupt_names, unlock
        counter = 0
        total = 0 
        is_on = True
        corrupt = 0
        unlock = False
        corrupt_names.clear()

    def reset_ui(self):
        self.button_return.place_forget()
        self.frame.pack_forget()
        self.frame2.pack_forget()
        self.terminal.place_forget() 


#Popup notification def
def displayNotification(message,title=None,subtitle=None,soundname=None):

    titlePart = 'with title "{0}"'.format(title) if title is not None else ''
    subtitlePart = '' if subtitle is None else 'subtitle "{0}"'.format(subtitle)
    soundnamePart = ''
    if soundname is not None:
        soundnamePart = 'sound name "{0}"'.format(soundname)

    appleScriptNotification = 'display notification "{0}" {1} {2} {3}'.format(message,titlePart,subtitlePart,soundnamePart)
    os.system("osascript -e '{0}'".format(appleScriptNotification))

#Azure OCR Polygons calc (Euclidian distance between points):
def dist(p1, p2):
    return math.sqrt((p1.x - p2.x) * (p1.x - p2.x) + (p1.y - p2.y) * (p1.y - p2.y))

#OCR module
def ocr(input_path, output_path, input_file):
    global counter, corrupt, corrupt_names, is_on, unlock
    
    input_filename = input_file.name
    input_filepath = os.path.join(input_path, input_file)
    input_dir = os.path.dirname(input_filepath)
    subdir = os.path.split(os.path.split(input_path)[0])[1]

    x = input_dir.split(subdir)
    output_dir = output_path + x[1]

    if not os.path.exists(output_path):
        os.makedirs(output_path)

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    output_filename = input_filename.split(".")[0]
    output_file = os.path.join(output_dir, output_filename + "_ocr.pdf")
    txt_file = os.path.join(output_dir, output_filename + "_ocr.txt")

    if not is_on:
        return

    # Loading input file
    print(f"----Chargement du fichier {input_filename}----\n")
    if not os.path.exists(output_file):
        try:
            # read existing PDF as images
            image_pages = convert_from_path(
                input_file,
                poppler_path=poppler_bin,
                dpi=150,
                fmt="JPEG",
                jpegopt={"quality": 90, "progressive": True, "optimize": True},
            )
        except Exception as error:
            print(f"✕ {datetime.now():%d/%m/%Y %H:%M:%S}: Le fichier {input_filename} semble compromis\n")
            print("An error occurred:", type(error).__name__, "–", error)
            counter += 1
            corrupt += 1
            corrupt_names.append(input_file)
            return
            

        # Running OCR using Azure Form Recognizer Read API
        print(f"{datetime.now():%d/%m/%Y %H:%M:%S}: Démarrage d'Azure Form Recognizer...\n")
        document_analysis_client = DocumentAnalysisClient(
        endpoint=endpoint,
        credential=AzureKeyCredential(key),
        headers={"x-ms-useragent": "searchable-pdf-blog/1.0.0"},
        )

        try:
            with open(input_file, "rb") as f:
                poller = document_analysis_client.begin_analyze_document(
                "prebuilt-read", document=f
                )
        except Exception as error:
            is_on = False
            unlock = True
            displayNotification("Erreur!",f"Une erreur est survenue: {error}", "", "Submarine")
            tk.messagebox.showerror(
                "Erreur!", f"Une erreur est survenue: {error}")
            sys.exit(
                 print("Une erreur est survenue:", type(error).__name__, "–", error)
                 )            

        ocr_results = poller.result()
        print(
            f"{datetime.now():%d/%m/%Y %H:%M:%S}: Azure Form Recognizer a fini l'OCRisation de {len(ocr_results.pages)} pages du document {input_filename}\n"
        )

        if len(ocr_results.paragraphs) > 0:
            print(
                f"----#{len(ocr_results.paragraphs)} paragraphes détectés dans le document {input_filename}----\n"
            )

            with open(txt_file, "w", encoding='utf-8') as f:
                for paragraph in ocr_results.paragraphs:
                    f.write(paragraph.content + "\n" + "\n")

            f.close()
            print(f"✓ {datetime.now():%d/%m/%Y %H:%M:%S}: Fichier .txt du PDF {input_filename} créé.\n")

        # Generate OCR overlay layer
        print(f"{datetime.now():%d/%m/%Y %H:%M:%S}: Création du fichier...\n")
        output = PdfWriter()
        default_font = "Al Nile"
        for page_id, page in enumerate(ocr_results.pages):
            ocr_overlay = io.BytesIO()

            # Calculate overlay PDF page size
            if image_pages[page_id].height > image_pages[page_id].width:
                page_scale = float(image_pages[page_id].height) / pagesizes.letter[1]
            else:
                page_scale = float(image_pages[page_id].width) / pagesizes.letter[1]

            page_width = float(image_pages[page_id].width) / page_scale
            page_height = float(image_pages[page_id].height) / page_scale

            scale = (page_width / page.width + page_height / page.height) / 2.0
            pdf_canvas = canvas.Canvas(ocr_overlay, pagesize=(page_width, page_height))

            # Add image into PDF page
            pdf_canvas.drawInlineImage(image_pages[page_id], 0, 0, width=page_width, height=page_height, preserveAspectRatio=True)
            
            text = pdf_canvas.beginText()
            # Set text rendering mode to invisible
            text.setTextRenderMode(3)
            for word in page.words:
                # Calculate optimal font size
                desired_text_width = max(dist(word.polygon[0], word.polygon[1]), dist(word.polygon[3], word.polygon[2])) * scale
                desired_text_height = max(dist(word.polygon[1], word.polygon[2]), dist(word.polygon[0], word.polygon[3])) * scale
                font_size = desired_text_height
                actual_text_width = pdf_canvas.stringWidth(word.content, default_font, font_size)

                # Calculate text rotation angle
                text_angle = math.atan2((word.polygon[1].y - word.polygon[0].y + word.polygon[2].y - word.polygon[3].y) / 2.0, 
                                        (word.polygon[1].x - word.polygon[0].x + word.polygon[2].x - word.polygon[3].x) / 2.0)
                text.setFont(default_font, font_size)
                text.setTextTransform(math.cos(text_angle), -math.sin(text_angle), math.sin(text_angle), math.cos(text_angle), word.polygon[3].x * scale, page_height - word.polygon[3].y * scale)
                text.setHorizScale(desired_text_width / actual_text_width * 100)
                
                # Reshaping ar chars
                if confidence == 0:
                    text.textOut(get_display(reshape(word.content + " ")))
                if confidence == 1:
                    if word.confidence >= 0.8:
                        text.textOut(get_display(reshape(word.content + " ")))
                    if 0.5 < word.confidence < 0.8:
                        yellowtransparent = Color(255, 211, 0, alpha=0.5)
                        pdf_canvas.setFillColor(yellowtransparent)
                        pdf_canvas.rect(text._x, text._y, desired_text_width, desired_text_height, fill = 1, stroke = 0)
                        pdf_canvas.setFillColor(black)
                        text.textOut(get_display(reshape(word.content + " ")))
                    if word.confidence <= 0.5:
                        redtransparent = Color( 100, 0, 0, alpha=0.5)
                        pdf_canvas.setFillColor(redtransparent)
                        pdf_canvas.rect(text._x, text._y, desired_text_width, desired_text_height, fill = 1, stroke = 0)
                        pdf_canvas.setFillColor(black)
                        text.textOut(get_display(reshape(word.content + " ")))
                        

            pdf_canvas.drawText(text)
            pdf_canvas.save()

            # Move to the beginning of the buffer
            ocr_overlay.seek(0)

            # Create a new PDF page
            new_pdf_page = PdfReader(ocr_overlay)
            output.add_page(new_pdf_page.pages[0])

        # Save output searchable PDF file
        with open(output_file, "wb") as outputStream:
            output.write(outputStream)
            f.close()
        print(f"✓ {datetime.now():%d/%m/%Y %H:%M:%S}: PDF créé: {output_filename}_ocr.pdf\n")
        counter += 1

    else:
        print(f"❐ {datetime.now():%d/%m/%Y %H:%M:%S}: Le fichier {output_filename}._ocr.pdf existe déjà\n")
        counter += 1
    
    return

#Batch module
def batch(input_path, output_path):
    global unlock, total
    if is_on:
        print(f"Clé Azure: {key}\n")
        print(f"Endpoint: {endpoint}\n")
        print(f"Dossier d'entrée: {input_path}\n")
        print(f"Dossier de sortie: {output_path}\n")
        list1 = []
        for root, dirs, files in os.walk(input_path):
            for file in files:
                if file.endswith(".pdf"):
                    list1.append(os.path.join(root, file))
                    total = len(list1)
        print(f"Nombre total de fichiers PDF : {total}\n")
        for input_file in Path(input_path).rglob("*.pdf"):
            ocr(input_path, output_path, input_file)

    elif not is_on:
        print(f"✕ {datetime.now():%d/%m/%Y %H:%M:%S}: L'opération a été interrompue\n")

    
    print(f"{datetime.now():%d/%m/%Y %H:%M:%S}: Fin de l'opération\n")
    unlock = True

    if corrupt > 0:
        print("Les fichiers suivants n'ont pu être OCRisés: \n")
        print("---Liste des fichiers vraisemblablement compromis---\n")
        for name in corrupt_names:
            print(f"{name}\n")

    #End credits = Directed by Robert B. Weide
    displayNotification("Fin", "Tâches accomplies", "", "Hero")
    tk.messagebox.showinfo("Fin de l'opération", "Toutes les tâches ont été finalisées")


def get_dpi():
    screen = Tk()
    current_dpi = screen.winfo_fpixels('1i')
    screen.destroy()
    print(current_dpi)
    return current_dpi

if __name__ == "__main__":
    root = ttkb.Window(themename="new")
    root.geometry("872x594")
    root.title("Azure Document Intelligence OCR")
    app = App(root)
    root.resizable(False, False)
    root.mainloop()
