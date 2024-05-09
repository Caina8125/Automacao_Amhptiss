import customtkinter
from tkinter import *
from itertools import count, cycle
from PIL import Image, ImageTk

class ImageLabel(customtkinter.CTkLabel):
   
    def load(self, im):
        if isinstance(im, str):
            im = Image.open(im)
        frames = []

        try:
            for i in count(1):
                frames.append(ImageTk.PhotoImage(im.copy()))
                im.seek(i)
        except EOFError:
            pass
        self.frames = cycle(frames)

        try:
            self.delay = im.info['duration']
        except:
            self.delay = 80

        if len(frames) == 1:
            self.config(image=next(self.frames))
        else:
            self.next_frame()

    def unload(self):
        self.config(image=None)
        self.frames = None

    def next_frame(self):
        if self.frames:
            self.configure(image=next(self.frames))
            self.after(self.delay, self.next_frame)

    def iniciarGif(self,janela,texto):
        self.info = customtkinter.CTkLabel(janela,text=texto, text_color="#274360",font=("Arial",20))
        self.info.pack(padx=10, pady=10)

        self.lbl = ImageLabel(janela, text="",bg_color="transparent", fg_color="transparent")
        self.lbl.pack(padx=10, pady=10)
        self.lbl.load(r"Infra\Arquivos\loader2.gif")

    def ocultarGif(self):
        # self.unload(self)
        self.info.pack_forget()
        self.lbl.pack_forget()