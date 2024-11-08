import logging
import os
import sys
from tkinter import *
import tkinter as tk
from tkinter import scrolledtext
from tkinter.ttk import Progressbar
from tkinter.ttk import Combobox

from tkinter import messagebox
from pdf_util import ProgressHandler


def resource_path(relative_path):
    """
    Get absolute path to resource, works for dev and for PyInstaller.
    pyinstaller unpacks your data into a temporary folder, and stores this
    directory path in the _MEIPASS2 environment variable.
    mode, I use this:"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath("..")

    return os.path.normpath(os.path.join(base_path, relative_path))


class GUICore:
    def __init__(self):
        self.root = Tk()
        self.root.resizable(False, False)
        self.root.title('PDF Creator v1.0.0')
        # Upon GUI initialization create local variables to hold:
        # - current working directory,
        # - name of the output file,
        # - name of the metadata *csv file.
        self.CWD = os.path.normpath(os.getcwd())
        self.OUTPUT_FILENAME = 'combined'
        self.TLF_METADATA_NAME = 'tlfmetadata'
        # Insert Atorus logo.
        img1 = PhotoImage(file=resource_path(os.path.join('../assets/images', 'atorus_logo.png')))
        banner1 = Label(self.root, image=img1)
        banner1.image = img1
        banner1.place(x=40, y=4)

        # Insert PDF Utility logo.
        img2 = PhotoImage(file=resource_path(os.path.join('../assets/images', 'pdf_utility_logo.png'))).subsample(2, 2)
        banner2 = Label(self.root, image=img2)
        banner2.image = img2
        banner2.place(x=190, y=12)
        # Insert Python logo.
        img3 = PhotoImage(file=resource_path(os.path.join('../assets/images', 'python_logo.png'))).subsample(2, 2)
        banner3 = Label(self.root, image=img3)
        banner3.image = img3
        banner3.place(x=405, y=12)


        # Labels.
        self.lbl2 = Label(self.root, text='Choose Folder:', font='Ubuntu, 10')
        self.lbl2.place(x=40, y=50)

        self.lbl3 = Label(self.root, text='TLF Metadata:', font='Ubuntu, 10')
        self.lbl3.place(x=40, y=85)

        self.lbl4 = Label(self.root, text='Output As:', font='Ubuntu, 10')
        self.lbl4.place(x=40, y=115)

        # Log Text field.
        self.txt1 = scrolledtext.ScrolledText(self.root, wrap=WORD, font='Ubuntu, 8', fg='black', state='disabled')
        self.txt1.place(x=40, y=250, width=625, height=250)

        # Progress Bar widget
        # Create ProgressBar after button 'GO!' is pressed.
        self.pb1 = Progressbar(self.root)

        # Create textLogger
        self.text_handler = ProgressHandler(self.txt1, self.pb1)
        # Add the handler to logger
        self.logger = logging.getLogger()
        self.logger.addHandler(self.text_handler)

        # For folder path entry and display.
        self.entry_var1 = StringVar()
        self.entr1 = Entry(self.root, textvariable=self.entry_var1)
        self.entr1.config(state='normal')
        self.entr1.place(x=150, y=51, width=400, height=25)
        # Assign by-default value to CWD - path to PY executable.
        self.entry_var1.set(os.path.normpath(self.CWD))

        # For TLF Metadata path.
        self.entry_var2 = StringVar()
        self.entr2 = Entry(self.root, textvariable=self.entry_var2)
        self.entr2.config(state='normal')
        self.entr2.place(x=150, y=86, width=400, height=25)
        # Set default output filename.
        self.entry_var2.set(os.path.normpath(os.path.join(self.CWD, self.TLF_METADATA_NAME + '.xlsx')))

        # For output filename entry and display.
        self.entry_var3 = StringVar()
        self.entr3 = Entry(self.root, textvariable=self.entry_var3)
        self.entr3.config(state='normal')
        self.entr3.place(x=150, y=115, width=370, height=25)
        # Set default output filename.
        self.entry_var3.set(self.OUTPUT_FILENAME)

        # A hint saying extention will be PDF.
        self.lbl5 = Label(self.root, text='.pdf', font=('Ubuntu, 10'))
        self.lbl5.place(x=525, y=115)

        # For browsing path to folder.
        self.btn1 = Button(self.root, text='Browse', font=('Ubuntu, 9'), width=11)
        self.btn1.place(x=580, y=50)

        # For browsing path to TLF Metadata.
        self.btn2 = Button(self.root, text='Browse', font=('Ubuntu, 9'), width=11)
        self.btn2.place(x=580, y=85)

        # Password set and store
        self.entry_var5 = StringVar()
        def set_pass():
            if self.pas_check_var:
                self.lbl6 = Label(self.root, text='Password:', font='Ubuntu, 10')
                self.lbl6.place(x=370, y=210)
                # For password entry and display.

                self.entr5 = Entry(self.root, textvariable=self.entry_var5, show='*')
                self.entr5.config(state='normal')
                self.entr5.place(x=450, y=210, width=210, height=25)
            else:
                pass

        self.pas_check_var = BooleanVar()
        self.pas_check_var.set(0)
        self.pass_check = Checkbutton(self.root, text='Set password', variable=self.pas_check_var, onvalue=1,
                                      offvalue=0, font=('Ubuntu, 9'), justify='left', command=set_pass)
        self.pass_check.place(x=370, y=180)

        self.final_run_var = BooleanVar()
        self.final_run_var.set(0)
        self.final_run = Checkbutton(self.root, text='Final run', variable=self.final_run_var, onvalue=1,
                                      offvalue=0, font=('Ubuntu, 9'), justify='left')
        self.final_run.place(x=370, y=145)

        # Combobox for select font
        self.lbl6 = Label(self.root, text='Use font for TOC:', font='Ubuntu, 10')
        self.lbl6.place(x=40, y=145)

        self.box_value = StringVar()
        self.locationBox = Combobox(self.root, textvariable=self.box_value, state = 'readonly', width=25)
        self.locationBox.place(x=150, y=145)
        self.locationBox.bind("<<ComboboxSelected>>", self.get_selected_value(self))
        self.locationBox['values'] = ('DejaVuSansMono (152U)', 'PT_Mono (152U)', 'Monospace (152U)', 'DroidSansMono (152P)', 'FiraMono (152P)',
                                      'JetBrainsMono (152P)', 'LiberationMono (152P)', 'NotoMono (152P)', 'CamingoCode (166P)',
                                      'Lekton (182P)',  'EversonMono (159P)', 'Monoid (160P)', 'VictorMono (167P)')
        self.locationBox.current(0)

        # Frame for bottons for additional options for bookmarks.
        self.frm1 = LabelFrame(self.root, relief='solid', bd='0.5')
        self.frm1.place(x=40, y=180, width=310, height=60)

        self.lbl_additional_opts = Label(self.root, text='Additional Bookmark Options:', font=('Ubuntu, 8 italic'),
                                         anchor='w', justify='left')
        self.lbl_additional_opts.place(x=45, y=185)

        # Checkbutton for adding optional 'Population' to bookmark.
        self.check_var2 = BooleanVar()
        self.check_var2.set(1)
        self.check2 = Checkbutton(self.root, text='Include Population', variable=self.check_var2, onvalue=1,
                                  offvalue=0, font=('Ubuntu, 8'))
        self.check2.place(x=45, y=200)


        self.lbl5 = Label(self.root, text='Title Separator:', font='Ubuntu, 8')
        self.lbl5.place(x=170, y=200)
        self.entry_var4 = StringVar()
        self.entr4 = Entry(self.root, textvariable=self.entry_var4, font='Ubuntu, 7')
        self.entr4.config(state='normal')
        self.entr4.xview(END)
        self.entry_var4.set('-')
        self.entr4.place(x=250, y=200, width=45, height=25)


        # GO button.
        self.btn_go = Button(self.root, text='GO!', font=('Ubuntu, 10'))
        self.btn_go.place(x=580, y=115, width=88)


    @property
    def add_population(self):
        """
        :return: Boolean value of checkbox asking to add population to bookmarks or not.
        """
        return bool(self.check_var2.get())
    @property
    def title_separator(self):
        """
        :return: --> str. Character used to separate title and title2. If not specified - single space is used.
        """
        if self.entry_var4:
            if self.entry_var4.get() == '':
                sep = ' '
            else:
                sep = ' ' + self.entry_var4.get() + ' '
        else:
            sep = ''
        return sep
    def get_selected_value(self, event):
        return self.locationBox.get()

    # Function to update config and link function to button.
    def link_btn_to_command(self, btn, command):
        btn.configure(command=command)
        self.root.update_idletasks()

    # Change background color.
    def change_bg_color(self, list, color):
        self.root.configure(bg=color)
        for item in list:
            try:
                item.configure(bg=color)
            except TclError:
                pass

    # Get output name from entry filed
    def get_output_as(self, suffix=''):
        return self.entry_var3.get() + suffix + '.pdf'