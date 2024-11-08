import logging
import os
import shutil
import sys
from tkinter import *
import tkinter as tk
from tkinter import scrolledtext
from tkinter.ttk import Progressbar
from tkinter.ttk import Combobox

from tkinter import messagebox
from src.pdf_util import ProgressHandler


def resource_path(relative_path):
    """Get absolute path to resource, works for dev and PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # Running in development mode
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

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
        img1 = PhotoImage(file=resource_path(os.path.join('assets', 'images', 'atorus_logo.png')))
        banner1 = Label(self.root, image=img1)
        banner1.image = img1
        banner1.place(x=40, y=4)

        # Insert PDF Utility logo.
        img2 = PhotoImage(file=resource_path(os.path.join('assets', 'images', 'pdf_utility_logo.png'))).subsample(2, 2)
        banner2 = Label(self.root, image=img2)
        banner2.image = img2
        banner2.place(x=190, y=12)
        # Insert Python logo.
        img3 = PhotoImage(file=resource_path(os.path.join('assets', 'images', 'python_logo.png'))).subsample(2, 2)
        banner3 = Label(self.root, image=img3)
        banner3.image = img3
        banner3.place(x=405, y=12)


        # Labels.
        self.lbl2 = Label(self.root, text='Choose Folder:', font='Ubuntu, 10')
        self.lbl2.place(x=40, y=50)

        self.lbl3 = Label(self.root, text='TLF Metadata:', font='Ubuntu, 10')
        self.lbl3.place(x=40, y=150)

        self.lbl4 = Label(self.root, text='Output As:', font='Ubuntu, 10')
        self.lbl4.place(x=40, y=115)

        # Log Text field.
        self.txt1 = scrolledtext.ScrolledText(self.root, wrap=WORD, font='Ubuntu, 8', fg='black', state='disabled')
        self.txt1.place(x=40, y=200, width=525, height=300)

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

        # Create password label and entry first (but don't place them)
        self.lbl6 = Label(self.root, text='Password:', font='Ubuntu, 10')
        self.entr5 = Entry(self.root, textvariable=self.entry_var5, show='*')
        self.entr5.config(state='normal')

        self.pas_check_var = BooleanVar()
        self.pas_check_var.set(0)
        self.pass_check = Checkbutton(self.root, text='Set password', variable=self.pas_check_var, onvalue=1,
                                      offvalue=0, font=('Ubuntu, 9'), justify='left', command=self.set_pass)
        self.pass_check.place(x=580, y=250)

        self.final_run_var = BooleanVar()
        self.final_run_var.set(0)
        self.final_run = Checkbutton(self.root, text='Final run', variable=self.final_run_var, onvalue=1,
                                      offvalue=0, font=('Ubuntu, 9'), justify='left')
        self.final_run.place(x=580, y=220)

        self.box_value = StringVar()
        self.locationBox = Combobox(self.root, textvariable=self.box_value, state = 'readonly', width=25)
        self.locationBox.place(x=150, y=185)
        self.locationBox.bind("<<ComboboxSelected>>", self.get_selected_value(self))
        self.locationBox['values'] = ('DejaVuSansMono (152U)', 'PT_Mono (152U)', 'Monospace (152U)', 'DroidSansMono (152P)', 'FiraMono (152P)',
                                      'JetBrainsMono (152P)', 'LiberationMono (152P)', 'NotoMono (152P)', 'CamingoCode (166P)',
                                      'Lekton (182P)',  'EversonMono (159P)', 'Monoid (160P)', 'VictorMono (167P)')
        self.locationBox.current(0)

        # Combobox for select font
        self.lbl7 = Label(self.root, text='Use font for TOC:', font='Ubuntu, 10')
        self.lbl7.place(x=40, y=145)

        # Frame for bottons for additional options for bookmarks.
        self.frm1 = LabelFrame(self.root, relief='solid', bd='0.5')
        self.frm1.place(x=40, y=220, width=310, height=60)

        self.lbl_additional_opts = Label(self.root, text='Additional Bookmark Options:', font=('Ubuntu, 8 italic'),
                                         anchor='w', justify='left')
        self.lbl_additional_opts.place(x=45, y=225)

        # Checkbutton for adding optional 'Population' to bookmark.
        self.check_var2 = BooleanVar()
        self.check_var2.set(1)
        self.check2 = Checkbutton(self.root, text='Include Population', variable=self.check_var2, onvalue=1,
                                  offvalue=0, font=('Ubuntu, 8'))
        self.check2.place(x=45, y=240)

        self.lbl5 = Label(self.root, text='Title Separator:', font='Ubuntu, 8')
        self.lbl5.place(x=170, y=240)
        self.entry_var4 = StringVar()
        self.entr4 = Entry(self.root, textvariable=self.entry_var4, font='Ubuntu, 7')
        self.entr4.config(state='normal')
        self.entr4.xview(END)
        self.entry_var4.set('-')
        self.entr4.place(x=250, y=240, width=45, height=25)

        # Add frame for metadata/TOC options
        self.meta_frame = LabelFrame(self.root, text='TOC Options', relief='solid', bd='0.5')
        self.meta_frame.place(x=40, y=85, width=625, height=55)

        # Add radio buttons for TOC options
        self.toc_var = StringVar()
        self.toc_var.set('no_toc')  # Default to no TOC

        self.rb_no_toc = Radiobutton(
            self.meta_frame,
            text='Combine PDFs without TOC',
            variable=self.toc_var,
            value='no_toc',
            command=self.update_toc_controls,
            font=('Ubuntu, 9')
        )
        self.rb_no_toc.place(x=10, y=5)

        self.rb_use_meta = Radiobutton(
            self.meta_frame,
            text='Use TLF Metadata',
            variable=self.toc_var,
            value='use_meta',
            command=self.update_toc_controls,
            font=('Ubuntu, 9')
        )
        self.rb_use_meta.place(x=200, y=5)

        # Add template button
        self.btn_template = Button(
            self.meta_frame,
            text='Create Metadata Using Template',
            command=self.open_metadata_template,
            font=('Ubuntu, 9')
        )
        self.btn_template.place(x=380, y=5)

        # Move existing metadata controls
        self.lbl3.place_forget()  # Hide original metadata label
        self.entry_var2.set('')  # Clear metadata path
        self.entr2.place_forget()  # Hide metadata entry
        self.btn2.place_forget()  # Hide original browse button

        # Initialize control states
        self.update_toc_controls()

        # GO button.
        self.btn_go = Button(self.root, text='GO!', font=('Ubuntu, 10'))
        self.btn_go.place(x=580, y=450, width=88)

    def set_pass(self):
        if self.pas_check_var.get():
            self.lbl6.place(x=580, y=280)
            self.entr5.place(x=645, y=280, width=70, height=25)
        else:
            # Hide password field and label
            self.lbl6.place_forget()
            self.entr5.place_forget()
            self.entry_var5.set('')  # Reset password value

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

    def update_toc_controls(self):
        """Update GUI controls based on TOC selection"""
        if self.toc_var.get() == 'no_toc':
            # Hide all TOC and metadata-related controls
            self.lbl3.place_forget()  # TLF Metadata label
            self.entr2.place_forget()  # TLF Metadata entry
            self.btn2.place_forget()  # TLF Metadata browse button
            self.locationBox.place_forget()  # Font selection combobox
            self.lbl7.place_forget()  # Use font for TOC label
            self.frm1.place_forget()  # Bookmark options frame
            self.entry_var2.set('')  # Clear metadata path

            # Hide Additional Bookmark Options
            self.lbl_additional_opts.place_forget()  # Additional Bookmark Options label
            self.check2.place_forget()  # Include Population checkbox
            self.lbl5.place_forget()  # Title Separator label
            self.entr4.place_forget()  # Title Separator entry

            # Expand log field to use available space
            # Move it up and increase height
            self.txt1.place(x=40, y=200, width=525, height=300)

        else:
            # Show metadata controls
            self.lbl3.place(x=40, y=150)
            self.entr2.place(x=150, y=150, width=400, height=25)
            self.btn2.place(x=580, y=150)

            # Show TOC controls
            self.locationBox.place(x=150, y=185)
            self.lbl7.place(x=40, y=185)

            # Show Bookmark frame and options
            self.frm1.place(x=40, y=220, width=310, height=60)
            self.lbl_additional_opts.place(x=45, y=225)
            self.check2.place(x=45, y=240)
            self.lbl5.place(x=170, y=240)
            self.entr4.place(x=250, y=240, width=45, height=25)

            # Shrink the log field
            self.txt1.place(x=40, y=300, width=525, height=200)

    def open_metadata_template(self):
        """Open the metadata template for editing"""
        template_path = resource_path(os.path.join('examples', 'metadata_example.csv'))
        if not os.path.exists(template_path):
            messagebox.showerror("Error", "Metadata template not found!")
            return

        # Create user's template if it doesn't exist
        user_template = os.path.join(self.CWD, 'metadata_template.csv')
        if not os.path.exists(user_template):
            shutil.copy2(template_path, user_template)

        try:
            os.startfile(user_template)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open template: {str(e)}")

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