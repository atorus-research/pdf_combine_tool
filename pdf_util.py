import logging
import time
from tkinter.filedialog import *
from tkinter import messagebox


import win32com.client
import pandas as pd
import os
import os.path
import psutil
import pathlib
import shutil
import math
import re


import fitz

import itertools


# pd.options.mode.chained_assignment = None  # default='warn'
class PDFUtility:
    def __init__(self, gui):
        global CWD
        self.gui = gui
        CWD = self.gui.entry_var1.get()
        self.meta_source = {}

    # Browse the current working directory for files to combine.
    # Pass the path to CWD.
    def select_folder(self):
        """
        Function for select current working directory path

        Args:
            self

        Returns:
            None
        Raises:
            None

        """
        self.gui.entry_var1.set(os.path.normpath(askdirectory()))
        self.assign_cwd()

    def assign_cwd(self):
        """
        Function for assign selected current working folder path to value

        Returns:
            None

        """
        CWD_dir = self.gui.entry_var1.get()
        # Set CWD_dir.
        if CWD_dir != '':
            os.chdir(CWD_dir)
            self.gui.logger.warning('INFO: Directory set to ' + str(CWD_dir))

    # Create subdirectory for holding converted separate PDFs.
    def mkdir(self, dirname: str = "_PDF"):
        """
        Create folder "_PDF" for store converted TLF files

        Args:
            dirname (string, required): Name for folder

        Returns:
            None

        Raises:
            Log message that folder already exists
        """

        try:
            # Create target Directory
            os.mkdir(dirname)
            self.gui.logger.warning('INFO: Directory ' + str(dirname) + ' was created')
        except:
            self.gui.logger.warning('INFO: Directory ' + str(dirname) + ' already exists')
            pass

    # Browse the path to TLF metadata.
    def select_metadata(self):
        """
        Function for select metadate files: shows only *xlsx and * sas7bdat files

        Returns:
            None
        Raises:
            None
        """
        self.gui.entry_var2.set(
            os.path.normpath(askopenfilename(filetypes=[("Meta-data", "*.csv"), ("Meta-data", "*.sas7bdat")])))
        self.assign_meta()

    def assign_meta(self):
        """
        Function for assign  selected path to metadata file into values

        Returns:
            None
        Raises:
            Log message in case when metadata file wasn't selected

        """
        global METADATA
        if self.gui.entry_var2.get() != '.':
            METADATA = self.gui.entry_var2.get()
            self.gui.logger.warning('INFO: Selected metadata file: ' + str(METADATA))
        else:
            self.gui.logger.warning('WARNING: Metadata file not selected')
            messagebox.showwarning(title="Warning", message="Please select metatada file")
        self.meta_source = self.meta_data_to_dict(METADATA, title_sep="*", add_popul=False)


    @staticmethod
    def get_event_number(path_to_metadata: str):
        """
                Count number of event for log based of number of files from metadata file
        :param path_to_metadata: full path to meta data excel file.
        :return: int; number of rows from metadata file

        Args:
            path_to_metadata (str): path to metadata file

        Returns:
            r (int): number of events calculated based of metadata file length.
        Raises:
            MessageBox with invite user to check metadata file for empty file (no files to convert)

        """
        df1 = pd.read_csv(path_to_metadata)
        df1 = df1.dropna(axis=1, how='all')
        df1 = df1.dropna(axis=0, how='all')

        r, c = df1.shape
        if r:
            return r
        else:
            messagebox.showwarning(title="Warning", message="Please selected metadata file is empty. Please, check"
                                                            "metadata file. ")
            return None



    @staticmethod
    def get_tlf_list(path_to_metadata: str):
        """
        Count number of tlf files and get list of tlf files

        Args:
            path_to_metadata (str): path to metadata file

        Returns:
            file_to_convert (list): list of files need to convert and/or combine based on metadata file
            tlf_count (int): total number of files for processing

        Raises:
            None
        """
        df1 = pd.read_csv(path_to_metadata)
        df1 = df1.dropna(axis=1, how='all')
        df1 = df1.dropna(axis=0, how='all')
        df1['Filename'] = df1['OutputName'].str.replace('-', '_')
        df1['Filename'] = df1['Filename'].str.replace('.', '_')
        df1['Filename'] = df1['Filename'] + ".rtf"

        file_to_convert = df1['Filename'].to_list()
        tlf_count = len(file_to_convert)

        return file_to_convert, tlf_count

    @staticmethod
    def meta_data_to_dict(meta_data_file, title_sep: str, add_popul: bool = True):
        """
        Convert input metadata excel file into PY dictionary

        :param meta_data_file: full path to meta data excel file.
        :return: dict; selected from input metadata columns as a PY dict.

        """
        df = pd.read_csv(meta_data_file)
        df = df.dropna(how='all')
        meta_df = df.copy()

        r, c = meta_df.shape
        output_id = []
        output_title = []
        for i in range(0, r):
            output_id.append(str(meta_df.at[i, 'Title3']).capitalize().strip())
            if add_popul:
                output_title.append((str(meta_df.at[i, 'Title4']).capitalize().strip() + str(title_sep) +
                                     str(meta_df.at[i, 'Title5']).strip()))
            else:
                output_title.append(str(meta_df.at[i, 'Title4']).strip())

        return dict(zip(output_id, output_title))

    @staticmethod
    def close_word_proc(proc_tuple=("word", "winword", "WINWORD", "splwow64.exe"), silent=False):
        """
         Check if word application or print service are running and kill them to avoid freeze tool
         :param proc_tuple: list of process to check, default values word and print service
         :return: None
         """
        for proc in psutil.process_iter():
            if any(procstr in proc.name() for procstr in proc_tuple):
                if not silent:
                    result = messagebox.askquestion(title="Word process running",
                                                    message='All Word related processes should be closed before run.' + \
                                                            "\nClose all Word processes?")

                if result or silent:
                    proc.kill()

    # TODO: TO_THINK: run  with multithreads, parallelization?
    def rtf_file_to_pdf(self, file_name: str, input_dir: str, output_dir: str, pause_time: float) -> None:
        word = None  # declare variable 'word'.
        wdFormatPDF = 17
        wdDoNotSaveChanges = 0

        word = win32com.client.gencache.EnsureDispatch('Word.Application')

        in_file = os.path.normpath(os.path.join(input_dir, file_name))
        output_file = os.path.splitext(file_name)[0]
        out_file = os.path.normpath(os.path.join(output_dir, output_file + '.pdf'))
        self.gui.logger.warning('Converting ' + str(file_name) + '...')

        if os.path.isfile(out_file):
            self.gui.logger.warning(str(file_name) + ' has been detected and do not need to convert to PDF.')

        else:
            if self.gui.final_run_var.get() == 1: #Raise Error and abort - Final run mood, no file
                try:
                    doc = word.Documents.Open(in_file, False, False, True)  # 'True' as a 3d param tell to open in ReadOnly.
                    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
                    doc.Close(SaveChanges=wdDoNotSaveChanges)
                    time.sleep(pause_time)
                    self.gui.logger.warning(str(file_name) + ' has been converted to PDF.')
                    # out_file.close()
                except Exception as e:
                    print(e)
                    self.gui.logger.error('ERROR: Error handle while converting ' + str(file_name) + ' file.')
                    messagebox.showerror(title='File convert error',
                                         message='ERROR: Error handle while converting ' + str(file_name) +
                                                                              '. Check metadata file and '
                                                                              'TLF  file.', default='ok')
                    os.abort()
            else:
                try:
                    doc = word.Documents.Open(in_file, False, False, True)  # 'True' as a 3d param tell to open in ReadOnly.
                    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
                    doc.Close(SaveChanges=wdDoNotSaveChanges)
                    time.sleep(pause_time)
                    self.gui.logger.warning(str(file_name) + ' has been converted to PDF.')

                except Exception as e:
                    self.gui.logger.error("Sorry, we couldn't find your file. Was it moved, renamed or deleted? "
                                          + str(file_name) + ' file.')
                    pass


    def convert_to_pdf(self, in_list: list, rtf_folder_dir: str, pdf_folder_dir: str) -> None:
        """
        Convert RTF files to separate PDF files.
        :param in_list: List of outputs to convert.
        :param rtf_folder_dir: Source directory.
        :param pdf_folder_dir: Where to put converted PDFs.

        :return: N/A
        """
        # ask user to close all word processes for avoid freeze during convertation
        self.close_word_proc()
        #check if all files from pdf_to_keep present into folder -> convert while not true
        in_folder_pdf = tuple(os.path.join(elem)[:-4]+'.rtf' for elem in os.listdir(pdf_folder_dir) if
                       pathlib.Path(elem).suffix == '.pdf')
        s = set(in_folder_pdf)
        _tm = tuple(x for x in in_list if x not in s)

        if _tm:
            if self.gui.final_run_var.get() == 1:
                while len(_tm) != 0:
                    for file in _tm:
                        self.rtf_file_to_pdf(file_name=file, input_dir=rtf_folder_dir,
                                             output_dir=pdf_folder_dir, pause_time=0.5)
                        in_folder_pdf = tuple(os.path.join(elem)[:-4] + '.rtf' for elem in os.listdir(pdf_folder_dir) if
                                              pathlib.Path(elem).suffix == '.pdf')
                        s = set(in_folder_pdf)
                        _tm = tuple(x for x in in_list if x not in s)
            else:
                for file in _tm:
                    self.rtf_file_to_pdf(file_name=file, input_dir=rtf_folder_dir,
                                         output_dir=pdf_folder_dir, pause_time=0.5)
                    in_folder_pdf = tuple(os.path.join(elem)[:-4] + '.rtf' for elem in os.listdir(pdf_folder_dir) if
                                          pathlib.Path(elem).suffix == '.pdf')
                    s = set(in_folder_pdf)
                    _tm = tuple(x for x in in_list if x not in s)
        else:
            # Reset Progress Bar
            self.gui.pb1['value'] = 0

    def add_bmk_to_file(self, input_dir: str, meta_data_file: str, title_sep: str, add_popul: bool = True) -> None:

        df = pd.read_csv(meta_data_file)
        df = df.dropna(how='all')
        df['Filename'] = df['OutputName'].str.replace('-', '_')
        df['Filename'] = df['Filename'].str.replace('.', '_')
        df['Filename'] = df['Filename'] + ".rtf"
        df['FilenamePDF'] = input_dir + '\\' + df.Filename.str.slice(0, -4) + '.pdf'
        if add_popul:
            df['Bookmark'] = df['Title3'] + str(title_sep) + df['Title4'] + str(title_sep) + df['Title5']
        else:
            df['Bookmark'] = df['Title3'] + str(title_sep) + df['Title4']

        file_bmk_dict = dict(zip(df.FilenamePDF, df.Bookmark))

        for file, bmk_txt in file_bmk_dict.items():
            print("FINAL RUN MOOD: ", self.gui.final_run_var.get())
            if self.gui.final_run_var.get(): #Final run - all files exists according to metadata file
                self.gui.logger.warning("Add bookmark to file " + str(file))
                self.gui.logger.warning("Bookmark to add: " + str(bmk_txt))
                with fitz.open(file) as _tmpfile:
                    _tmpfile.set_toc([[1, bmk_txt, 1]])
                    _tmpfile.name = file
                    print(_tmpfile.can_save_incrementally())
                    _tmpfile.saveIncr()

            else: #temp run - not all files from metadata are into tfl's folder
                if os.path.exists(file): #file exists - need to add bookmark
                    self.gui.logger.warning("Add bookmark to file " + str(file))
                    self.gui.logger.warning("Bookmark to add: " + str(bmk_txt))
                    with fitz.open(file) as _tmpfile:
                        _tmpfile.set_toc([[1, bmk_txt, 1]])
                        _tmpfile.name = file
                        print(_tmpfile.can_save_incrementally())
                        _tmpfile.saveIncr()

                else: #file not exist - need to create it and add bookmark
                    self.gui.logger.warning("Create file: " + str(file))
                    bmk_txt = str(os.path.basename(file))[:-4] + "NO SUCH FILE IN TLF's FOLDER->Re-RUN to get bookmark"
                    self.gui.logger.warning("Bookmark to add_: " + str(bmk_txt))

                    doc = fitz.open()
                    page = doc.newPage()
                    where = fitz.Point(50, 100)
                    page.insertText(where, """NO SUCH FILE IN TLF's FOLDER""", fontsize=35)
                    doc.save(file)

                    with fitz.open(file) as _tmpfile:
                        _tmpfile.set_toc([[1, bmk_txt, 1]])
                        _tmpfile.name = file
                        print(_tmpfile.can_save_incrementally())
                        _tmpfile.saveIncr()



    def go_combine_selected_pdf(self, dir, meta_data_, out_name, title_sep: str, add_popul: bool = True,
                                prot_fl: bool =False):

        def fitz_combine(pdf_files, out_name_, prot_fl_=False):
            with fitz.open() as result:
                general_toc, tmp_toc = None, None

                for pdf in pdf_files:
                    with fitz.open(pdf) as mfile:
                        pages = len(result)
                        result.insert_pdf(mfile)
                        if not general_toc:
                            general_toc = mfile.get_toc(simple=True)
                        else:
                            tmp_toc = mfile.get_toc(simple=True)
                            for t in tmp_toc:  # increase toc2 page numbers
                                t[2] += pages  # by old len(doc1)
                            general_toc += tmp_toc

                # general_toc.sort()
                general_toc = list(k for k, _ in itertools.groupby(general_toc))
                result.set_toc(general_toc)
                if prot_fl_:
                    result.save(out_name_, pretty=True, garbage=4, deflate=True, encryption=4, user_pw="ewq321")
                else:
                    result.save(out_name_, pretty=True, garbage=4, deflate=True)


        # TODO: move to fucn (Dataframe to order list)
        df = pd.read_csv(meta_data_)
        df = df.dropna(how='all')
        df['Filename'] = df['OutputName'].str.replace('-', '_')
        df['Filename'] = df['Filename'].str.replace('.', '_')
        df['Filename'] = df['Filename'] + ".rtf"

        if add_popul:
            df['Bookmark'] = df['Title3'] + str(title_sep) + df['Title4'] + str(title_sep) + df['Title5']
        else:
            df['Bookmark'] = df['Title3'] + str(title_sep) + df['Title4']

        tmp_file_order_dict = dict(zip(df['Order'], df['Filename']))
        tmp_file_bookmark_dict = dict(zip(df['Filename'], df['Bookmark']))
        tmp_file_order_dict = {k: v[:-3] + 'pdf' for k, v in tmp_file_order_dict.items()}
        tmp_file_bookmark_dict = {k[:-3] + 'pdf': v for k, v in tmp_file_bookmark_dict.items()}

        order_dict = dict(sorted(tmp_file_order_dict.items()))
        pdf_files_ = tuple(os.path.join(dir, v) for k, v in dict(sorted(order_dict.items())).items())

        FILE_CONST = 400
        if len(pdf_files_) <= FILE_CONST:
            fitz_combine(pdf_files=pdf_files_, out_name_=out_name, prot_fl_=prot_fl)
        else:
            a_ = math.ceil(int(len(pdf_files_))/FILE_CONST)
            a_view = order_dict.items()
            a_list = list(a_view)
            st = 0
            end = FILE_CONST
            dct_lst = []
            for i in range(0, a_):
                dct_lst.append(f'a_{i}')
                exec("a_{} = a_list[st:end]".format(i))
                st = end
                end += FILE_CONST


            for elem in dct_lst:
                dict_1 = dict()
                exec("""for num, f_name in {0}:
                    dict_1.setdefault(num, []).append(f_name)""".format(elem))
                for key, value in dict_1.items():
                    dict_1[key] = value[0]



                pdf_files_ = tuple(os.path.join(dir, v) for k, v in dict(sorted(dict_1.items())).items())
                fitz_combine(pdf_files=pdf_files_, out_name_=elem+".pdf", prot_fl_=prot_fl)

            dct_lst = [elem+'.pdf' for elem in dct_lst]
            fitz_combine(pdf_files=dct_lst, out_name_=out_name, prot_fl_=prot_fl)
            for elem in dct_lst:
                try:
                    os.remove(elem)
                except Exception as e:
                    print(e)




class ProgressHandler(logging.Handler):
    """This class allows you to log to a Tkinter Text or ScrolledText widget"""

    def __init__(self, text, pb):
        # run the regular Handler __init__
        logging.Handler.__init__(self)
        # Store a reference to the Text it will log to
        self.text = text
        self.pb = pb
        self.formatter = logging.Formatter(fmt='[%(asctime)s] %(message)s\n', datefmt='%Y-%m-%d %H:%M:%S')


    def emit(self, record):
        msg = self.format(record)
        def append():
            self.text.configure(state='normal')
            self.text.insert("end", msg + '\n')
            self.text.configure(state='disabled')
            # Autoscroll to the bottom
            self.text.yview("end")
        #This is necessary because we can't modify the Text from other threads
        self.text.after(0, append)
        self.text.after(100, self.text.update())
        self.update_progress_bar()

    def update_progress_bar(self):
        self.pb['value'] += 1
        self.pb.after(100, self.pb.update())


