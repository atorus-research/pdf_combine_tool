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

    def combine_pdfs_simple(self, pdf_files, output_name, use_password=False, password=None):
        """Combine PDFs without TOC or bookmarks"""
        try:
            with fitz.open() as result:
                for pdf in pdf_files:
                    with fitz.open(pdf) as mfile:
                        result.insert_pdf(mfile)

                if use_password:
                    result.save(
                        output_name,
                        encryption=fitz.PDF_ENCRYPT_AES_256,
                        owner_pw=password,
                        garbage=4,
                        deflate=True
                    )
                else:
                    result.save(output_name, garbage=4, deflate=True)

            return True
        except Exception as e:
            self.gui.logger.error(f'ERROR: Failed to combine PDFs: {str(e)}')
            return False

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

    def close_word_proc(self, proc_tuple=("word", "winword", "WINWORD", "splwow64.exe"), silent=False):
        """
        Check if word application or print service are running and kill them
        """
        for proc in psutil.process_iter():
            if any(procstr in proc.name() for procstr in proc_tuple):
                if not silent:
                    result = messagebox.askquestion(
                        title="Word process running",
                        message='All Word related processes should be closed before run.\nClose all Word processes?'
                    )
                    if result == 'yes':
                        try:
                            proc.kill()
                        except:
                            pass
                else:
                    # In silent mode, kill process without asking
                    try:
                        proc.kill()
                    except:
                        pass

    # TODO: TO_THINK: run  with multithreads, parallelization?
    def rtf_file_to_pdf(self, file_name: str, input_dir: str, output_dir: str, pause_time: float) -> None:
        """Convert RTF to PDF with improved process handling"""
        word = None
        max_retries = 3
        retry_count = 0

        while retry_count < max_retries:
            try:
                word = win32com.client.gencache.EnsureDispatch('Word.Application')
                word.Visible = False

                in_file = os.path.normpath(os.path.join(input_dir, file_name))
                output_file = os.path.splitext(file_name)[0]
                out_file = os.path.normpath(os.path.join(output_dir, output_file + '.pdf'))

                if os.path.isfile(out_file):
                    self.gui.logger.warning(f'{file_name} already exists as PDF')
                    return

                doc = word.Documents.Open(in_file, False, False, True)
                doc.SaveAs(out_file, FileFormat=17)  # wdFormatPDF = 17
                doc.Close(SaveChanges=0)  # wdDoNotSaveChanges = 0
                time.sleep(pause_time)
                self.gui.logger.warning(f'{file_name} has been converted to PDF')
                return

            except Exception as e:
                retry_count += 1
                self.gui.logger.warning(f'Attempt {retry_count} failed for {file_name}: {str(e)}')
                try:
                    if doc:
                        doc.Close(SaveChanges=0)
                except:
                    pass

            finally:
                try:
                    if word:
                        word.Quit()
                except:
                    pass

                # Force cleanup of any hanging processes
                self.close_word_proc(silent=True)

        if retry_count >= max_retries:
            error_msg = f'Failed to convert {file_name} after {max_retries} attempts'
            self.gui.logger.error(error_msg)
            if self.gui.final_run_var.get() == 1:
                messagebox.showerror('File convert error', error_msg)
                os.abort()

    def convert_to_pdf(self, in_list: list, rtf_folder_dir: str, pdf_folder_dir: str) -> None:
        """
        Convert RTF files to separate PDF files.
        :param in_list: List of outputs to convert.
        :param rtf_folder_dir: Source directory.
        :param pdf_folder_dir: Where to put converted PDFs.

        :return: N/A
        """
        # ask user to close all word processes for avoid freeze during conversion
        self.close_word_proc()

        in_folder_pdf = {os.path.join(elem)[:-4] + '.rtf' for elem in os.listdir(pdf_folder_dir)
                         if pathlib.Path(elem).suffix == '.pdf'}
        files_to_convert = [x for x in in_list if x not in in_folder_pdf]

        if files_to_convert:
            if self.gui.final_run_var.get() == 1:
                for file in files_to_convert:
                    self.rtf_file_to_pdf(file_name=file, input_dir=rtf_folder_dir,
                                         output_dir=pdf_folder_dir, pause_time=0.5)
            else:
                for file in files_to_convert:
                    self.rtf_file_to_pdf(file_name=file, input_dir=rtf_folder_dir,
                                         output_dir=pdf_folder_dir, pause_time=0.5)
        else:
            self.gui.logger.warning('No new files to convert')

    def add_bmk_to_file(self, input_dir: str, meta_data_file: str, title_sep: str, add_popul: bool = True) -> None:
        """Add bookmarks to PDF files based on metadata"""
        df = pd.read_csv(meta_data_file)
        df = df.dropna(how='all')
        df['Filename'] = df['OutputName'].str.replace('-', '_')
        df['Filename'] = df['Filename'].str.replace('.', '_')
        df['Filename'] = df['Filename'] + ".rtf"
        df['FilenamePDF'] = df['Filename'].str[:-4] + '.pdf'
        df['FilenamePDF'] = df['FilenamePDF'].apply(lambda x: os.path.join(input_dir, x))

        if add_popul:
            df['Bookmark'] = df['Title3'] + str(title_sep) + df['Title4'] + str(title_sep) + df['Title5']
        else:
            df['Bookmark'] = df['Title3'] + str(title_sep) + df['Title4']

        # Filter to only existing files
        existing_files = df[df['FilenamePDF'].apply(os.path.exists)]

        if existing_files.empty:
            self.gui.logger.warning("No PDF files found matching metadata entries")
            return

        # Process each file individually
        for _, row in existing_files.iterrows():
            file = row['FilenamePDF']
            bmk_txt = row['Bookmark']

            self.gui.logger.warning("Add bookmark to file " + str(file))
            self.gui.logger.warning("Bookmark to add: " + str(bmk_txt))

            try:
                temp_file = file + ".tmp"
                doc = fitz.open(file)
                new_doc = fitz.open()
                new_doc.insert_pdf(doc)
                new_doc.set_toc([[1, bmk_txt, 1]])
                new_doc.save(temp_file, garbage=4, deflate=True)
                new_doc.close()
                doc.close()

                try:
                    os.replace(temp_file, file)
                except PermissionError:
                    os.remove(file)
                    os.rename(temp_file, file)

            except Exception as e:
                self.gui.logger.error(f"Error processing file {file}: {str(e)}")
                if os.path.exists(temp_file):
                    try:
                        os.remove(temp_file)
                    except:
                        pass

    def go_combine_selected_pdf(self, dir, meta_data_, out_name, title_sep: str, add_popul: bool = True,
                                prot_fl: bool = False):
        df = pd.read_csv(meta_data_)
        df = df.dropna(how='all')
        df['Filename'] = df['OutputName'].str.replace('-', '_')
        df['Filename'] = df['Filename'].str.replace('.', '_')
        df['Filename'] = df['Filename'] + ".rtf"

        df['FilenamePDF'] = df['Filename'].str[:-4] + '.pdf'
        existing_files = df[df['FilenamePDF'].apply(lambda x: os.path.exists(os.path.join(dir, x)))]

        if existing_files.empty:
            self.gui.logger.warning("No PDF files found to combine")
            return False

        if add_popul:
            existing_files.loc[:, 'Bookmark'] = existing_files['Title3'] + str(title_sep) + existing_files[
                'Title4'] + str(title_sep) + existing_files['Title5']
        else:
            existing_files.loc[:, 'Bookmark'] = existing_files['Title3'] + str(title_sep) + existing_files['Title4']

        order_dict = dict(zip(existing_files['Order'], existing_files['FilenamePDF']))
        pdf_files_ = tuple(os.path.join(dir, v) for k, v in dict(sorted(order_dict.items())).items())

        if pdf_files_:
            # Use self.gui.CWD to get the correct working directory
            out_path = os.path.join(dir, '..', out_name)
            self._fitz_combine(pdf_files_, out_path, prot_fl)
            self.gui.logger.warning(f'\nINFO: Job finished! {len(pdf_files_)} files were combined into {out_name}')
            self.gui.logger.warning(f'\nINFO: {out_name} is saved in {os.path.dirname(out_path)}')
            return True
        return False

    def _fitz_combine(self, pdf_files_, output_name, prot_fl=False):
        """Combine PDFs using PyMuPDF (fitz)"""
        result = fitz.open()
        general_toc = []
        current_page = 1

        for pdf in pdf_files_:
            with fitz.open(pdf) as mfile:
                pages = len(result)
                result.insert_pdf(mfile)
                tmp_toc = mfile.get_toc(simple=True)
                if tmp_toc:
                    for t in tmp_toc:
                        t[2] += pages
                    general_toc.extend(tmp_toc)

        if general_toc:
            result.set_toc(general_toc)

        if self.gui.pas_check_var.get() and self.gui.entry_var5.get():
            result.save(output_name,
                        encryption=fitz.PDF_ENCRYPT_AES_256,
                        owner_pw=self.gui.entry_var5.get(),
                        garbage=4,
                        deflate=True)
        else:
            result.save(output_name, garbage=4, deflate=True)
        result.close()
        return True


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


