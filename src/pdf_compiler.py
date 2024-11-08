from tkinter import messagebox
import fitz
import textwrap
import pandas as pd
import numpy as np
from fpdf import FPDF
import os
import shutil
import sys

from collections import namedtuple

from src.gui import resource_path

sys.setrecursionlimit(1500000)




#TODO: Check all "open" statement and replcase to "open with"

class PDFCompiler:
    def __init__(self, gui, util):

        self.gui = gui
        self.util = util
        self.pathToRTF = ''
        self.pathToPDF = ''
        self.CWD = ''

    def combine_pdfs(self):
        # Disable 'GO' button after pressed.
        self.gui.btn_go.configure(state='disabled')

        global font_folder, usage_font_c, METADATA

        usage_font_c = str(self.gui.box_value.get())[:-7]

        self.gui.logger.warning('INFO: Selected font for TOC: ' + str(usage_font_c) + '.')
        font_folder = resource_path(os.path.join('assets', 'fonts', str(usage_font_c) + '.ttf'))
        self.CWD = self.gui.entry_var1.get()
        self.pathToRTF = os.path.join(self.CWD)
        # Assign output filename.
        self.gui.OUTPUT_FILENAME = self.gui.get_output_as()
        self.util.mkdir(os.path.join(self.CWD, '_PDF'))
        self.pathToPDF = os.path.join(self.CWD, '_PDF')
        self.out_file_txt = os.path.normpath(os.path.join(self.pathToPDF, 'toc_file.txt'))
        self.pathToFile = str(self.pathToPDF)[:-4] + str(self.gui.OUTPUT_FILENAME)
        self.outFilePdfToc = self.pathToFile[:-4] + '_with_TOC.pdf'
        self.outFilePdf = os.path.normpath(os.path.join(self.pathToPDF, 'toc_file.pdf'))
        # remove toc_pdf file from previously running
        try:
            os.remove(self.outFilePdf)
            os.remove(os.path.normpath(os.path.join(self.pathToRTF, 'toc_file.pdf')))
        except Exception as e:
            print("Not found file "+ str(self.outFilePdf ))
            pass

        METADATA = self.gui.entry_var2.get()

        # Get tlfmetadata and move to py dict.
        self.util.assign_meta()

        # Place ProgressBar onto main Frame after button 'GO!' is pressed.
        self.gui.pb1.place(x=40, y=490, width=625, height=10)
        self.gui.pb1['value'] = 0  # reset Progress Bar.

        pbConstant = 8
        numberOfLogEvents = self.util.get_event_number(METADATA)

        self.gui.pb1['maximum'] = pbConstant + numberOfLogEvents * 3

        # If check-box is off - then we need to convert raw files to PDF.
        # After converted we can proceed and combine files.
        # if self.gui.pas_check_var.get() == 0:
        tlfs, tlfs_count = self.util.get_tlf_list(METADATA)
        self.util.convert_to_pdf(in_list=tlfs, rtf_folder_dir=self.CWD, pdf_folder_dir=self.pathToPDF)


        # Count how many files we have to combine.
        count = tlfs_count
        # Combine PDF files listed in list_pdf.
        if tlfs_count > 0:
            self.gui.logger.warning('\nNow combining outputs into PDF...')
            self.gui.logger.warning('\nSearching in ' + str(self.pathToPDF))
            self.util.add_bmk_to_file(input_dir=self.pathToPDF,
                                      meta_data_file=METADATA,
                                      title_sep=self.gui.title_separator,
                                      add_popul=self.gui.add_population)


            self.util.go_combine_selected_pdf(dir=self.pathToPDF,
                                              meta_data_=METADATA,
                                              out_name=self.gui.OUTPUT_FILENAME,
                                              prot_fl=False,
                                              title_sep=self.gui.title_separator,
                                              add_popul=self.gui.add_population)

            self.gui.logger.warning(
                '\nINFO: Job finished! ' + str(count) + ' files were added to ' + self.gui.OUTPUT_FILENAME)
            self.gui.logger.warning('\nINFO: ' + self.gui.OUTPUT_FILENAME + ' is saved in ' + str(os.getcwd()))
        else:
            self.gui.logger.warning('WARNING: No files to concatenate. Check ' + str(self.pathToRTF) + '.')
            # Reset Progress Bar
            self.gui.pb1['value'] = 0

    @staticmethod
    def get_toc_page_numb(path_to_pdf: str):
        """
        Function to get total number of pages from pdf TOC file to use as cut-off during final merge
        Param:
        path_to_pdf - path to pdf TOC file

        Return:
            doc_page_numb - total pages into pdf TOC file
        """

        with fitz.open(path_to_pdf) as f:
            doc_page_numb = len(f)
        return doc_page_numb

    @staticmethod
    def extcract_bmk_to_list(path_to_pdf_file: str, w_page: int,
                             tab_ch: str, pg_char: str, out_df_headers: list) -> pd.DataFrame:
        layer_ = list()
        toc = list()
        tab_bmk_page_lst = list()

        doc = fitz.open(path_to_pdf_file)
        toc = doc.get_toc(simple=True)  # The format of a bookmark likes: [layer, name, page]
        # layer = 1 main header, layer = 2 - sub-header and go on
        # name - text of bookmark
        # page - page where it bookmark root
        doc.close()

        for i in range(0, len(toc)):
            layer_.append(toc[i][0])
            toc[i][0] = tab_ch * int(toc[i][0] - 1)  # convert layer number to space tabulation
            # add chars string at the begining of page numbers
            toc[i][2] = pg_char + str(toc[i][2])
        # make new list of list: first elem in each elem present concatenation of tabulation symbols and bmk text
        # second - bmk page
        for elem in toc:
            tab_bmk_page_lst.append([''.join(elem[:2]), elem[2]])

        # wrap bmk text and page number to page width
        wrapper = textwrap.TextWrapper(width=w_page - len(pg_char) - 5)  # add 6 for 99_999 pages
        for elem in tab_bmk_page_lst:
            dedented_text = textwrap.dedent(text=elem[0])
            original = wrapper.fill(text=dedented_text)
            elem[0] = original.rstrip(' ')
            new_line_index = str(elem[0]).find('\n')
            # if new line char is in string find len of 'new line' part of original string
            if new_line_index != -1:
                symb_to_add = w_page - (len(elem[0]) - new_line_index + 1 + len(str(elem[1])))
                elem[1] = str(elem[1]).rstrip().rjust(symb_to_add, ".")
            elif new_line_index == -1:
                symb_to_add = w_page - (len(elem[0]))
                elem[1] = str(elem[1]).rstrip().rjust(symb_to_add, ".")

        bmk_df = pd.DataFrame(tab_bmk_page_lst, columns=out_df_headers)

        return bmk_df

    @staticmethod
    def update_toc_pages(input_file: str, page_char: str, w_page: int, page_numb_to_add: int) -> str:
        with open(input_file, "r", encoding="latin-1") as f:
            new_content = ""
            for line in f:
                if page_char in line:
                    page_index = line.index(page_char)
                    # slice page numbers from line
                    s_0 = line[:page_index]  # string without page nunber
                    s_1 = line[page_index + len(page_char):]  # page number for string

                    numeric_filter = filter(str.isdigit, s_1)
                    numeric_string = "".join(numeric_filter)
                    n_string = int(numeric_string)
                    n_string += page_numb_to_add
                    page_str_numb = page_char + str(n_string)
                    tmp_line = s_0 + page_str_numb
                    if len(tmp_line) < w_page + 1:
                        add_dot = w_page - len(tmp_line) + 1
                        tmp_line = s_0 + "." * add_dot + str(page_str_numb)
                    elif len(tmp_line) > w_page + 1:
                        length_dif = len(tmp_line) - w_page - 1
                        ind_to_rem = len(str(page_str_numb)) + length_dif
                        # check for additional '.' symbols -> shifts after incerease page number: 10 (len=2) instead
                        # TODO: use any other more "smart way" for replace chars
                        if tmp_line[-ind_to_rem:-(ind_to_rem - length_dif)] == ('.',):
                            tmp_line = s_0[:-1] + str(page_str_numb)
                        if tmp_line[-ind_to_rem:-(ind_to_rem - length_dif)] in ('.',):
                            tmp_line = s_0[:-1] + str(page_str_numb)
                        elif tmp_line[-ind_to_rem:-(ind_to_rem - length_dif)] in ('..',):
                            tmp_line = s_0[:-2] + str(page_str_numb)
                        elif tmp_line[-ind_to_rem:-(ind_to_rem - length_dif)] in ('...',):
                            tmp_line = s_0[:-3] + str(page_str_numb)
                        elif tmp_line[-ind_to_rem:-(ind_to_rem - length_dif)] in ('....',):
                            tmp_line = s_0[:-4] + str(page_str_numb)
                        elif tmp_line[-ind_to_rem:-(ind_to_rem - length_dif)] in ('.....',):
                            tmp_line = s_0[:-5] + str(page_str_numb)

                    new_content += tmp_line + "\n"
                else:
                    new_content += line
        return new_content

    @staticmethod
    def make_toc_pdf(input_file: str, page_char: str, w_page: int, usage_font: str, doc_page_size: tuple,
                     out_file: str, bimo_char: str, bimo_tab_replace: str, page_size_const: any,
                     all_doc_page_numb: any, out_file_type: str, folder_with_font: str) -> None:
        # out_file_type 'temp' or 'final': temp - for get pdf_file pages number, final - for merge with main doc

        pdf = FPDF(orientation='L', unit='in',
                   format=(doc_page_size.height * page_size_const, doc_page_size.width * page_size_const))
        pdf.set_margins(0.38, 0.78, 0.38)
        # Add a page
        # set style and size of font that you want in the pdf
        pdf.set_doc_option('core_fonts_encoding', 'windows-1252')
        pdf.add_font(usage_font, '', folder_with_font, uni=True)
        pdf.set_font(usage_font, size=9)
        pdf.set_auto_page_break(True, 0.38)
        pdf.add_page()

        # insert the texts in pdf
        to_toc = pdf.add_link()
        pdf.cell(0, h=0.2, txt="Table of Contents", border=0, ln=2, align='C', link=to_toc)
        pdf.set_link(to_toc, y=0.0, page=-1)
        pdf.set_font(usage_font, size=8)

        link_str_list = []

        # open the text file in read mode
        with open(input_file, "r", encoding="latin-1") as f:
            if out_file_type == 'final':
                for line in f:
                    if page_char in line:
                        link_str_list.append(line)
                        page_index = line.index(page_char)
                        s_1 = line[page_index + len(page_char):]  # page number for string
                        n_string = int(s_1)
                        to_pg = pdf.add_link()
                        for toc_txt in link_str_list:

                            toc_txt = toc_txt.replace(bimo_char, bimo_tab_replace)
                            toc_txt = toc_txt[:-1]

                            if page_char in toc_txt and len(toc_txt) > w_page + 2:
                                len_diff = len(page_char) - (len(toc_txt) - w_page - 2)
                                toc_txt = toc_txt.replace(page_char, '.' * len_diff)
                            else:
                                toc_txt = toc_txt.replace(page_char, '.' * len(page_char))

                            pdf.cell(0, h=0.2, txt=toc_txt, ln=1, align='L', link=to_pg)
                            pdf.set_link(to_pg, y=0.0, page=n_string)
                        link_str_list = []
                    elif page_char not in line:
                        link_str_list.append(line)

                for i in range(0, all_doc_page_numb):
                    pdf.add_page()
                    # save the pdf with name .pdf
            elif out_file_type == 'temp':
                pdf.cell(0, h=0.2, txt="Table of Contents", border=0, ln=2, align='C')
                pdf.set_font(usage_font, size=8)
                for x in f:
                    pdf.cell(0, h=0.2, txt=x, ln=1, align='L')

        pdf.output(out_file)

    # Prepare TOC shell file
    def add_toc(self):
        """Add table of contents to the combined PDF"""
        t_char = '*page:'
        bimo_tab_char = '$$$$'
        bimo_tabulation_replace = '    '

        col_header = ['name', 'page']
        PR_TO_IN = 1 / 72
        pdf_del = False
        page_w = 152

        self.out_file_txt = str(self.pathToPDF) + "\\" + 'toc_file.txt'

        # custom value for specific fonts
        if usage_font_c == 'VictorMono':
            page_w = 167
        elif usage_font_c == 'Monoid':
            page_w = 160
        elif usage_font_c == 'EversonMono':
            page_w = 159
        elif usage_font_c == 'Lekton':
            page_w = 182
        elif usage_font_c == 'CamingoCode':
            page_w = 166

        pathToFile = str(self.pathToPDF)[:-4] + str(self.gui.OUTPUT_FILENAME)

        df = self.extcract_bmk_to_list(pathToFile, page_w, bimo_tab_char, t_char, col_header)

        # REPLACE CHARACTER INTO BOOKMARK DATAFRAME IN CASE IF SELECTED FONT FROM PART SUPPORT FONT LIST
        # TODO: move to function
        r, c = df.shape
        for i in range(0, r):
            if u'\u2013' in df.loc[i, 'name']:
                self.gui.logger.warning('\nReplace \\u2013 to – in ' + str(df.loc[i, 'name']))
            if u'\u2011' in df.loc[i, 'name']:
                self.gui.logger.warning('\nReplace \\u2011 to ‑ in ' + str(df.loc[i, 'name']))
            if u'\u2265' in df.loc[i, 'name']:
                self.gui.logger.warning('\nReplace \\u2065 to >= in ' + str(df.loc[i, 'name']))
            if u'\u2264' in df.loc[i, 'name']:
                self.gui.logger.warning('\nReplace \\u2064 to <= in ' + str(df.loc[i, 'name']))
            if u'\u03bc' in df.loc[i, 'name']:
                self.gui.logger.warning('\nReplace \\u03bc to μ in ' + str(df.loc[i, 'name']))
            if u'\u2019' in df.loc[i, 'name']:
                self.gui.logger.warning('\nReplace \\u2019  to \' in ' + str(df.loc[i, 'name']))

            df.loc[i, 'name'] = df.loc[i, 'name'].replace(u'\u2013', '-')
            df.loc[i, 'name'] = df.loc[i, 'name'].replace(u'\u2011', '-')
            df.loc[i, 'name'] = df.loc[i, 'name'].replace(u'\u03bc', chr(181))
            df.loc[i, 'name'] = df.loc[i, 'name'].replace(u"\u2019", '\'')
            df.loc[i, 'name'] = df.loc[i, 'name'].replace(u'\u2265', chr(62) + chr(61))
            df.loc[i, 'name'] = df.loc[i, 'name'].replace(u'\u2264', chr(60) + chr(61))

        with open(self.out_file_txt, "w", encoding="utf-8") as f:
            np.savetxt(f, df.to_numpy(), fmt='%s')

        # Get page size from main document
        doc = fitz.open(pathToFile)
        page = doc[0]  # Get first page
        main_doc_page_size = page.rect.br  # Get bottom-right point of page rect
        doc.close()

        Page = namedtuple("Page", "width height")
        page_size = Page(main_doc_page_size.x, main_doc_page_size.y)

        # convert txt file to pdf to get toc-pdf file number of pages
        self.make_toc_pdf(input_file=self.out_file_txt, page_char=t_char, w_page=page_w,
                          usage_font=usage_font_c, doc_page_size=page_size,
                          out_file=self.outFilePdf, bimo_char=bimo_tab_char, bimo_tab_replace=bimo_tabulation_replace,
                          page_size_const=PR_TO_IN, all_doc_page_numb=None, out_file_type='temp',
                          folder_with_font=font_folder)

        toc_pdf_file_page_numb = self.get_toc_page_numb(self.outFilePdf)

        self.gui.logger.warning('INFO: Total TOC pages number:  ' + str(toc_pdf_file_page_numb) + '.')

        comb_page_numb = self.get_toc_page_numb(self.gui.OUTPUT_FILENAME)

        self.gui.logger.warning('INFO: Total combained pdf file pages number :  ' + str(comb_page_numb) + '.')
        self.gui.logger.warning('INFO: Creating TOC...')

        # return to toc txt file and add to page number add toc-pdf file total pages number
        new_file_content = self.update_toc_pages(input_file=self.out_file_txt, page_char=t_char,
                                                 w_page=page_w, page_numb_to_add=toc_pdf_file_page_numb)

        writing_file = open(self.out_file_txt, "w", encoding="latin-1")
        writing_file.write(new_file_content)
        writing_file.close()

        # TODO: Move to gui module
        # open TOC-shell file to view and edit
        q = messagebox.askokcancel(title=None, message="TOC template ready and save at " + str(
            self.out_file_txt) + " Do you want to open and edit the file?",
                                   default='ok')
        if q == True:
            # self.out_file_txt.close()
            os.startfile(self.out_file_txt)
            messagebox.showinfo(title="TOC ready", message='Press OK after view TOC')
        elif q == False:
            pass

        # convert txt file to output pdf
        self.make_toc_pdf(input_file=self.out_file_txt, page_char=t_char, w_page=page_w,
                          usage_font=usage_font_c, doc_page_size=page_size,
                          out_file=self.outFilePdf, bimo_char=bimo_tab_char, bimo_tab_replace=bimo_tabulation_replace,
                          page_size_const=PR_TO_IN, all_doc_page_numb=comb_page_numb, out_file_type='final',
                          folder_with_font=font_folder)


        # concatenate both file toc-pdf file and original input file
        pdfs = [str(self.outFilePdf), str(pathToFile)]  # toc file, main_cont file
        result = fitz.open()
        for pdf in pdfs:
            mfile= fitz.open(pdf)
            if pdf == self.outFilePdf:
                result.insert_pdf(mfile, from_page=0, to_page=toc_pdf_file_page_numb - 1, links=True,
                                  annots=True)
                general_toc = [[1, 'Table of Contents', 1]]
            else:
                result.insert_pdf(mfile, links=True, annots=True)
                tmp_toc = mfile.get_toc(simple=False)
                for t in tmp_toc:  # increase toc2 page numbers
                    t[2] += toc_pdf_file_page_numb  # by old len(doc1)
                general_toc += tmp_toc
            mfile.close()

        with fitz.open(self.outFilePdf) as mfile:
            link_cnti = 0
            link_skip = 0
            for pinput in mfile:  # iterate through input pages
                links = pinput.get_links()  # get list of links
                link_cnti += len(links)  # count how many
                pout = result[pinput.number]  # read corresp. output page
                for l in links:  # iterate though the links
                    if l["kind"] == fitz.LINK_NAMED:  # we do not handle named links
                        link_skip += 1  # count them
                        continue
                    pout.insert_link(l)  # simply output the others
        try:
            mfile.close()
        except:
            pass

        result.set_toc(general_toc)

        if self.gui.pas_check_var.get() == 1:
            print('############################################################################')
            print("Set password: ", self.gui.pas_check_var.get())
            print(self.gui.entry_var5.get())
            print('############################################################################')
            tlfs, tlfs_count = self.util.get_tlf_list(METADATA)

            result.save(self.outFilePdfToc, pretty=True, garbage=4, deflate=True,
                        encryption=fitz.PDF_ENCRYPT_AES_256,
                        user_pw=tlfs[0], owner_pw=self.gui.entry_var5.get())
        if self.gui.pas_check_var.get() == 0:
            result.save(self.outFilePdfToc, pretty=True, garbage=4, deflate=True)

        result.close()


        self.gui.logger.warning('INFO: TOC successfully created!')

        # Cleanup section
        try:
            # Remove temporary PDF file
            if os.path.exists(self.outFilePdf):
                os.remove(self.outFilePdf)
        except Exception as e:
            self.gui.logger.error(f"Error removing temporary PDF: {str(e)}")

        try:
            # Remove temporary txt file
            if os.path.exists(self.out_file_txt):
                os.remove(self.out_file_txt)
        except Exception as e:
            self.gui.logger.error(f"Error removing temporary txt file: {str(e)}")

        # Remove font cache file if it exists
        try:
            font_cache = f"{usage_font_c}.pkl"
            if os.path.exists(font_cache):
                os.remove(font_cache)
        except Exception as e:
            self.gui.logger.warning(f"Could not remove font cache file: {str(e)}")
            # Non-critical error, we can continue

        # Show final message and file
        if os.path.exists(self.outFilePdfToc):
            q = messagebox.askokcancel(
                title="PDF Created",
                message=f"Combined pdf ready and saved at {self.outFilePdfToc}. Do you want to open the file?",
                default='ok'
            )
            if q:
                os.startfile(self.outFilePdfToc)

        # Make 'GO' button active again
        self.gui.btn_go.config(state='normal')
