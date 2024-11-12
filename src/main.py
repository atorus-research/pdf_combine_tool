#!/usr/bin/env python
# coding: utf-8

import logging
from src.gui import GUICore
from src.pdf_compiler import PDFCompiler
from src.pdf_util import PDFUtility

# Set logging level.
logging.basicConfig(level=logging.WARNING)

# Create GUI.
gui = GUICore()

# Set parameters of main window.
gui.root.geometry('720x520')

# Create object of Utility class.
util = PDFUtility(gui)

# Create object of Action class.
pc = PDFCompiler(gui=gui, util=util)

# Link BTN1 from GUI with SELECT_FOLDER command from Action class.
gui.link_btn_to_command(btn=gui.btn1, command=util.select_folder)
# Link BTN2 from GUI with SELECT_METADATA command from Action class.
gui.link_btn_to_command(btn=gui.btn2, command=util.select_metadata)


def execute_pdf_operations():
    if gui.toc_var.get() == 'no_toc':
        pc.combine_pdfs()  # This will now return early for no_toc
    else:
        pc.combine_pdfs()
        pc.add_toc()


# Link BTN_GO from GUI with COMBINE_PDFs command from Action class.
gui.link_btn_to_command(btn=gui.btn_go, command=execute_pdf_operations)

# Change background color for all widgets.
gui.change_bg_color(gui.root.winfo_children(), 'white')

# Show GUI.
gui.root.mainloop()
