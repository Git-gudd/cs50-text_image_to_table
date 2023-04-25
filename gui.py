import tkinter as tk
from tkinter import font as tkFont
from tkinter import filedialog as fd
from tkinter import ttk
from tkinter import messagebox as msg

from PIL import ImageGrab, Image, ImageTk
from win32com.client import Dispatch
import configparser

import logging
logging.basicConfig(level=logging.INFO)

import os
FOLDER_PATH = os.path.dirname(os.path.abspath(__file__))
TEMP_IMAGE_ADDRESS = os.path.join(FOLDER_PATH, "resources", "temp.jpg")
TEMP_CSV_ADDRESS = os.path.join(FOLDER_PATH, "result.csv")
PNG_ADDRESS = os.path.join(FOLDER_PATH, "resources", "icon", "png")
CONFIG_ADDRESS = os.path.join(FOLDER_PATH, "resources", "settings.ini")
EXAMPLE_ADDRESS = os.path.join(FOLDER_PATH, "examples")

import OCR_detector as ocr_detector
from OCR_detector import NotEnoughCol, NotEnoughRow

class OOP():
    def __init__(self):
        self.win = tk.Tk()
        self.win.title("Image to table data")
        self.win.geometry("-90+0")
        self.win.configure(background="White")

        # Color
        self.component_color="White"
        self.background_color="White"
        self.font_color="Black"
        self.active_color="#00CCB3"
        self.ttkStyle = ttk.Style()
        self.ttkStyle.configure('new.TLabelframe', background='White', foreground='White', highlightbackground="White")
        self.tiny_font = tkFont.Font(size=1)
        # Measurement
        self.gui_width_char = 30
        self.gui_width_pixel = 400
        self.frame_padding = 10
        self.grid_padding = 5

        self.canvas_width = self.gui_width_pixel
        self.canvas_height = 250

        self.button_width = 10
        self.button_height = 1

        # Object & ID
        self.config = configparser.ConfigParser()
        
        self.excel = None
        self.excelObject = None

        self.fullsized_image = None

        self.create_widgets()
        self.set_currency_option_state()
        self.set_remove_option_state()
        self.set_adv_option_states()
        self.display_image(image_path=TEMP_IMAGE_ADDRESS)
        
    # UI modules

    def create_widgets(self):
        #__________________________________________________________________________________
        # Frame
        self.create_widgets_frame()

        #__________________________________________________________________________________
        # Input frame's widget
        self.create_input_frame_widget()

        #__________________________________________________________________________________
        # Window bar frame's widget
        self.create_exit_bar_widget()

        #__________________________________________________________________________________
        # Options value
        ## Basic
        self.alignment_value = tk.IntVar()
        self.use_excel_value = tk.IntVar()
        self.header_value = tk.IntVar()
        self.currency_value = tk.IntVar()
        self.currency_character = tk.StringVar()
        self.remove_seperator_value = tk.IntVar()
        self.thousand_seperator = tk.StringVar()
        
        ## Advance
        ### Dynamic
        self.clustering_mode_value=tk.IntVar()
        self.col_dist_value=tk.IntVar()
        self.row_dist_value=tk.IntVar()
        self.min_col_value=tk.IntVar()
        self.min_row_value=tk.IntVar()
        self.accuracy_value = tk.IntVar()
        ### Fixed
        self.num_of_col=tk.IntVar()
        self.num_of_row=tk.IntVar()

        #__________________________________________________________________________________
        # Basic options's widget
        self.create_basic_option_widget()

        #__________________________________________________________________________________
        # Advance options's widget
        self.create_adv_option_widget()

        self.read_settings()

    def create_widgets_frame(self):

        self.table_input = tk.Frame(self.win,
                                    background=self.background_color,
                                    padx=self.frame_padding)
        self.table_input.grid(column=0, row=1, sticky='we')


        self.table_options = ttk.LabelFrame(self.win,
                                            text="Options",
                                            style='new.TLabelframe')
        self.table_options.grid(column=0, row=2, sticky='nesw', pady=self.grid_padding)

        self.table_adv_options = ttk.LabelFrame(self.win,
                                            text="Advanced Options",
                                            style='new.TLabelframe')
        self.table_adv_options.grid(column=0, row=3, sticky='we')

        self.adv_dynamic_options = ttk.LabelFrame(self.table_adv_options,
                                            text="Dynamic",
                                            style='new.TLabelframe')
        self.adv_dynamic_options.grid(column=0, row=1, columnspan=2, sticky='we', padx=self.grid_padding)
        self.adv_dynamic_options.columnconfigure(0, weight=1)
        self.adv_dynamic_options.columnconfigure(1, weight=1)

        self.adv_fixed_options = ttk.LabelFrame(self.table_adv_options,
                                            text="Fixed",
                                            style='new.TLabelframe')
        self.adv_fixed_options.grid(column=0, row=2, columnspan=2, sticky='we', padx=self.grid_padding, pady=self.grid_padding)

        self.col_dynamic_options = tk.Frame(self.adv_dynamic_options,
                                    background=self.background_color,
                                    padx=self.frame_padding, 
                                    pady=self.frame_padding,
                                    highlightthickness=1)
        self.col_dynamic_options.grid(column=0, row=2, columnspan=2, sticky='we')
        self.col_dynamic_options.columnconfigure(0, weight=1)
        

        self.row_dynamic_options = tk.Frame(self.adv_dynamic_options,
                                    background=self.background_color,
                                    padx=self.frame_padding, 
                                    pady=self.frame_padding,
                                    highlightthickness=1)
        self.row_dynamic_options.grid(column=0, row=3, columnspan=2, sticky='we')
        self.row_dynamic_options.columnconfigure(0, weight=1)

    def create_exit_bar_widget(self):
        # To make sure the row and col in grid expand full
        self.win.rowconfigure(0, weight=1)
        self.win.columnconfigure(0, weight=1)

        self.quit_button = tk.Button(self.win,
                                     text="Exit",
                                     background=self.active_color,
                                     foreground="White",
                                     activebackground="#E81123",
                                     highlightthickness=1,
                                     #width=6,
                                     height=1,
                                     relief="flat",
                                     command=self._quit)
        self.quit_button.grid(column=0, row=0, padx=0, pady=0, ipadx=0, ipady=0,sticky='we')     

    def create_input_frame_widget(self):
        self.imageDisplay = tk.Button(self.table_input, 
                                      background=self.component_color,
                                      width=self.canvas_width, 
                                      height=self.canvas_height,
                                      relief="groove",
                                      command=self.open_image_full_size)
        self.imageDisplay.grid(column=0, row=0, columnspan=2, sticky='nsew', pady=self.grid_padding)


        self.file_dialog_button = tk.Button(self.table_input, 
                                text='Select from computer', 
                                foreground=self.font_color,
                                background=self.component_color,
                                height=self.button_height,
                                command=self.get_file_from_filedialog,
                                padx=10)
        self.file_dialog_button.grid(column=0, row=1)

        self.paste_clipboard_button = tk.Button(self.table_input, 
                                text='Paste from clipboard', 
                                foreground=self.font_color,
                                background=self.component_color,
                                height=self.button_height,
                                command=self.get_file_from_clipboard,
                                padx=10)
        self.paste_clipboard_button.grid(column=1, row=1)

        self.previous_image_button = tk.Button(self.table_input, 
                                text='Run on current image', 
                                foreground=self.font_color,
                                background=self.component_color,
                                height=self.button_height,
                                command=self.get_previous_file,
                                padx=10)
        self.previous_image_button.grid(column=0, row=2, columnspan=2, pady=self.grid_padding)

    def create_basic_option_widget(self):
        ### Alignment options
        self.alignmentFrame = tk.Frame(self.table_options,
                                       background=self.background_color,
                                       width=self.gui_width_pixel)
                                       #background=self.background_color,
                                       
        self.alignmentFrame.grid(column=0, row=0, columnspan=3, sticky="we", padx=self.grid_padding)
        

        self.alignment_label = tk.Label(self.alignmentFrame, 
                                        text="Choose text alignment",
                                        foreground=self.font_color,
                                        background=self.component_color)
        self.left_icon = ImageTk.PhotoImage(Image.open(os.path.join(PNG_ADDRESS, "left.png")))
        self.center_icon = ImageTk.PhotoImage(Image.open(os.path.join(PNG_ADDRESS, "center.png")))
        self.right_icon = ImageTk.PhotoImage(Image.open(os.path.join(PNG_ADDRESS, "right.png")))
        
        self.left_align = tk.Radiobutton(self.alignmentFrame, image=self.left_icon, variable=self.alignment_value, value=-1)
        self.center_align = tk.Radiobutton(self.alignmentFrame, image=self.center_icon, variable=self.alignment_value, value=0)
        self.right_align = tk.Radiobutton(self.alignmentFrame, image=self.right_icon, variable=self.alignment_value, value=1)

        for widget in (self.left_align, self.center_align, self.right_align):
            widget.configure(foreground=self.font_color,
                             background=self.component_color,
                             indicatoron=0,
                             borderwidth=0,
                             selectcolor="#D8DCD8")
            
        self.paddingGap1 = tk.Canvas(self.alignmentFrame,
                             width=self.gui_width_pixel,
                             background=self.component_color,
                             highlightthickness=0,
                             height=10)
        self.paddingGap1.grid(column=0, row=2, columnspan=3, sticky='we')

        self.alignment_label.grid(column=0, row=0, columnspan=3)
        self.left_align.grid(column=0, row=1, sticky='e')
        self.center_align.grid(column=1, row=1)
        self.right_align.grid(column=2, row=1, sticky='w')
        
        self.left_align.select() # Default is left alignment
        
        ### Output to Excel options
        self.excel_output = tk.Checkbutton(self.table_options, 
                                     text="Output to excel", 
                                     variable=self.use_excel_value,
                                     foreground=self.font_color,
                                     background=self.component_color,
                                     selectcolor=self.active_color)
        self.excel_output.grid(column=1, row=1, sticky="w", padx=self.grid_padding)

        ### Table Headers options
        self.header = tk.Checkbutton(self.table_options, 
                                     text="Contains headers", 
                                     variable=self.header_value,
                                     foreground=self.font_color,
                                     background=self.component_color,
                                     selectcolor=self.active_color)
        self.header.grid(column=2, row=1, sticky="w", padx=self.grid_padding)

        ### Remove Currency format options
        self.currency = tk.Checkbutton(self.table_options, 
                                       text="Remove currency", 
                                       variable=self.currency_value,
                                       foreground=self.font_color,
                                       background=self.component_color,
                                       selectcolor=self.active_color,
                                       command=self.set_currency_option_state)
        self.currency.grid(column=1, row=2, sticky="w", padx=self.grid_padding)

        self.currency_entry = ttk.Combobox(self.table_options,
                                           foreground=self.active_color,
                                           background=self.component_color,
                                           textvariable = self.currency_character,
                                           justify="center",
                                           width=5)
        self.currency_entry['values']=("$", "£", "€")
        self.currency_entry.grid(column=1, row=3, sticky="w", padx=self.grid_padding)

        ### Remove thousand seperator options
        self.remove_seperator = tk.Checkbutton(self.table_options, 
                                       text="Remove character(s)", 
                                       variable=self.remove_seperator_value,
                                       foreground=self.font_color,
                                       background=self.component_color,
                                       selectcolor=self.active_color,
                                       command=self.set_remove_option_state)
        self.remove_seperator.grid(column=2, row=2, sticky="w", padx=self.grid_padding)

        self.seperator_entry = tk.Entry(self.table_options,
                                       foreground=self.active_color,
                                       background=self.component_color,
                                       textvariable = self.thousand_seperator,
                                       justify="center",
                                       width=10)
        self.seperator_entry.grid(column=2, row=3, sticky="w", padx=self.grid_padding)


        self.paddingGap2 = tk.Canvas(self.table_options,
                             background=self.component_color,
                             highlightthickness=0,
                             width=20,
                             height=35)
        self.paddingGap2.grid(column=0, row=1, rowspan=3, sticky='ns')

        self.paddingGap3 = tk.Canvas(self.table_options,
                             width=self.gui_width_pixel,
                             background=self.component_color,
                             highlightthickness=0,
                             height=5)
        self.paddingGap3.grid(column=0, row=4, columnspan=3, sticky='we')

    def create_adv_option_widget(self):
        self.dynamic_clustering_mode = tk.Radiobutton(self.table_adv_options,
                                                      text="Dynamic", 
                                                      variable=self.clustering_mode_value,
                                                      selectcolor=self.active_color, 
                                                      background=self.background_color,
                                                      value=0, 
                                                      command=self.set_adv_option_states)
        
        self.fixed_clustering_mode = tk.Radiobutton(self.table_adv_options, 
                                                    text="Fixed", 
                                                    variable=self.clustering_mode_value,
                                                    selectcolor=self.active_color,
                                                    background=self.background_color,
                                                    value=1, 
                                                    command=self.set_adv_option_states)
        self.dynamic_clustering_mode.grid(column=0, row=0)
        self.fixed_clustering_mode.grid(column=1, row=0)

        self.create_dynamic_option()
        self.create_fixed_option()

        self.set_adv_option_states()

    def create_dynamic_option(self):
        ### Minimum col cell option
        # Label
        self.min_col_label = tk.Label(self.adv_dynamic_options,
                                      text="Min number of columns:",
                                      foreground=self.font_color,
                                      background=self.component_color,
                                      highlightthickness=0)
        self.min_col_label.grid(column=0, row=0)
        # Input
        self.min_col_entry = tk.Spinbox(self.adv_dynamic_options,
                                        foreground=self.active_color,
                                        background=self.component_color,
                                        textvariable = self.min_col_value,
                                        justify="center",
                                        width=5,
                                        from_=1, to=20)

        self.min_col_entry.grid(column=0, row=1, pady=self.grid_padding)

        ### Minimum row cell option
        # Label
        self.min_row_label = tk.Label(self.adv_dynamic_options,
                                      text="Min number of rows:",
                                      foreground=self.font_color,
                                      background=self.component_color,
                                      highlightthickness=0)
        self.min_row_label.grid(column=1, row=0)
        # Input
        self.min_row_entry = tk.Spinbox(self.adv_dynamic_options,
                                        foreground=self.active_color,
                                        background=self.component_color,
                                        textvariable = self.min_row_value,
                                        justify="center",
                                        width=5,
                                        from_=1, to=20)
        self.min_row_entry.grid(column=1, row=1, pady=self.grid_padding)

        ### Distance threshold options

        self.col_less_label = tk.Label(self.col_dynamic_options,
                                      text="Less columns",
                                      foreground=self.font_color,
                                      background=self.component_color,
                                      highlightthickness=0)
        self.col_less_label.grid(column=0, row=3, sticky="w")

        self.col_more_label = tk.Label(self.col_dynamic_options,
                                      text="More columns",
                                      foreground=self.font_color,
                                      background=self.component_color,
                                      highlightthickness=0)
        self.col_more_label.grid(column=1, row=3, sticky="e")

        self.col_dist_slider = tk.Scale(self.col_dynamic_options,
                                    variable=self.col_dist_value,
                                    from_=10,
                                    to=0,
                                    resolution=1,
                                    tickinterval=-1,
                                    showvalue=0,
                                    orient='horizontal',
                                    #length=self.gui_width_pixel,
                                    foreground=self.component_color,
                                    background=self.component_color,
                                    highlightthickness=0,
                                    relief='flat',
                                    font=self.tiny_font)
        self.col_dist_slider.grid(column=0, row=4, columnspan=2, pady=self.grid_padding, sticky='we')

        ### Row options

        self.row_less_label = tk.Label(self.row_dynamic_options,
                                      text="Less row",
                                      foreground=self.font_color,
                                      background=self.component_color,
                                      highlightthickness=0)
        self.row_less_label.grid(column=0, row=6, sticky="w")

        self.row_more_label = tk.Label(self.row_dynamic_options,
                                      text="More row",
                                      foreground=self.font_color,
                                      background=self.component_color,
                                      highlightthickness=0)
        self.row_more_label.grid(column=1, row=6, sticky="e")
        self.row_dist_slider = tk.Scale(self.row_dynamic_options,
                                    variable=self.row_dist_value,
                                    from_=10,
                                    to=0,
                                    resolution=1,
                                    tickinterval=-1,
                                    showvalue=0,
                                    orient='horizontal',
                                    #length=self.gui_width_pixel,
                                    foreground=self.component_color,
                                    background=self.component_color,
                                    highlightthickness=0,
                                    relief='flat',
                                    font=self.tiny_font)
        self.row_dist_slider.grid(column=0, row=7, columnspan=2, pady=self.grid_padding, sticky='we')

    def create_fixed_option(self):
        ### Minimum col cell option
        # Label
        self.fixed_col_label = tk.Label(self.adv_fixed_options,
                                      text="Fixed columns:",
                                      foreground=self.font_color,
                                      background=self.component_color,
                                      highlightthickness=0)
        self.fixed_col_label.grid(column=0, row=0)
        # Input

        self.fixed_col_entry = tk.Spinbox(self.adv_fixed_options,
                                          foreground=self.active_color,
                                          background=self.component_color,
                                          textvariable = self.num_of_col,
                                          justify="center",
                                          width=5,
                                          from_=1, to=20)
        self.fixed_col_entry.grid(column=0, row=1)

        ### Minimum row cell option
        # Label
        self.fixed_row_label = tk.Label(self.adv_fixed_options,
                                      text="Fixed rows:",
                                      foreground=self.font_color,
                                      background=self.component_color,
                                      highlightthickness=0)
        self.fixed_row_label.grid(column=1, row=0)
        # Input
        self.fixed_row_entry = tk.Spinbox(self.adv_fixed_options,
                                          foreground=self.active_color,
                                          background=self.component_color,
                                          textvariable = self.num_of_row,
                                          justify="center",
                                          width=5,
                                          from_=1, to=20)
        self.fixed_row_entry.grid(column=1, row=1)

        self.paddingGap4= tk.Canvas(self.adv_fixed_options,
                             width=self.gui_width_pixel,
                             background=self.component_color,
                             highlightthickness=0,
                             height=10)
        self.paddingGap4.grid(column=0, row=2, columnspan=2, sticky='we')

    # Backend functional modules

    def get_file_from_filedialog(self):
        if self.use_excel_value.get() == 1:
            self.close_excel_file()

        filetypes = (('JPG files', '*.jpg'), ('PNG files','*.png'))
        filename = fd.askopenfilename(title='Open a file', initialdir=EXAMPLE_ADDRESS, filetypes=filetypes)
        if filename:
            image = Image.open(filename)
            image.save(TEMP_IMAGE_ADDRESS)
            self.display_image(image_path=TEMP_IMAGE_ADDRESS)
            self.run_ocr(TEMP_IMAGE_ADDRESS)
    
    def get_file_from_clipboard(self):
        if self.use_excel_value.get() == 1:
            self.close_excel_file()

        image = ImageGrab.grabclipboard()
        if not image:
            return
        
        rgb_im = image.convert('RGB')
        rgb_im.save(TEMP_IMAGE_ADDRESS)
        self.display_image(image_path=TEMP_IMAGE_ADDRESS)

        self.run_ocr(TEMP_IMAGE_ADDRESS)

    def get_previous_file(self):
        if self.use_excel_value.get() == 1:
            self.close_excel_file()
        self.display_image(image_path=TEMP_IMAGE_ADDRESS)
        self.run_ocr(TEMP_IMAGE_ADDRESS)

    def set_adv_option_states(self):
        mode = self.clustering_mode_value.get()

        if mode == 0:
            self.set_dynamic_states(1)
            self.set_fixed_states(0)
        else:
            self.set_dynamic_states(0)
            self.set_fixed_states(1)

    def set_dynamic_states(self, value):
        color = {"disabled": "Gray", "normal": self.active_color}
        current_state = 'normal' if value == 1 else 'disabled'
        self.min_col_entry.configure(state=current_state)
        self.min_row_entry.configure(state=current_state)
        self.col_dist_slider.configure(state=current_state, troughcolor=color[current_state])
        self.row_dist_slider.configure(state=current_state, troughcolor=color[current_state])

    def set_fixed_states(self, value):
        current_state = 'normal' if value == 1 else 'disabled'
        self.fixed_col_entry.configure(state=current_state)
        self.fixed_row_entry.configure(state=current_state)

    def set_currency_option_state(self):
        current_state = 'normal' if self.currency_value.get() == 1 else "disabled"
        self.currency_entry.configure(state=current_state)

    def set_remove_option_state(self):
        current_state = 'normal' if self.remove_seperator_value.get() == 1 else "disabled"
        self.seperator_entry.configure(state=current_state)

    def get_fixed_col_row(self):
        # Clustering mode: Dynamic or Fixed
        # Fixed mode
        if self.clustering_mode_value.get() == 0:
            fixed_num_of_col = 0
            fixed_num_of_row = 0
        # Dynamic mode
        else:
            fixed_num_of_col = self.num_of_col.get()
            fixed_num_of_row = self.num_of_row.get()
            if fixed_num_of_col == 0:
                msg.showwarning("Number of columns is set fixed at 0", "The program will now apply dynamic settings when constructing columns from this image")
            if fixed_num_of_row == 0:
                msg.showwarning("Number of rows is set fixed at 0", "The program will now apply dynamic settings when constructing rows from this image")

        return fixed_num_of_col, fixed_num_of_row

    def get_min_col_row(self):
        # Fixed mode
        if self.clustering_mode_value.get() == 1:
            min_num_of_col = 1
            min_num_of_row = 1
        # Dynamic mode
        else:
            min_num_of_col = self.min_col_value.get()
            min_num_of_row = self.min_row_value.get()

        return min_num_of_col, min_num_of_row

    def get_char_to_remove(self):
        # Removing characters
        char_to_remove = ""
        if self.currency_value.get() != 0:
            char_to_remove += self.currency_character.get()
        if self.remove_seperator_value.get() != 0:
            char_to_remove += self.thousand_seperator.get()

        return char_to_remove

    def run_ocr(self, image_path):
        # Preparing parameters to send to tesseract module

        # Calculate dynamic distance threshold

        # Fixed number of columns and rows
        fixed_num_of_col, fixed_num_of_row = self.get_fixed_col_row()
        min_num_of_col, min_num_of_row = self.get_min_col_row()

        # Removing characters
        char_to_remove = self.get_char_to_remove()

        args = {"image": image_path, 
                "output": TEMP_CSV_ADDRESS, 
                "have_header": self.header_value.get(),
                "remove_character": char_to_remove,
                "column_alignment": self.alignment_value.get(),
                "col_dist_thresh": self.col_dist_value.get(),
                "row_dist_thresh": self.row_dist_value.get(),
                "fixed_col": fixed_num_of_col,
                "fixed_row": fixed_num_of_row,              
                "min_col_cell": min_num_of_col,
                "min_row_cell": min_num_of_row,
                "min_conf": 0,
                "fulltable": 1}
        try:
            image_array = ocr_detector.main(args)
        except ValueError:
            msg.showerror("Error", "No text is detected in this image. Try another image")
            return
        except NotEnoughCol:
            msg.showwarning("Cannot construct table", "Text is detected in this image. But does not have the minimum number of columns required. Try lowering the minimum number of columns")
            return
        except NotEnoughRow:
            msg.showwarning("Cannot construct table", "Text is detected in this image. But does not have the minimum number of rows required. Try lowering the minimum number of rows")
            return

        self.display_image(image_array=image_array)

        if self.use_excel_value.get() == 1:
            self.open_excel_file(TEMP_CSV_ADDRESS)
        return

    def display_image(self, image_array=None, image_path=None):
        if image_path:
            image = Image.open(image_path)
        else:
            image = Image.fromarray(image_array)
        self.fullsized_image = image
        ratio = min(self.canvas_height / image.height, self.canvas_width / image.width)
        new_size = (int(image.width * ratio) , int(image.height * ratio))
        self.resized_image = ImageTk.PhotoImage(image.resize(new_size))

        self.imageDisplay.configure(image=self.resized_image)

    def open_image_full_size(self):
        self.fullsized_image.show()

    def open_excel_file(self, filePath):
        if not self.excel:
            try:
                self.excel = Dispatch('Excel.Application')
            except Exception:
                msg.showerror("Cannot open .csv file with Excel", "It has appeared that you don't have Microsoft Excel program on this computer. Please try using other programs that can open .csv file")
                return
            self.excel.Visible = True #If we want to see it change
            
        self.excelObject = self.excel.Workbooks.Open(filePath)

    def close_excel_file(self):
        if not self.excelObject:
            return
        try:
            self.excelObject.Close(True)
        except Exception:
            logging.debug("Excel has already closed")
            self.excel.Quit()
            self.excel = None

        self.excelObject = None

    def read_settings(self):
        self.config.read(CONFIG_ADDRESS)
        #Basic
        basic_config = self.config['basic.options']
        self.alignment_value.set(int(basic_config['alignment.value']))
        self.use_excel_value.set(int(basic_config['use.excel.value']))
        self.header_value.set(int(basic_config['header.value']))
        self.currency_value.set(int(basic_config['currency.value']))
        self.currency_character.set(basic_config['currency.char'])
        self.remove_seperator_value.set(int(basic_config['seperator.value']))
        self.thousand_seperator.set(basic_config['seperator.char'])
        
        ## Advance
        ### Dynamic
        dynamic_config = self.config['dynamic.options']
        self.clustering_mode_value.set(int(dynamic_config['cluster.mode']))
        self.col_dist_value.set(int(dynamic_config['col.dist']))
        self.row_dist_value.set(int(dynamic_config['row.dist']))
        self.min_col_value.set(int(dynamic_config['min.col']))
        self.min_row_value.set(int(dynamic_config['min.row']))
        self.accuracy_value.set(int(dynamic_config['accuracy.value']))
        ### Fixed
        fixed_config = self.config['fixed.options']
        self.num_of_col.set(int(fixed_config['num.col']))
        self.num_of_row.set(int(fixed_config['num.row']))

    def save_settings(self):
        #Basic
        self.config['basic.options'] = {}
        basic_config = self.config['basic.options']
        basic_config['alignment.value'] = str(self.alignment_value.get())
        basic_config['use.excel.value'] = str(self.use_excel_value.get())
        basic_config['header.value'] = str(self.header_value.get())
        basic_config['currency.value'] = str(self.currency_value.get())
        basic_config['currency.char'] = self.currency_character.get()
        basic_config['seperator.value'] = str(self.remove_seperator_value.get())
        basic_config['seperator.char'] = self.thousand_seperator.get()
        
        ## Advance
        ### Dynamic
        self.config['dynamic.options'] = {}
        dynamic_config = self.config['dynamic.options']
        dynamic_config['cluster.mode'] = str(self.clustering_mode_value.get())
        dynamic_config['col.dist'] = str(self.col_dist_value.get())
        dynamic_config['row.dist'] = str(self.row_dist_value.get())
        dynamic_config['min.col'] = str(self.min_col_value.get())
        dynamic_config['min.row'] = str(self.min_row_value.get())
        dynamic_config['accuracy.value'] = str(self.accuracy_value.get())
        ### Fixed
        self.config['fixed.options'] = {}
        fixed_config = self.config['fixed.options']
        fixed_config['num.col'] = str(self.num_of_col.get())
        fixed_config['num.row'] = str(self.num_of_row.get())

        with open(CONFIG_ADDRESS, 'w') as configfile:
            self.config.write(configfile)

    def _quit(self):
        logging.debug("Quitting")
        self.save_settings()
        self.close_excel_file()
        if self.excel:
            self.excel.Quit()

        self.win.quit()
        self.win.destroy()
        exit()

if __name__ == "__main__":
    oop=OOP()
    oop.win.mainloop()
        
