import os
import tkinter as tk
import tkinter.filedialog as filedialog
import tkinter.messagebox as messagebox
import tkinter.ttk as ttk
from tkinter import *

from openpyxl import load_workbook

import VocabProblemFiles as vpf
from VocabProblemFiles import FileGenerator


class MainWidget:
    def __init__(self):
        # Added filenames including its directory
        self.filenames = []

        # Each detail -> (prob_pivot, ans_pivot, index_included)
        self.file_details = dict()

        # Each detail -> (dirname, name, ftype, size, random_seed)
        self.output_details = [None for _ in range(5)]

        # User interface configurations
        self.root = Tk()
        self.root.title("Vocabulary")
        self.root.resizable(False, False)
        current_dirname = os.path.split(os.path.abspath(__file__))[0].strip()
        icon = PhotoImage(file = os.path.join(current_dirname, 'logo_image.png'))
        self.root.iconphoto(False, icon)

        # Menu bar
        menubar = Menu(self.root)

        # File Menu
        menu_file = Menu(menubar, tearoff=0)
        menu_file.add_command(label="Add files", command=self.file_add)
        menu_file.add_command(label="Delete files", command=self.file_del)
        menu_file.add_separator()
        menu_file.add_command(label="Exit", command=self.root.destroy)
        menubar.add_cascade(label="File", menu=menu_file)

        # Option Menu
        menu_option = Menu(menubar, tearoff=0)
        menu_option.add_command(label='Added file details', command=self.file_details_widget)
        menu_option.add_command(label='Detailed options', command=self.detailed_option_widget)
        menubar.add_cascade(label='Option', menu=menu_option)

        self.root.config(menu=menubar)

        # List Frame -> Display the list of added files
        list_frame = Frame(self.root)
        list_frame.pack(fill='both', padx=5, pady=5)

        x_scrollbar = Scrollbar(list_frame, orient='horizontal')
        x_scrollbar.pack(side='bottom', fill='x')

        y_scrollbar = Scrollbar(list_frame)
        y_scrollbar.pack(side='right', fill='y')

        self.list_file = Listbox(list_frame, selectmode='extended', height=15, xscrollcommand=x_scrollbar.set, yscrollcommand=y_scrollbar.set)
        self.list_file.pack(side='left', fill='both', expand=True)
        x_scrollbar.config(command=self.list_file.xview)
        y_scrollbar.config(command=self.list_file.yview)

        # Path Frame -> Select the directory for saving the result
        path_frame = LabelFrame(self.root, text='Save path')
        path_frame.pack(fill='x', padx=5, pady=5)

        self.txt_dest_path = Entry(path_frame)
        self.txt_dest_path.insert(0, vpf.default_dir)
        self.txt_dest_path.pack(side='left', fill='x', expand=True, ipady=4, padx=5, pady=5)

        btn_dest_path = Button(path_frame, text='find', width=5, command=self.file_save_dir)
        btn_dest_path.pack(side='right', padx=5, pady=5)

        # Filename Frame
        frame_filename = LabelFrame(self.root, text='Filename')
        frame_filename.pack(fill='x', padx=5, pady=5)

        # Filename -> Name
        Label(frame_filename, text='Name', width=5).pack(side='left', padx=5, pady=5)

        self.txt_name = Entry(frame_filename, width=40)
        self.txt_name.insert(0, vpf.default_name)
        self.txt_name.pack(side='left', padx=5, pady=5)

        # Filename -> format(ftype)
        Label(frame_filename, text='Format', width=5).pack(side='left', padx=5, pady=5)

        self.cmb_ftype = ttk.Combobox(frame_filename, state='readonly', values=vpf.ftypes, width=10)
        self.cmb_ftype.current(0)
        self.cmb_ftype.pack(side='left', padx=5, pady=5)
    
        # Progress Frame
        frame_progress = LabelFrame(self.root, text='progress')
        frame_progress.pack(fill='x', padx=5, pady=5)

        self.p_var = DoubleVar()
        self.progress_bar = ttk.Progressbar(frame_progress, maximum=1, variable=self.p_var)
        self.progress_bar.pack(fill='x', padx=5, pady=5)

        # Run Frame
        frame_run = Frame(self.root)
        frame_run.pack(fill='both', padx=5, pady=5)

        btn_start = Button(frame_run, padx=3, pady=3, text='start', width=8, command=self.start_process)
        btn_start.pack(side='right', padx=5, pady=5)

    def run(self):
        self.root.mainloop()

    def file_add(self):
        init_dir = vpf.document_path
        types = vpf.supporting_types
        filenames = filedialog.askopenfilenames(initialdir=init_dir, title='select file', filetypes=types)
        for filename in filenames:
            if filename not in self.filenames: self.filenames.append(filename)
        names = [os.path.split(filename)[1] + ' << ' + filename for filename in self.filenames]
        self.list_file.delete(0, END)
        for name in names: self.list_file.insert(END, name)

    def file_del(self):
        for idx in self.list_file.curselection():
            filename = self.list_file.get(idx).split('<<')[-1].strip()
            self.list_file.delete(idx)
            self.filenames.remove(filename)

    def file_details_widget(self):
        if len(self.filenames) == 0:
            messagebox.showerror("Error", "There are no added files")
            return

        widget = FileDetailsWidget(self)
        widget.run()

    def detailed_option_widget(self):
        widget = DetailedOptionWidget(self)
        widget.run()

    def file_save_dir(self):
        dirname = filedialog.askdirectory()
        if dirname != '':
            self.dirname = dirname
            self.txt_dest_path.delete(0, tk.END)
            self.txt_dest_path.insert(0, self.dirname)

    def start_process(self):
        if len(self.filenames) == 0:
            messagebox.showerror("Error", "There are not enough files")
            return

        fg = FileGenerator()

        for idx, filename in enumerate(self.filenames):
            if filename not in self.file_details.keys(): self.file_details[filename] = (0, 1, True)
            fg.read(filename, *self.file_details[filename])
            
            # progress bar update
            self.p_var.set((idx+1) / (len(self.filenames)+1))
            self.progress_bar.update()

        self.output_details[0] = self.txt_dest_path.get()
        self.output_details[1] = self.txt_name.get()
        self.output_details[2] = self.cmb_ftype.get()
        fg.make_both(*self.output_details)

        # progress bar update
        self.p_var.set(1)
        self.progress_bar.update()

        # show the result
        os.startfile(vpf.default_dir if self.output_details[0] is None else self.output_details[0])
        

class FileDetailsWidget:
    def __init__(self, master):
        self.master = master
        self.filenames = master.filenames

        self.root = Toplevel()
        self.root.title('Added file details')
        self.root.resizable(False, False)
        current_dirname = os.path.split(os.path.abspath(__file__))[0].strip()
        icon = PhotoImage(file = os.path.join(current_dirname, 'logo_image.png'))
        self.root.iconphoto(False, icon)

        self.cmb_details_dict = dict()

        for filename in self.filenames:
            frame_pivot = LabelFrame(self.root, text=os.path.split(filename)[-1])
            frame_pivot.pack(padx=5, pady=5)

            pivot_opt = ExcelWorkbookMethods.get_index(filename)
            index_opt = ['True', 'False']

            # problem pivot section
            Label(frame_pivot, text='problem pivot', width=12).pack(side='left', padx=5, pady=5)
            cmb_prob_pivot = ttk.Combobox(frame_pivot, state='readonly', values=pivot_opt, width=10)
            cmb_prob_pivot.current(0 if filename not in self.master.file_details.keys() else self.master.file_details[filename][0])
            cmb_prob_pivot.pack(side='left', padx=5, pady=5)

            # answer pivot section
            Label(frame_pivot, text='answer pivot', width=12).pack(side='left', padx=5, pady=5)
            cmb_ans_pivot = ttk.Combobox(frame_pivot, state='readonly', values=pivot_opt, width=10)
            cmb_ans_pivot.current(1 if filename not in self.master.file_details.keys() else self.master.file_details[filename][1])
            cmb_ans_pivot.pack(side='left', padx=5, pady=5)

            # index included section
            Label(frame_pivot, text='index included', width=12).pack(side='left', padx=5, pady=5)
            cmb_idx_inc = ttk.Combobox(frame_pivot, state='readonly', values=index_opt, width=10)
            cmb_idx_inc.set("True" if filename not in self.master.file_details.keys() else str(self.master.file_details[filename][2]))
            cmb_idx_inc.pack(side='left', padx=5, pady=5)

            self.cmb_details_dict[filename] = (cmb_prob_pivot, cmb_ans_pivot, cmb_idx_inc)

        btn_start = Button(self.root, padx=3, pady=3, text='apply', width=8, command=self.apply)
        btn_start.pack(side='right', padx=5, pady=5)

    def run(self):
        self.root.mainloop()
    
    def apply(self):
        self.master.file_details = dict()
        for filename in self.filenames:
            prob_piv = int(self.cmb_details_dict[filename][0].get().split()[0].strip())
            ans_piv = int(self.cmb_details_dict[filename][1].get().split()[0].strip())
            idx_inc = bool(self.cmb_details_dict[filename][2].get())
            self.master.file_details[filename] = (prob_piv, ans_piv, idx_inc)
        self.root.destroy()


class DetailedOptionWidget:
    def __init__(self, master):
        self.master = master
        self.filenames = master.filenames

        self.root = Toplevel()
        self.root.title('Detailed file options')
        self.root.resizable(False, False)
        current_dirname = os.path.split(os.path.abspath(__file__))[0].strip()
        icon = PhotoImage(file = os.path.join(current_dirname, 'logo_image.png'))
        self.root.iconphoto(False, icon)

        # Option Frame
        frame_option = LabelFrame(self.root, text='Options')
        frame_option.pack(fill='x', padx=5, pady=5)

        # Size option
        Label(frame_option, text='Size').pack(side='left', padx=5, pady=5)

        self.txt_size = Entry(frame_option, width=8)
        self.txt_size.insert(0, "All" if self.master.output_details[3] == None else self.master.output_details[3])
        self.txt_size.pack(side='left', fill='x', expand=True, ipady=4, padx=5, pady=5)

        btn_start = Button(self.root, padx=3, pady=3, text='apply', width=8, command=self.apply)
        btn_start.pack(side='right', padx=5, pady=5)

        # Random state option
        Label(frame_option, text='Random seed').pack(side='left', padx=5, pady=5)
        
        seed_opt = ["None"] + list(range(0, 101, 20))
        self.cmb_random_seed = ttk.Combobox(frame_option, state='readonly', values=seed_opt, width=10)
        self.cmb_random_seed.set(str(self.master.output_details[4]))
        self.cmb_random_seed.pack(side='left', padx=5, pady=5)

    def run(self):
        self.root.mainloop()
    
    def apply(self):
        self.master.output_details[3] = None if self.txt_size.get() == 'All' else int(self.txt_size.get())
        self.master.output_details[4] = None if self.cmb_random_seed.get() == 'None' else int(self.cmb_random_seed.get())
        self.root.destroy()


class ExcelWorkbookMethods:
    @staticmethod
    def get_index(filename):
        wb = load_workbook(filename, read_only=True)
        return ["{} {}".format(idx, cell) for idx, cell in enumerate(list(wb.active.values)[0])]


if __name__ == '__main__':
    main = MainWidget()
    main.run()