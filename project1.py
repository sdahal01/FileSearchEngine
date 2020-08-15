from tkinter import *
from tkinter.filedialog import askdirectory
import tkinter as tk
from tkinter.font import Font
import os
import textract


SUPPORTED_EXT = ['.txt', '.docx', '.pptx', '.xlsx']


def doc2string(path):
    # define 'fileContent'
    fileContent = ""
    path = path.replace('/', '\\')
    fileContent = textract.process(path)
    parsedFile = fileContent.decode('utf-8').split('\n')
    return parsedFile


def has_text(filename, path, text):
    lines = doc2string(path)
    print(filename)
    for line in lines:
        if line.lower().find(text.lower()) is not -1:
            return True
    return False


class MyFrame(Frame):
    def __init__(self):
        Frame.__init__(self)
        # Title and Iconbitmap
        self.master.title("File Search Engine")
        #self.master.iconbitmap(r'icon.ico')
        self.master.rowconfigure(5, weight=1)
        self.master.columnconfigure(5, weight=1)
        self.grid(sticky=W + E + N + S)
        # Setting the label and folder input
        # '#995511' = orange, '#063a45' = darkgreen '#052f38' = light green
        self.frame_folder = tk.Frame(self, bg='#052f38')
        self.folder_label = tk.Label(self.frame_folder,
                                     bg='#052f38',
                                     text="Select a folder to search: ", font='verdana 8 bold  italic',
                                     fg='white', width=45)

        self.folder_input = Entry(self.frame_folder, width=60, bd=5, bg='#0e869e')
        self.folder_input.insert(0, "C:/")
        self.folder_button = tk.Button(self.frame_folder,
                                       text="Select Folder", font='verdana 8 bold  italic',
                                       bg='#b4d0d6', fg='black',
                                       bd=5, width=10,
                                       command=self.load_file)

        self.folder_label.grid(row=0, column=0, sticky=NE)
        self.folder_input.grid(row=0, column=1)
        self.folder_button.grid(row=0, column=1, sticky=E)

        # search lable and button
        self.frame_text = tk.Frame(self, bg='#052f38')
        self.text_search_label = tk.Label(self.frame_text,
                                          text="Search Text: ", font='verdana 8 bold  italic',
                                          bg='#052f38',
                                          fg='white',
                                          width=45)

        self.text_to_search = Entry(self.frame_text, width=60, bd=5, bg='#0e869e')
        self.search_button = tk.Button(self.frame_text,
                                       text="Search", font='verdana 8 bold  italic',
                                       bg='#b4d0d6',
                                       fg='black',
                                       bd=5,
                                       command=self.search_text,
                                       width=10)
        self.text_search_label.grid(row=0, column=0, sticky=E)
        self.text_to_search.grid(row=0, column=1)
        self.search_button.grid(row=0, column=1, sticky=E)
        # Listbox
        self.frame_listbox = Frame(self, bg='#052f38')
        self.listbox = Listbox(self.frame_listbox,
                               selectmode=SINGLE,
                               width=100,
                               bd=5, bg='#0e869e')
        # Scrollbar
        self.scrollbar = Scrollbar(self.frame_listbox, orient="vertical", bd=5, bg='#80c1ff')
        self.scrollbar.config(command=self.listbox.yview)
        self.listbox.config(yscrollcommand=self.scrollbar.set)
        self.listbox.grid(row=0, column=0, sticky=NE)
        self.listbox.bind('<Double-Button-1>', self.open_file)
        self.scrollbar.grid(row=0, column=1, sticky=NS)
        # Open File
        self.frame_open_file = Frame(self, bg='#052f38')
        self.open_button = tk.Button(self.frame_open_file,
                                     text="Open", font='verdana 8 bold  italic',
                                     bg='#b4d0d6',
                                     fg='black',
                                     bd=5,
                                     command=self.open_file,
                                     width=10)
        self.file_content = Text(self.frame_open_file, height=15, width=75, bd=5, bg='#0e869e')
        self.scrollbar1 = Scrollbar(self.frame_open_file, orient="vertical", bd=5)
        self.scrollbar1.config(command=self.file_content.yview)
        self.file_content.config(yscrollcommand=self.scrollbar1.set)
        self.open_button.grid(row=0, column=0, sticky=N)
        self.file_content.grid(row=1, sticky=E)
        self.scrollbar1.grid(row=1, column=1, sticky=NS)
        # Setting up the layout
        self.frame_folder.grid(row=0, sticky=E)
        self.frame_text.grid(row=1, sticky=E)
        self.frame_listbox.grid(row=2, sticky=NSEW)
        self.frame_open_file.grid(row=3, sticky=NSEW)

    def load_file(self):
        self.folder_input.delete(0, END)
        self.folder_input.insert(0, askdirectory())
        if self.folder_input.get() == '':
            self.folder_input.insert(0, "C:/")

    def search_text(self):
        folder = self.folder_input.get()
        search_text = self.text_to_search.get()
        self.listbox.delete(0, END)
        print(folder)
        for root, dirs, files in os.walk(folder, topdown=False):
            print("Number of folder searched: " + str(len(dirs)))
            print("Number of file searched: " + str(len(files)))
            for name in files:
                if search_text.lower() in name.lower():
                    self.listbox.insert(END, os.path.join(root, name))
                elif name.endswith(tuple(SUPPORTED_EXT)):
                    if has_text(name, os.path.join(root, name), search_text):
                        self.listbox.insert(END, os.path.join(root, name))

    def open_file(self, event=0):
        self.file_content.delete("1.0", END)
        search_text = self.text_to_search.get()
        path = self.listbox.get(self.listbox.curselection())
        lines = doc2string(path)
        for line in lines:
            if search_text in line:
                l1 = line.partition(search_text)
                for l in l1:
                    if l != search_text:
                        self.file_content.insert(END, l)
                    else:
                        self.file_content.tag_config("highlight", background='orange')
                        self.file_content.insert(END, l, "highlight")
            else:
                self.file_content.insert(END, line + '\n')


if __name__ == "__main__":
    MyFrame().mainloop()
