from tkinter import *
from tkinter import ttk, filedialog

from lib.ExcelManipulator import ExcelManipulator
from lib.DataProcessor import DataProcessor

class GUIApp:
  def __init__(self):
    self.selected_file = None

    self.root = Tk()

    self.job = StringVar(value='group')
    self.selected_sheet = StringVar()

    self.sheets = []

    self.data_processor = DataProcessor()


  def start(self):
    self.root.title('MJ - Excel Cleaner')

    mainframe = ttk.Frame(self.root, padding='3 3 12 12')
    mainframe.grid(column=0, row=0, sticky=(N, W, E, S))

    self.root.columnconfigure(0, weight=1)
    self.root.rowconfigure(0, weight=1)

    select_file_text = ttk.Label(mainframe, text='Select an Excel file:')
    select_file_text.grid(row=1, column=1, sticky=(W, S), padx=(0, 200))

    select_file_button = ttk.Button(mainframe, text='Select file', command=self._select_file)
    select_file_button.grid(row=1, column=2, sticky=(N, E))

    self.sheet_combobox = ttk.Combobox(mainframe, textvariable=self.selected_sheet, values=self.sheets)
    self.sheet_combobox.grid(row=2, column=2, sticky=(W, E))

    group_radio_button = ttk.Radiobutton(mainframe, text='Group Duplicates', value='group', variable=self.job)
    merge_radio_button = ttk.Radiobutton(mainframe, text='Merge Duplicates', value='merge', variable=self.job)

    group_radio_button.grid(row=2, column=1, sticky=(W, S), pady=(8, 0))
    merge_radio_button.grid(row=3, column=1, sticky=(W, S))

    self.process_button = ttk.Button(mainframe, text='Process', command=self._process, state=DISABLED)
    self.process_button.grid(row=4, column=1, columnspan=2, sticky=(W, E, S), pady=(8, 0))

    self.result_text_content = StringVar(value='Select a file and click "Process" to see a result: ')
    
    result_text = ttk.Label(mainframe, textvariable=self.result_text_content)
    self.result_button = ttk.Button(mainframe, state=DISABLED, text='Download', command=self._download)
    
    result_text.grid(row=5, column=1, sticky=(W, S), pady=(8, 0))
    self.result_button.grid(row=5, column=2, sticky=(W, E))

    self.root.mainloop()

  def _select_file(self):
    file_path = filedialog.askopenfilename(title='Select an Excel file')

    if file_path: 
      self.selected_file = file_path
      self.process_button['state'] = NORMAL
      self.result_text_content.set('Click "process" to get a result: ')

      self.excel_manipulator = ExcelManipulator(file_path)
      self.sheets = self.excel_manipulator.get_sheet_names()
      self.sheet_combobox['values'] = self.sheets

      self.selected_sheet.set(self.sheets[0])
    
  def _process(self):
    self.excel_manipulator.reload_file()
    self.excel_manipulator.set_worksheet(self.selected_sheet.get())
    
    self.result_text_content.set('Here is the result, click to download: ')
    self.result_button['state'] = NORMAL

    data = self.excel_manipulator.extract_data()
    grouped_data = self.data_processor.group_duplicates(data, [0, 2], [0, 2])
    
    if self.job.get() == 'group':
      self.excel_manipulator.write_new_data(grouped_data)
    elif self.job.get() == 'merge':
      merged_data = self.data_processor.merge_data(grouped_data, [0, 2])
      self.excel_manipulator.write_new_data(merged_data)

  def _download(self):
    file_path = filedialog.asksaveasfilename(title='Save Excel file', defaultextension='xlsx')
    self.excel_manipulator.save_file(file_path)

    self.result_text_content.set('The result has been downloaded to: \n' + file_path)
