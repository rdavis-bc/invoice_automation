from __future__ import annotations
from typing import Optional, Union
import openpyxl
import datetime as dt
import os 
import sys
from pathlib import Path
import re
import xlwings as xw
import glob
from operator import itemgetter 


def create_data_directory(base_dir: Union[str, None] = None) -> str:
    """
    A utility function to quickly create a new directory with the option
    to be placed in a base directory
    Params:
        base_dir: str, a relative path in which you want to place your folder
    Returns: the relative path to the created directory
    Raises:
    >>> results = create_data_directory(base_dir='base_dir')
    >>> results

    """
    try:
        the_day = dt.datetime.now().strftime('%Y-%m-%d') # getting the current day
        if not base_dir: #creating the directory name
            directory = f"{the_day}"
        else:
            directory = f"{base_dir}/{the_day}"
        path = Path(os.getcwd()) # absolute path to current dir to join with dir name and create new dir
        dir_path = os.path.join(path, directory)
        
        os.mkdir(dir_path)
    except Exception as e:
        print(f'{directory} or {base_dir} already exists Error:{e!r}')
    return directory

# def excel_file_to_pdf(filepath:str, sheet: int, loc: str, pdf_name: str):

#     return

class UserInput():
    all_inputs: list = []
    """
    Represents the user input
    >>> UserInput(base_dir=sys.argv[1], dest_dir=sys.argv[2])
    """
    def __init__(self, base_dir: Optional[str], dest_dir: Optional[str]) -> None:
        self.base_dir = base_dir
        self.dest_dir = dest_dir
        self.all_inputs.append((self.base_dir, self.dest_dir,))

    def copying_file_to_dest(self) -> None :
        print(f'copying data to {self.dest_dir}')

class AdditionalUserInput(UserInput):
    """Inherits from UserInput which was used previously as a helper"""

    all_inputs: list = []
    def __init__(self, base_dir: Optional[str], dest_dir: Optional[str]) -> None:
        super().__init__(base_dir, dest_dir)

    def logging_choices(self) -> None:
        self.all_inputs.append((self.base_dir, self.dest_dir,))
        print(f'{self.base_dir!r} and {self.dest_dir!r} logged')

# define sheet serializer classs

class SheetSerializer():

    def __init__(self, sheet, client_name:str, month:str, dest_dir) -> None:
        self.sheet = sheet
        self.client_name = client_name
        self.month = month
        self.dest_dir = dest_dir

class PDFSerializer(SheetSerializer):

    def __init__(self, sheet, client_name: str, month: str, dest_dir: str) -> None:
        super().__init__(sheet, client_name, month, dest_dir)

    def serialize(self) -> str:
        global wkbk
        global directory
        xw_wkbk = xw.Book(f"{wkbk}")
        xw_sheet = xw_wkbk.sheets[self.sheet]
        pdf_path = f'{directory}/{self.client_name + self.month}.pdf'
        try:
            xw_wkbk.to_pdf(path=pdf_path, include=self.sheet)
        except Exception as e:
            print(e)

class WorkbookParser():
    def __init__(self, workbook: openpyxl.Workbook) -> None:
        self.wkbk = workbook
        self.wkshts : list["Worksheet"] = None
        # self.clientname: Optional[str] = None
        # self.month: Optional[str] = None
        
    def parser(self) -> None:
        global anas_input
        sheets = self.wkbk.sheetnames[1:]
        for ix, sheet in enumerate(sheets):
            client = self.wkbk.worksheets[ix]
            mx_rw, mx_col = client.max_row, client.max_column
            prev_val = None
            for row in range(1, mx_rw + 1): # starting with the first row (top)
                for col in range(mx_col, 0, -1): # starting with the last column
                    val = client.cell(row=row, column=col).value #value of the coordinates
                    lw_cs_val = val.lower() if type(val) == str else val # trnasofmring to lower case if datatype is str
                    if re.match(r'.*date.*', lw_cs_val if type(val) == str else 'negative') and type(prev_val) == dt.datetime :# if date is in the current value and the 
                        # cell_contents = prev_val.split(' ')                               # is of type datetime      # re.match(r'.*\d{4}
                        prev_val_month :dt.datetime = prev_val.strftime("%b")
                        # insert code for creating pdf
                        pdf_serializer = PDFSerializer(sheet=sheet, client_name=sheet, month=prev_val_month, dest_dir=anas_input.dest_dir)
                        pdf_serializer.serialize()
                        prev_val = None
                        break
                    prev_val = val



if __name__ == "__main__":
    

    # argument should go python main.py <directory of excel sheets> <desired directory for pdfs>
    anas_input = UserInput(base_dir=sys.argv[1], dest_dir=sys.argv[2])
    directory = create_data_directory(base_dir=anas_input.dest_dir)

    directory_of_excel_sheets = sys.argv[1]
    desired_directory_for_pdfs = sys.argv[2]

    the_workbooks = glob.glob(f"{directory_of_excel_sheets}/*.xlsx")
    # the_sheets = glob.glob(f"excel_sheets/*.xlsx")
    # files = itemgetter(*[0])(the_sheets) not needed

    # file should be searched by suffix
    for wkbk in the_workbooks:
        book = openpyxl.load_workbook(f"{wkbk}", data_only=True)
        workbook_parser = WorkbookParser(workbook=book)
        workbook_parser.parser()

    exit(0)
    # the sheet name has the client's names
    book.sheetnames

    for a in book.worksheets:
        print(a.title)

    # data range
    client1 = book.worksheets[1]
    mx_rw, mx_col = client1.max_row, client1.max_column

    # reading data (1-based ix) 
    # go to row then change col instead
    prev_val = None
    for row in range(1, mx_rw + 1): # starting with the first row (top)
        for col in range(mx_col, 0, -1): # starting with the last column
            val = client1.cell(row=row, column=col).value #value of the coordinates
            # if val:
            #     print(val)
            #     print(col,row)
            lw_cs_val = val.lower() if (type.val) == str else val # trnasofmring to lower case if datatype is str
            if re.match(r'.*date.*', lw_cs_val) and type(prev_val) == dt.datetime :# if date is in the current value and the prev_val (column to right)
                # cell_contents = prev_val.split(' ')                               # is of type datetime      # re.match(r'.*\d{4}', prev_val):
                # month = cell_contents[1]
                prev_val_month :dt.datetime = prev_val.strftime("%b")
                # insert code for creating pdf
            prev_val = val
            print(val)

    # I want to take each sheet serialized separately
    # then convert this file to pdf
    # https://stackoverflow.com/questions/57724345/print-excel-to-pdf-with-xlwings
    # client1.cell(row=9, column=8).value.strftime("%b")
    # re.match(r'.*date.*', client1.cell(row=6, column=7).value.lower())
    # re.match(r'.*date.*', 'dates')

    # use this to change the format of the sheets (ultimately take allow the option of taking arguments)
    # create a function that takes a sheet number of the xlsx and then create the pdf of it
    book = xw.Book("excel_sheets/2022.02.28 Test Invoice to automate.xlsx")
    sheet = book.sheets[1]

    pdf_path = os.path.join(os.getcwd(), 'file.pdf') 

    book.to_pdf(path=pdf_path, include=2)
