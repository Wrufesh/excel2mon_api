import xlrd
import pdb
from pymongo import MongoClient

connection = MongoClient('mongodb://localhost:27017/')
db = connection['db_name']
collections = []
coll_field_dict = {}


def get_row_data_cells_only(start_at, diff, *args):
    data_list = []
    for cell in args[start_at::diff]:
        data_list.append(cell)
    return data_list


def make_document(data):
    for i, cell in enumerate(data):
        try:
            type(coll_field_dict[collections[i]])
        except KeyError:
            coll_field_dict[collections[i]] = []

        if cell.ctype != 0:
            coll_field_dict[collections[i]].append(cell.value)
    pass

f = xlrd.open_workbook(filename='input_files/mongo_model.xlsx',
                       on_demand=False
                       )
sheet_names = f.sheet_names()
xl_sheet = f.sheet_by_name(sheet_names[0])

i = 0
while i >= 0:
    try:
        row = xl_sheet.row(i)
        if i == 0:
            data = get_row_data_cells_only(0, 3, *row)
            collections = [cell.value for cell in data]
        else:
            data = get_row_data_cells_only(1, 3, *row)
            make_document(data)
        i = i + 1
    except IndexError:
        break

# pdb.set_trace()
print collections
print coll_field_dict.keys()[0]
print len(coll_field_dict[coll_field_dict.keys()[0]])
print type(row)
