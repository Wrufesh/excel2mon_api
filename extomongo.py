import xlrd

# Start and spacing in excel document row
TITLE_START = 0
TITLE_SPACING = 3
DATA_START = 1
DATA_SPACING = 3


def get_row_data_cells_only(start_at, diff, *args):
    data_list = []
    for cell in args[start_at::diff]:
        data_list.append(cell)
    return data_list


class GetDataFromExcel(object):

    def __init__(self, filename):
        self.collections = []
        self.coll_field_dict = {}

        f = xlrd.open_workbook(filename)
        sheet_names = f.sheet_names()
        xl_sheet = f.sheet_by_name(sheet_names[0])

        i = 0
        while i >= 0:
            try:
                row = xl_sheet.row(i)
                if i == 0:
                    data = get_row_data_cells_only(TITLE_START,
                                                   TITLE_SPACING, *row)
                    self.collections = [cell.value for cell in data]
                else:
                    data = get_row_data_cells_only(DATA_START,
                                                   DATA_SPACING, *row)
                    self.make_document(data)
                i = i + 1
            except IndexError:
                break
        pass

    def make_document(self, data):
        for i, cell in enumerate(data):
            try:
                type(self.coll_field_dict[self.collections[i]])
            except KeyError:
                self.coll_field_dict[self.collections[i]] = []

            if cell.ctype != 0:
                self.coll_field_dict[self.collections[i]].append(cell.value)
        pass

    def get_collections_names(self):
        return self.collections

    def get_collection_fields(self, collection_name):
        return self.coll_field_dict.get(collection_name, None)

    # Usages Pymongo==3.0
    def write_list_to_coll(self, db, collection_name, lyst):
        collection = db[collection_name]
        dic = {}
        fields = self.coll_field_dict.get(collection_name, None)
        if len(fields) == len(lyst):

            for key, item in zip(fields, lyst):
                dic[key] = item
        if len(dic) == 0:
            return False
        else:
            collection.insert(dic)
            return True

    def write_from_input(self, db, collection_name):
        collection = db[collection_name]
        dic = {}
        fields = self.coll_field_dict.get(collection_name, None)
        print 'Note: Input str data enclosed in inverted comma'
        for key in fields:
            try:
                in_data = input('Enter the value for ' + key + ': ')
                dic[key] = in_data
            except NameError:
                print 'ERROR:: Note: Input str data must be enclosed in inverted comma'
                break
        if len(dic) < len(fields):
            return False
        else:
            collection.insert(dic)
            return True

