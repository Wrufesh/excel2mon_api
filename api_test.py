from extomongo import GetDataFromExcel
from pymongo import MongoClient

connection = MongoClient('mongodb://localhost:27017/')
db = connection['db_test']

x = GetDataFromExcel('input_files/mongo_model.xlsx')

print x.get_collections_names()

print x.get_collection_fields('Account')

x.write_list_to_coll(db, 'Geography__c', [1,2,3,4,5,6,7,8,9,10,11,12,13])

x.write_from_input(db, 'Geography__c')
