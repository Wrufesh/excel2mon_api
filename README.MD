# class GetDataFromExcel(*filename*)

## Attributes of GetDataFromExcel

* *collections* : This gives the name of tables present in excel document
* *coll_field_dict* : This gives the dictionary of all the tables and their attribute
* *get_collections_names()* : returns same result as given by collection attribute
* *get_collection_fields()*: returns the list of fields present in particular collection


# Usages

## Create the class instance and use it methods 
* `data = GetDataFromExcel('model.xlsx')  # Instance creation`
* `data.collections                       # Gives collections list`
* `data.get_collection_fields('Account')   # Returns the fields present in Account table`          

## Using with mongoDB
* `connection = MongoClient('mongodb://localhost:27017/')`
* `db = connection['db_name']`
* `x.write_list_to_coll(db, 'Geography__c', [1,2,3,4,5,6,7,8,9,10,11,12,13])`
* `x.write_from_input(db, 'Geography__c')`
