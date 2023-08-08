import pymongo
from pymongo import MongoClient
import  json

class MongoDBHandler:
    def __init__(self, uri):
        self.client = MongoClient(uri)
        self.db = self.client.get_database()

    def create_or_get_collection(self, collection_name):
        if collection_name in self.db.list_collection_names():
            print(f"Collection '{collection_name}' already exists.")
        else:
            self.db.create_collection(collection_name)
            print(f"Collection '{collection_name}' created successfully!")

        return self.db[collection_name]

    def get_next_sequence(self, sequence_name):
        sequence = self.db[sequence_name].find_one_and_update(
            {'_id': sequence_name},
            {'$inc': {'seq': 1}},
            upsert=True,
            return_document=pymongo.ReturnDocument.AFTER
        )
        return sequence['seq']

    def insert_document(self, collection, document):
        collection.insert_one(document)

    def close_connection(self):
        self.client.close()


def main():
    mongo_uri = "mongodb://root:geleiinfo@8.140.25.143:27017/test?authMechanism=SCRAM-SHA-256"
    mongodb_handler = MongoDBHandler(mongo_uri)

    collection_name = "ppt_test"
    collection = mongodb_handler.create_or_get_collection(collection_name)
    with open(rf'F:\pptx2md\json\大气工作总结计划汇报PPT模板.json', 'r',encoding='utf-8') as file:
        json_string = file.read()

    document = {
        '_id': mongodb_handler.get_next_sequence('ppt_id'),
        'json_field': json.loads(json_string),
        'middle_id': 'iddle_id1',
        'backup_field1': 'value1',
        'backup_field2': 'value2',
        'backup_field3': 'value3'
    }

    mongodb_handler.insert_document(collection, document)
    mongodb_handler.close_connection()


if __name__ == "__main__":
    main()
