import pymongo
from pymongo import MongoClient
mongo_uri = "mongodb://root:geleiinfo@8.140.25.143:27017/test?authMechanism=SCRAM-SHA-256"
mongodb_handler = MongoDBHandler(mongo_uri)
collection_name = "ppt_test"
collection = mongodb_handler.create_or_get_collection(collection_name)