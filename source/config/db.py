import os
from elasticsearch import Elasticsearch
import pymongo
import psycopg2

def ElastichConnection():
    username = os.getenv("ES_VORTEX_USER")
    password = os.getenv("ES_VORTEX_PASS")
    uri = os.getenv("ES_VORTEX_URI")
    es = Elasticsearch(uri, http_auth=(username, password))

    return es

def MongoConnecction():
    srv = os.getenv("DB_MONGO_SRV")
    db_name = os.getenv("DB_NAME")

    mongo_client = pymongo.MongoClient(srv)
    mongo_index = mongo_client[db_name]
    collection = mongo_index['report']

    return collection

def PostgreConnection():
    conn = psycopg2.connect(
    host=os.getenv("POSTGRE_HOST"),
    database=os.getenv("POSTGRE_DB"),
    user=os.getenv("POSTGRE_USER"),
    password=os.getenv("POSTGRE_PASS"),
    port=os.getenv("POSTGRE_PORT")
    )

    return conn