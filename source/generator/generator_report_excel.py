import pandas as pd
import os
from .util import *
from generator.query_builder import *
from generator.sql_builder import *
import boto3
from botocore.client import Config
from config.db import *


def generate_report_campaign_excel(data: dict): 
    es = ElastichConnection()
    collection = MongoConnecction()

    match_id = collection.find_one({"_id": data.get("id")})

    data = fetch_all_query(es, data)

    polda = data.get("polda")
    polsek = data.get("polsek")
    polres = data.get("polres")
    
    try:
        time = datetime.now().strftime('%Y-%m-%d')
        file_output = f"result/campaign_report_{time}.xlsx"
        file_name = f"campaign_report_{time}.xlsx"

        writer = pd.ExcelWriter(file_output, engine='xlsxwriter')
        polda.to_excel(writer, sheet_name='Polda', index=False)
        polres.to_excel(writer, sheet_name='Polres', index=False)
        polsek.to_excel(writer, sheet_name='Polsek', index=False)
        writer.close()

        s3_client = boto3.client(
            endpoint_url = os.getenv("S3_VORTEX_ADDRESS"),
            service_name = 's3',
            aws_access_key_id= os.getenv("S3_VORTEX_ACCESS_KEY"),
            aws_secret_access_key=os.getenv("S3_VORTEX_SECRET_KEY"),
            config=Config(connect_timeout=90, retries={'max_attempts': 6})
        )

        bucket = os.getenv("S3_VORTEX_BUCKET_NETWORK")
        s3_path = f"report/campaign/{file_name}"    
        update_query = {"$set": {"status": "success", "s3Path": f"s3://campaign-management/report/campaign/{file_name}"}}
        collection.update_one(match_id, update_query)
        s3_client.upload_file(file_output , bucket, s3_path)
        
        if os.path.exists(file_output):
            os.remove(file_output)
        else:
            print("File not exist")

        return f"campaign_report_{time}.xlsx"
    
    except Exception as e:
        print("error when generating report: ", e)

def generate_report_kpi_excel(data: dict):
    collection = MongoConnecction()
    conn = PostgreConnection()

    match_id = collection.find_one({"_id": data.get("id")})

    data = fetch_sql_query(conn, data)

    polda = data.get("polda")
    polres = data.get("polres")
    polsek = data.get("polsek")

    try:
        time = datetime.now().strftime('%Y-%m-%d')
        file_output = f"result/kpi_report_{time}.xlsx"
        file_name = f"kpi_report_{time}.xlsx"

        writer = pd.ExcelWriter(file_output, engine='xlsxwriter')
        polda.to_excel(writer, sheet_name='Polda', index=False)
        polres.to_excel(writer, sheet_name='Polres', index=False)
        polsek.to_excel(writer, sheet_name='Polsek', index=False)
        writer.close()

        bucket = os.getenv("S3_VORTEX_BUCKET_NETWORK")
        s3_client = boto3.client(
            endpoint_url = os.getenv("S3_VORTEX_ADDRESS"),
            service_name = 's3',
            aws_access_key_id= os.getenv("S3_VORTEX_ACCESS_KEY"),
            aws_secret_access_key=os.getenv("S3_VORTEX_SECRET_KEY"),
            config=Config(connect_timeout=90, retries={'max_attempts': 6})
        )

        s3_path = f"report/kpi/{file_name}"    
        update_query = {"$set": {"status": "success", "s3Path": f"s3://campaign-management/report/kpi/{file_name}"}}
        collection.update_one(match_id, update_query)
        s3_client.upload_file(file_output , bucket, s3_path)

        if os.path.exists(file_output):
            os.remove(file_output)
        else:
            print("File not exist")

        return f"kpi_report_{time}.xlsx"
    except Exception as e:
        print("error when generating report", e)
        