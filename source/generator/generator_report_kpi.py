from datetime import datetime
import os
import pymongo

from pptx import Presentation
from python_pptx_text_replacer import TextReplacer
from .util import *
from generator.query_builder import fetch_kpi_query
from generator.sql_builder import fetch_sql_query
from generator.query_builder import ReportParams
from elasticsearch import Elasticsearch
from dotenv import load_dotenv
import boto3
from botocore.client import Config
import psycopg2
from config.db import *

load_dotenv()

def generate_report_kpi(data: dict):
    es = ElastichConnection()
    collection = MongoConnecction()
    conn = PostgreConnection()

    data_elastich = fetch_kpi_query(es, data)
    tren_eksposure_matrick = data_elastich.get("tren_eksposure_matrick")
    tren_eksposure_total = data_elastich.get("tren_eksposure_total")
    tren_eksposure_trendline = data_elastich.get("tren_eksposure_trendline")

    data_sql = fetch_sql_query(conn, data)
    key_performance_indicator = data_sql.get("key_performance_indicator")
    digital_platform_indicator = data_sql.get("digital_platform_indicator")
    digital_interaction_indicator = data_sql.get("digital_interaction_indicator")
    public_perception_indicator = data_sql.get("public_perception_indicator")

    match_id = collection.find_one({"_id": data.get("id")})

    try:
        print('=============== Generating ReportKPI ===============')
        prs = Presentation("template/tmp-REPORT_KPI.pptx")
        find_chart(prs)

        generate_chart(prs, tren_eksposure_trendline, 1, 0)

        save_file = ("result/REPORT_KPI_chart.pptx")
        prs.save(save_file)

        platform_post = retrieve_data_object(tren_eksposure_matrick, sub_key='doc_count', length_min=5, default_fill=0)
        formatted_numbers = [f"{num:,}".replace(",", ".") for num in platform_post]

        platform_name = retrieve_data_object(tren_eksposure_matrick, sub_key='key', length_min=5)
        total_exsposure = tren_eksposure_total.get('eksposure')
        total_engagement = tren_eksposure_total.get('engagement')

        formatted_exposure = f"{total_exsposure:,.2f}"
        formatted_engagement = f"{total_engagement:,.2f}"

        formatted_exposure = formatted_exposure.replace(",", "X").replace(".", ",").replace("X", ".")
        formatted_engagement = formatted_engagement.replace(",", "X").replace(".", ",").replace("X", ".")

        grafik_date = []
        sorted_date = sorted(tren_eksposure_trendline, key=lambda x: -x['doc_count'], reverse=False)
        for item in sorted_date:
            dt = datetime.fromtimestamp(item['key'] / 1000).strftime("%d %B %Y")
            grafik_date.append(dt)

        grafik_value = []
        sorted_value = sorted(tren_eksposure_trendline, key=lambda x: -x['doc_count'], reverse=False)
        for item in sorted_value:
            grafik_value.append(item['doc_count'])
        formatted_grafik = [f"{num:,}".replace(",", ".") for num in grafik_value]

        kpi_polda_name = retrieve_data_object(key_performance_indicator, sub_key='key', length_min=34)
        kpi_value = retrieve_data_object(key_performance_indicator, sub_key='kpi', length_min=34, default_fill=0)
        kpi_dpi_value = retrieve_data_object(key_performance_indicator, sub_key='dpi', length_min=34, default_fill=0)
        kpi_dii_value = retrieve_data_object(key_performance_indicator, sub_key='dii', length_min=34, default_fill=0)
        kpi_ppi_value = retrieve_data_object(key_performance_indicator, sub_key='ppi', length_min=34, default_fill=0)

        dpi_polda_name = retrieve_data_object(digital_platform_indicator, sub_key='key', length_min=34)
        dpi_exsposure = retrieve_data_object(digital_platform_indicator, sub_key='exsposure', length_min=34, default_fill=0)
        dpi_engagement = retrieve_data_object(digital_platform_indicator, sub_key='engagemnet', length_min=34, default_fill=0)
        dpi_value = retrieve_data_object(digital_platform_indicator, sub_key='dpi', length_min=34, default_fill=0)

        dii_polda_name = retrieve_data_object(digital_interaction_indicator, sub_key='key', length_min=34)
        dii_exsposure = retrieve_data_object(digital_interaction_indicator, sub_key='exsposure', length_min=34, default_fill=0)
        dii_engagement = retrieve_data_object(digital_interaction_indicator, sub_key='engagemnet', length_min=34, default_fill=0)
        dii_sentiment_exsposure = retrieve_data_object(digital_interaction_indicator, sub_key='sentiment_exsposure', length_min=34, default_fill=0)
        dii_sentiment_engagement = retrieve_data_object(digital_interaction_indicator, sub_key='sentiment_engagement', length_min=34, default_fill=0)
        dii_value = retrieve_data_object(digital_interaction_indicator, sub_key='dii', length_min=34, default_fill=0)

        ppi_polda_name = retrieve_data_object(public_perception_indicator, sub_key='key', length_min=34)
        ppi_exsposure = retrieve_data_object(public_perception_indicator, sub_key='exsposure', length_min=34, default_fill=0)
        ppi_engagement = retrieve_data_object(public_perception_indicator, sub_key='engagemnet', length_min=34, default_fill=0)
        ppi_sentiment_exsposure = retrieve_data_object(public_perception_indicator, sub_key='sentiment_exsposure', length_min=34, default_fill=0)
        ppi_sentiment_engagement = retrieve_data_object(public_perception_indicator, sub_key='sentiment_engagement', length_min=34, default_fill=0)
        ppi_value = retrieve_data_object(public_perception_indicator, sub_key='ppi', length_min=34, default_fill=0)

        from_time = data["timeframe"].get("from")
        to_time = data["timeframe"].get("to")

        bulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
        from_time = datetime.utcfromtimestamp(from_time / 1000)
        to_time = datetime.utcfromtimestamp(to_time / 1000)
        date = f"{from_time.day} {bulan[from_time.month - 1]} {from_time.year} - {to_time.day} {bulan[to_time.month - 1]} {to_time.year}"

        replacer = TextReplacer(save_file, slides='', tables=True, charts=False, textframes=True)
        replacer.replace_text([
            ("{date}", date),

            # platform
            ("{insta}", str(formatted_numbers[0])),
            ("{twitter}", str(formatted_numbers[1])),
            ("{facebook}", str(formatted_numbers[2])),
            ("{tiktok}", str(formatted_numbers[3])),
            ("{youtube}", str(formatted_numbers[4])),

            # total exsposure & engagement
            ("{v_exposure}", str(formatted_exposure)),
            ("{v_engagement}", str(formatted_engagement)),

            # summary
            ("{rank_1}", str(platform_name[0])),
            ("{rank_2}", str(platform_name[1])),
            ("{rank_3}", str(platform_name[2])),
            ("{rank_4}", str(platform_name[3])),
            ("{rank_5}", str(platform_name[4])),

            ("{v_rank_1}", str(formatted_numbers[0])),  
            ("{v_rank_2}", str(formatted_numbers[1])),
            ("{v_rank_3}", str(formatted_numbers[2])),
            ("{v_rank_4}", str(formatted_numbers[3])),
            ("{v_rank_5}", str(formatted_numbers[4])),

            ("{trendline_1}", str(grafik_date[0])),
            ("{trendline_2}", str(grafik_date[1])),
            ("{trendline_3}", str(grafik_date[2])),
            ("{trendline_4}", str(grafik_date[3])),
            ("{trendline_5}", str(grafik_date[4])),

            ("{v_trendline_1}", str(formatted_grafik[0])),
            ("{v_trendline_2}", str(formatted_grafik[1])),
            ("{v_trendline_3}", str(formatted_grafik[2])),
            ("{v_trendline_4}", str(formatted_grafik[3])),
            ("{v_trendline_5}", str(formatted_grafik[4])),

            ("{v_trendline_last}", str(formatted_grafik[-1])),
            ("{trendline_last}", str(grafik_date[-1])),

            # kpi section
            ("{kpi_rank_1}", str(kpi_polda_name[0])),
            ("{kpi_rank_2}", str(kpi_polda_name[1])),
            ("{kpi_rank_3}", str(kpi_polda_name[2])),
            ("{kpi_rank_4}", str(kpi_polda_name[3])),
            ("{kpi_rank_5}", str(kpi_polda_name[4])),
            ("{kpi_rank_6}", str(kpi_polda_name[5])),
            ("{kpi_rank_7}", str(kpi_polda_name[6])),
            ("{kpi_rank_8}", str(kpi_polda_name[7])),
            ("{kpi_rank_9}", str(kpi_polda_name[8])),
            ("{kpi_rank_10}", str(kpi_polda_name[9])),
            ("{kpi_rank_11}", str(kpi_polda_name[10])),
            ("{kpi_rank_12}", str(kpi_polda_name[11])),
            ("{kpi_rank_13}", str(kpi_polda_name[12])),
            ("{kpi_rank_14}", str(kpi_polda_name[13])),
            ("{kpi_rank_15}", str(kpi_polda_name[14])),
            ("{kpi_rank_16}", str(kpi_polda_name[15])),
            ("{kpi_rank_17}", str(kpi_polda_name[16])),
            ("{kpi_rank_18}", str(kpi_polda_name[17])),
            ("{kpi_rank_19}", str(kpi_polda_name[18])),
            ("{kpi_rank_20}", str(kpi_polda_name[19])),
            ("{kpi_rank_21}", str(kpi_polda_name[20])),
            ("{kpi_rank_22}", str(kpi_polda_name[21])),
            ("{kpi_rank_23}", str(kpi_polda_name[22])),
            ("{kpi_rank_24}", str(kpi_polda_name[23])),
            ("{kpi_rank_25}", str(kpi_polda_name[24])),
            ("{kpi_rank_26}", str(kpi_polda_name[25])),
            ("{kpi_rank_27}", str(kpi_polda_name[26])),
            ("{kpi_rank_28}", str(kpi_polda_name[27])),
            ("{kpi_rank_29}", str(kpi_polda_name[28])),
            ("{kpi_rank_30}", str(kpi_polda_name[29])),
            ("{kpi_rank_31}", str(kpi_polda_name[30])),
            ("{kpi_rank_32}", str(kpi_polda_name[31])),
            ("{kpi_rank_33}", str(kpi_polda_name[32])),
            ("{kpi_rank_34}", str(kpi_polda_name[33])),

            # kpi value
            ("{kkpi_v_1}", str(kpi_value[0])),
            ("{kkpi_v_2}", str(kpi_value[1])),
            ("{kkpi_v_3}", str(kpi_value[2])),
            ("{kkpi_v_4}", str(kpi_value[3])),
            ("{kkpi_v_5}", str(kpi_value[4])),
            ("{kkpi_v_6}", str(kpi_value[5])),
            ("{kkpi_v_7}", str(kpi_value[6])),
            ("{kkpi_v_8}", str(kpi_value[7])),
            ("{kkpi_v_9}", str(kpi_value[8])),
            ("{kkpi_v_10}", str(kpi_value[9])),
            ("{kkpi_v_11}", str(kpi_value[10])),
            ("{kkpi_v_12}", str(kpi_value[11])),
            ("{kkpi_v_13}", str(kpi_value[12])),
            ("{kkpi_v_14}", str(kpi_value[13])),
            ("{kkpi_v_15}", str(kpi_value[14])),
            ("{kkpi_v_16}", str(kpi_value[15])),
            ("{kkpi_v_17}", str(kpi_value[16])),
            ("{kkpi_v_18}", str(kpi_value[17])),
            ("{kkpi_v_19}", str(kpi_value[18])),
            ("{kkpi_v_20}", str(kpi_value[19])),
            ("{kkpi_v_21}", str(kpi_value[20])),
            ("{kkpi_v_22}", str(kpi_value[21])),
            ("{kkpi_v_23}", str(kpi_value[22])),
            ("{kkpi_v_24}", str(kpi_value[23])),
            ("{kkpi_v_25}", str(kpi_value[24])),
            ("{kkpi_v_26}", str(kpi_value[25])),
            ("{kkpi_v_27}", str(kpi_value[26])),
            ("{kkpi_v_28}", str(kpi_value[27])),
            ("{kkpi_v_29}", str(kpi_value[28])),
            ("{kkpi_v_30}", str(kpi_value[29])),
            ("{kkpi_v_31}", str(kpi_value[30])),
            ("{kkpi_v_32}", str(kpi_value[31])),
            ("{kkpi_v_33}", str(kpi_value[32])),
            ("{kkpi_v_34}", str(kpi_value[33])),
            
            # kpi dpi
            ("{kdpi_v_1}", str(kpi_dpi_value[0])),
            ("{kdpi_v_2}", str(kpi_dpi_value[1])),
            ("{kdpi_v_3}", str(kpi_dpi_value[2])),
            ("{kdpi_v_4}", str(kpi_dpi_value[3])),
            ("{kdpi_v_5}", str(kpi_dpi_value[4])),
            ("{kdpi_v_6}", str(kpi_dpi_value[5])),
            ("{kdpi_v_7}", str(kpi_dpi_value[6])),
            ("{kdpi_v_8}", str(kpi_dpi_value[7])),
            ("{kdpi_v_9}", str(kpi_dpi_value[8])),
            ("{kdpi_v_10}", str(kpi_dpi_value[9])),
            ("{kdpi_v_11}", str(kpi_dpi_value[10])),
            ("{kdpi_v_12}", str(kpi_dpi_value[11])),
            ("{kdpi_v_13}", str(kpi_dpi_value[12])),
            ("{kdpi_v_14}", str(kpi_dpi_value[13])),
            ("{kdpi_v_15}", str(kpi_dpi_value[14])),
            ("{kdpi_v_16}", str(kpi_dpi_value[15])),
            ("{kdpi_v_17}", str(kpi_dpi_value[16])),
            ("{kdpi_v_18}", str(kpi_dpi_value[17])),
            ("{kdpi_v_19}", str(kpi_dpi_value[18])),
            ("{kdpi_v_20}", str(kpi_dpi_value[19])),
            ("{kdpi_v_21}", str(kpi_dpi_value[20])),
            ("{kdpi_v_22}", str(kpi_dpi_value[21])),
            ("{kdpi_v_23}", str(kpi_dpi_value[22])),
            ("{kdpi_v_24}", str(kpi_dpi_value[23])),
            ("{kdpi_v_25}", str(kpi_dpi_value[24])),
            ("{kdpi_v_26}", str(kpi_dpi_value[25])),
            ("{kdpi_v_27}", str(kpi_dpi_value[26])),
            ("{kdpi_v_28}", str(kpi_dpi_value[27])),
            ("{kdpi_v_29}", str(kpi_dpi_value[28])),
            ("{kdpi_v_30}", str(kpi_dpi_value[29])),
            ("{kdpi_v_31}", str(kpi_dpi_value[30])),
            ("{kdpi_v_32}", str(kpi_dpi_value[31])),
            ("{kdpi_v_33}", str(kpi_dpi_value[32])),
            ("{kdpi_v_34}", str(kpi_dpi_value[33])),

            # kpi dii
            ("{kdii_v_1}", str(kpi_dii_value[0])),
            ("{kdii_v_2}", str(kpi_dii_value[1])),
            ("{kdii_v_3}", str(kpi_dii_value[2])),
            ("{kdii_v_4}", str(kpi_dii_value[3])),
            ("{kdii_v_5}", str(kpi_dii_value[4])),
            ("{kdii_v_6}", str(kpi_dii_value[5])),
            ("{kdii_v_7}", str(kpi_dii_value[6])),
            ("{kdii_v_8}", str(kpi_dii_value[7])),
            ("{kdii_v_9}", str(kpi_dii_value[8])),
            ("{kdii_v_10}", str(kpi_dii_value[9])),
            ("{kdii_v_11}", str(kpi_dii_value[10])),
            ("{kdii_v_12}", str(kpi_dii_value[11])),
            ("{kdii_v_13}", str(kpi_dii_value[12])),
            ("{kdii_v_14}", str(kpi_dii_value[13])),
            ("{kdii_v_15}", str(kpi_dii_value[14])),
            ("{kdii_v_16}", str(kpi_dii_value[15])),
            ("{kdii_v_17}", str(kpi_dii_value[16])),
            ("{kdii_v_18}", str(kpi_dii_value[17])),
            ("{kdii_v_19}", str(kpi_dii_value[18])),
            ("{kdii_v_20}", str(kpi_dii_value[19])),
            ("{kdii_v_21}", str(kpi_dii_value[20])),
            ("{kdii_v_22}", str(kpi_dii_value[21])),
            ("{kdii_v_23}", str(kpi_dii_value[22])),
            ("{kdii_v_24}", str(kpi_dii_value[23])),
            ("{kdii_v_25}", str(kpi_dii_value[24])),
            ("{kdii_v_26}", str(kpi_dii_value[25])),
            ("{kdii_v_27}", str(kpi_dii_value[26])),
            ("{kdii_v_28}", str(kpi_dii_value[27])),
            ("{kdii_v_29}", str(kpi_dii_value[28])),
            ("{kdii_v_30}", str(kpi_dii_value[29])),
            ("{kdii_v_31}", str(kpi_dii_value[30])),
            ("{kdii_v_32}", str(kpi_dii_value[31])),
            ("{kdii_v_33}", str(kpi_dii_value[32])),
            ("{kdii_v_34}", str(kpi_dii_value[33])),

            # kpi ppi
            ("{kppi_v_1}", str(kpi_ppi_value[0])),
            ("{kppi_v_2}", str(kpi_ppi_value[1])),
            ("{kppi_v_3}", str(kpi_ppi_value[2])),
            ("{kppi_v_4}", str(kpi_ppi_value[3])),
            ("{kppi_v_5}", str(kpi_ppi_value[4])),
            ("{kppi_v_6}", str(kpi_ppi_value[5])),
            ("{kppi_v_7}", str(kpi_ppi_value[6])),
            ("{kppi_v_8}", str(kpi_ppi_value[7])),
            ("{kppi_v_9}", str(kpi_ppi_value[8])),
            ("{kppi_v_10}", str(kpi_ppi_value[9])),
            ("{kppi_v_11}", str(kpi_ppi_value[10])),
            ("{kppi_v_12}", str(kpi_ppi_value[11])),
            ("{kppi_v_13}", str(kpi_ppi_value[12])),
            ("{kppi_v_14}", str(kpi_ppi_value[13])),
            ("{kppi_v_15}", str(kpi_ppi_value[14])),
            ("{kppi_v_16}", str(kpi_ppi_value[15])),
            ("{kppi_v_17}", str(kpi_ppi_value[16])),
            ("{kppi_v_18}", str(kpi_ppi_value[17])),
            ("{kppi_v_19}", str(kpi_ppi_value[18])),
            ("{kppi_v_20}", str(kpi_ppi_value[19])),
            ("{kppi_v_21}", str(kpi_ppi_value[20])),
            ("{kppi_v_22}", str(kpi_ppi_value[21])),
            ("{kppi_v_23}", str(kpi_ppi_value[22])),
            ("{kppi_v_24}", str(kpi_ppi_value[23])),
            ("{kppi_v_25}", str(kpi_ppi_value[24])),
            ("{kppi_v_26}", str(kpi_ppi_value[25])),
            ("{kppi_v_27}", str(kpi_ppi_value[26])),
            ("{kppi_v_28}", str(kpi_ppi_value[27])),
            ("{kppi_v_29}", str(kpi_ppi_value[28])),
            ("{kppi_v_30}", str(kpi_ppi_value[29])),
            ("{kppi_v_31}", str(kpi_ppi_value[30])),
            ("{kppi_v_32}", str(kpi_ppi_value[31])),
            ("{kppi_v_33}", str(kpi_ppi_value[32])),
            ("{kppi_v_34}", str(kpi_ppi_value[33])),

            # dpi section
            ("{dpi_rank_1}", str(dpi_polda_name[0])),
            ("{dpi_rank_2}", str(dpi_polda_name[1])),
            ("{dpi_rank_3}", str(dpi_polda_name[2])),
            ("{dpi_rank_4}", str(dpi_polda_name[3])),
            ("{dpi_rank_5}", str(dpi_polda_name[4])),
            ("{dpi_rank_6}", str(dpi_polda_name[5])),
            ("{dpi_rank_7}", str(dpi_polda_name[6])),
            ("{dpi_rank_8}", str(dpi_polda_name[7])),
            ("{dpi_rank_9}", str(dpi_polda_name[8])),
            ("{dpi_rank_10}", str(dpi_polda_name[9])),
            ("{dpi_rank_11}", str(dpi_polda_name[10])),
            ("{dpi_rank_12}", str(dpi_polda_name[11])),
            ("{dpi_rank_13}", str(dpi_polda_name[12])),
            ("{dpi_rank_14}", str(dpi_polda_name[13])),
            ("{dpi_rank_15}", str(dpi_polda_name[14])),
            ("{dpi_rank_16}", str(dpi_polda_name[15])),
            ("{dpi_rank_17}", str(dpi_polda_name[16])),
            ("{dpi_rank_18}", str(dpi_polda_name[17])),
            ("{dpi_rank_19}", str(dpi_polda_name[18])),
            ("{dpi_rank_20}", str(dpi_polda_name[19])),
            ("{dpi_rank_21}", str(dpi_polda_name[20])),
            ("{dpi_rank_22}", str(dpi_polda_name[21])),
            ("{dpi_rank_23}", str(dpi_polda_name[22])),
            ("{dpi_rank_24}", str(dpi_polda_name[23])),
            ("{dpi_rank_25}", str(dpi_polda_name[24])),
            ("{dpi_rank_26}", str(dpi_polda_name[25])),
            ("{dpi_rank_27}", str(dpi_polda_name[26])),
            ("{dpi_rank_28}", str(dpi_polda_name[27])),
            ("{dpi_rank_29}", str(dpi_polda_name[28])),
            ("{dpi_rank_30}", str(dpi_polda_name[29])),
            ("{dpi_rank_31}", str(dpi_polda_name[30])),
            ("{dpi_rank_32}", str(dpi_polda_name[31])),
            ("{dpi_rank_33}", str(dpi_polda_name[32])),
            ("{dpi_rank_34}", str(dpi_polda_name[33])),

            # dpi exsposure
            ("{dpi_exp_1}", str(dpi_exsposure[0])),
            ("{dpi_exp_2}", str(dpi_exsposure[1])),
            ("{dpi_exp_3}", str(dpi_exsposure[2])),
            ("{dpi_exp_4}", str(dpi_exsposure[3])),
            ("{dpi_exp_5}", str(dpi_exsposure[4])),
            ("{dpi_exp_6}", str(dpi_exsposure[5])),
            ("{dpi_exp_7}", str(dpi_exsposure[6])),
            ("{dpi_exp_8}", str(dpi_exsposure[7])),
            ("{dpi_exp_9}", str(dpi_exsposure[8])),
            ("{dpi_exp_10}", str(dpi_exsposure[9])),
            ("{dpi_exp_11}", str(dpi_exsposure[10])),
            ("{dpi_exp_12}", str(dpi_exsposure[11])),
            ("{dpi_exp_13}", str(dpi_exsposure[12])),
            ("{dpi_exp_14}", str(dpi_exsposure[13])),
            ("{dpi_exp_15}", str(dpi_exsposure[14])),
            ("{dpi_exp_16}", str(dpi_exsposure[15])),
            ("{dpi_exp_17}", str(dpi_exsposure[16])),
            ("{dpi_exp_18}", str(dpi_exsposure[17])),
            ("{dpi_exp_19}", str(dpi_exsposure[18])),
            ("{dpi_exp_20}", str(dpi_exsposure[19])),
            ("{dpi_exp_21}", str(dpi_exsposure[20])),
            ("{dpi_exp_22}", str(dpi_exsposure[21])),
            ("{dpi_exp_23}", str(dpi_exsposure[22])),
            ("{dpi_exp_24}", str(dpi_exsposure[23])),
            ("{dpi_exp_25}", str(dpi_exsposure[24])),
            ("{dpi_exp_26}", str(dpi_exsposure[25])),
            ("{dpi_exp_27}", str(dpi_exsposure[26])),
            ("{dpi_exp_28}", str(dpi_exsposure[27])),
            ("{dpi_exp_29}", str(dpi_exsposure[28])),
            ("{dpi_exp_30}", str(dpi_exsposure[29])),
            ("{dpi_exp_31}", str(dpi_exsposure[30])),
            ("{dpi_exp_32}", str(dpi_exsposure[31])),
            ("{dpi_exp_33}", str(dpi_exsposure[32])),
            ("{dpi_exp_34}", str(dpi_exsposure[33])),

            # dpi engagement
            ("{dpi_eng_1}", str(dpi_engagement[0])),
            ("{dpi_eng_2}", str(dpi_engagement[1])),
            ("{dpi_eng_3}", str(dpi_engagement[2])),
            ("{dpi_eng_4}", str(dpi_engagement[3])),
            ("{dpi_eng_5}", str(dpi_engagement[4])),
            ("{dpi_eng_6}", str(dpi_engagement[5])),
            ("{dpi_eng_7}", str(dpi_engagement[6])),
            ("{dpi_eng_8}", str(dpi_engagement[7])),
            ("{dpi_eng_9}", str(dpi_engagement[8])),
            ("{dpi_eng_10}", str(dpi_engagement[9])),
            ("{dpi_eng_11}", str(dpi_engagement[10])),
            ("{dpi_eng_12}", str(dpi_engagement[11])),
            ("{dpi_eng_13}", str(dpi_engagement[12])),
            ("{dpi_eng_14}", str(dpi_engagement[13])),
            ("{dpi_eng_15}", str(dpi_engagement[14])),
            ("{dpi_eng_16}", str(dpi_engagement[15])),
            ("{dpi_eng_17}", str(dpi_engagement[16])),
            ("{dpi_eng_18}", str(dpi_engagement[17])),
            ("{dpi_eng_19}", str(dpi_engagement[18])),
            ("{dpi_eng_20}", str(dpi_engagement[19])),
            ("{dpi_eng_21}", str(dpi_engagement[20])),
            ("{dpi_eng_22}", str(dpi_engagement[21])),
            ("{dpi_eng_23}", str(dpi_engagement[22])),
            ("{dpi_eng_24}", str(dpi_engagement[23])),
            ("{dpi_eng_25}", str(dpi_engagement[24])),
            ("{dpi_eng_26}", str(dpi_engagement[25])),
            ("{dpi_eng_27}", str(dpi_engagement[26])),
            ("{dpi_eng_28}", str(dpi_engagement[27])),
            ("{dpi_eng_29}", str(dpi_engagement[28])),
            ("{dpi_eng_30}", str(dpi_engagement[29])),
            ("{dpi_eng_31}", str(dpi_engagement[30])),
            ("{dpi_eng_32}", str(dpi_engagement[31])),
            ("{dpi_eng_33}", str(dpi_engagement[32])),
            ("{dpi_eng_34}", str(dpi_engagement[33])),
            
            #  dpi value
            ("{dpi_v1}", str(dpi_value[0])),
            ("{dpi_v2}", str(dpi_value[1])),
            ("{dpi_v3}", str(dpi_value[2])),
            ("{dpi_v4}", str(dpi_value[3])),
            ("{dpi_v5}", str(dpi_value[4])),
            ("{dpi_v6}", str(dpi_value[5])),
            ("{dpi_v7}", str(dpi_value[6])),
            ("{dpi_v8}", str(dpi_value[7])),
            ("{dpi_v9}", str(dpi_value[8])),
            ("{dpi_v10}", str(dpi_value[9])),
            ("{dpi_v11}", str(dpi_value[10])),
            ("{dpi_v12}", str(dpi_value[11])),
            ("{dpi_v13}", str(dpi_value[12])),
            ("{dpi_v14}", str(dpi_value[13])),
            ("{dpi_v15}", str(dpi_value[14])),
            ("{dpi_v16}", str(dpi_value[15])),
            ("{dpi_v17}", str(dpi_value[16])),
            ("{dpi_v18}", str(dpi_value[17])),
            ("{dpi_v19}", str(dpi_value[18])),
            ("{dpi_v20}", str(dpi_value[19])),
            ("{dpi_v21}", str(dpi_value[20])),
            ("{dpi_v22}", str(dpi_value[21])),
            ("{dpi_v23}", str(dpi_value[22])),
            ("{dpi_v24}", str(dpi_value[23])),
            ("{dpi_v25}", str(dpi_value[24])),
            ("{dpi_v26}", str(dpi_value[25])),
            ("{dpi_v27}", str(dpi_value[26])),
            ("{dpi_v28}", str(dpi_value[27])),
            ("{dpi_v29}", str(dpi_value[28])),
            ("{dpi_v30}", str(dpi_value[29])),
            ("{dpi_v31}", str(dpi_value[30])),
            ("{dpi_v32}", str(dpi_value[31])),
            ("{dpi_v33}", str(dpi_value[32])),
            ("{dpi_v34}", str(dpi_value[33])),

            # dii section
            ("{dii_rank_1}", str(dii_polda_name[0])),
            ("{dii_rank_2}", str(dii_polda_name[1])),
            ("{dii_rank_3}", str(dii_polda_name[2])),
            ("{dii_rank_4}", str(dii_polda_name[3])),
            ("{dii_rank_5}", str(dii_polda_name[4])),
            ("{dii_rank_6}", str(dii_polda_name[5])),
            ("{dii_rank_7}", str(dii_polda_name[6])),
            ("{dii_rank_8}", str(dii_polda_name[7])),
            ("{dii_rank_9}", str(dii_polda_name[8])),
            ("{dii_rank_10}", str(dii_polda_name[9])),
            ("{dii_rank_11}", str(dii_polda_name[10])),
            ("{dii_rank_12}", str(dii_polda_name[11])),
            ("{dii_rank_13}", str(dii_polda_name[12])),
            ("{dii_rank_14}", str(dii_polda_name[13])),
            ("{dii_rank_15}", str(dii_polda_name[14])),
            ("{dii_rank_16}", str(dii_polda_name[15])),
            ("{dii_rank_17}", str(dii_polda_name[16])),
            ("{dii_rank_18}", str(dii_polda_name[17])),
            ("{dii_rank_19}", str(dii_polda_name[18])),
            ("{dii_rank_20}", str(dii_polda_name[19])),
            ("{dii_rank_21}", str(dii_polda_name[20])),
            ("{dii_rank_22}", str(dii_polda_name[21])),
            ("{dii_rank_23}", str(dii_polda_name[22])),
            ("{dii_rank_24}", str(dii_polda_name[23])),
            ("{dii_rank_25}", str(dii_polda_name[24])),
            ("{dii_rank_26}", str(dii_polda_name[25])),
            ("{dii_rank_27}", str(dii_polda_name[26])),
            ("{dii_rank_28}", str(dii_polda_name[27])),
            ("{dii_rank_29}", str(dii_polda_name[28])),
            ("{dii_rank_30}", str(dii_polda_name[29])),
            ("{dii_rank_31}", str(dii_polda_name[30])),
            ("{dii_rank_32}", str(dii_polda_name[31])),
            ("{dii_rank_33}", str(dii_polda_name[32])),
            ("{dii_rank_34}", str(dii_polda_name[33])),

            # dii exsposure
            ("{dii_exp_1}", str(dii_exsposure[0])),
            ("{dii_exp_2}", str(dii_exsposure[1])),
            ("{dii_exp_3}", str(dii_exsposure[2])),
            ("{dii_exp_4}", str(dii_exsposure[3])),
            ("{dii_exp_5}", str(dii_exsposure[4])),
            ("{dii_exp_6}", str(dii_exsposure[5])),
            ("{dii_exp_7}", str(dii_exsposure[6])),
            ("{dii_exp_8}", str(dii_exsposure[7])),
            ("{dii_exp_9}", str(dii_exsposure[8])),
            ("{dii_exp_10}", str(dii_exsposure[9])),
            ("{dii_exp_11}", str(dii_exsposure[10])),
            ("{dii_exp_12}", str(dii_exsposure[11])),
            ("{dii_exp_13}", str(dii_exsposure[12])),
            ("{dii_exp_14}", str(dii_exsposure[13])),
            ("{dii_exp_15}", str(dii_exsposure[14])),
            ("{dii_exp_16}", str(dii_exsposure[15])),
            ("{dii_exp_17}", str(dii_exsposure[16])),
            ("{dii_exp_18}", str(dii_exsposure[17])),
            ("{dii_exp_19}", str(dii_exsposure[18])),
            ("{dii_exp_20}", str(dii_exsposure[19])),
            ("{dii_exp_21}", str(dii_exsposure[20])),
            ("{dii_exp_22}", str(dii_exsposure[21])),
            ("{dii_exp_23}", str(dii_exsposure[22])),
            ("{dii_exp_24}", str(dii_exsposure[23])),
            ("{dii_exp_25}", str(dii_exsposure[24])),
            ("{dii_exp_26}", str(dii_exsposure[25])),
            ("{dii_exp_27}", str(dii_exsposure[26])),
            ("{dii_exp_28}", str(dii_exsposure[27])),
            ("{dii_exp_29}", str(dii_exsposure[28])),
            ("{dii_exp_30}", str(dii_exsposure[29])),
            ("{dii_exp_31}", str(dii_exsposure[30])),
            ("{dii_exp_32}", str(dii_exsposure[31])),
            ("{dii_exp_33}", str(dii_exsposure[32])),
            ("{dii_exp_34}", str(dii_exsposure[33])),

            # dii engagement
            ("{dii_eng_1}", str(dii_engagement[0])),
            ("{dii_eng_2}", str(dii_engagement[1])),
            ("{dii_eng_3}", str(dii_engagement[2])),
            ("{dii_eng_4}", str(dii_engagement[3])),
            ("{dii_eng_5}", str(dii_engagement[4])),
            ("{dii_eng_6}", str(dii_engagement[5])),
            ("{dii_eng_7}", str(dii_engagement[6])),
            ("{dii_eng_8}", str(dii_engagement[7])),
            ("{dii_eng_9}", str(dii_engagement[8])),
            ("{dii_eng_10}", str(dii_engagement[9])),
            ("{dii_eng_11}", str(dii_engagement[10])),
            ("{dii_eng_12}", str(dii_engagement[11])),
            ("{dii_eng_13}", str(dii_engagement[12])),
            ("{dii_eng_14}", str(dii_engagement[13])),
            ("{dii_eng_15}", str(dii_engagement[14])),
            ("{dii_eng_16}", str(dii_engagement[15])),
            ("{dii_eng_17}", str(dii_engagement[16])),
            ("{dii_eng_18}", str(dii_engagement[17])),
            ("{dii_eng_19}", str(dii_engagement[18])),
            ("{dii_eng_20}", str(dii_engagement[19])),
            ("{dii_eng_21}", str(dii_engagement[20])),
            ("{dii_eng_22}", str(dii_engagement[21])),
            ("{dii_eng_23}", str(dii_engagement[22])),
            ("{dii_eng_24}", str(dii_engagement[23])),
            ("{dii_eng_25}", str(dii_engagement[24])),
            ("{dii_eng_26}", str(dii_engagement[25])),
            ("{dii_eng_27}", str(dii_engagement[26])),
            ("{dii_eng_28}", str(dii_engagement[27])),
            ("{dii_eng_29}", str(dii_engagement[28])),
            ("{dii_eng_30}", str(dii_engagement[29])),
            ("{dii_eng_31}", str(dii_engagement[30])),
            ("{dii_eng_32}", str(dii_engagement[31])),
            ("{dii_eng_33}", str(dii_engagement[32])),
            ("{dii_eng_34}", str(dii_engagement[33])),

            # dii sentiment exsposure
            ("{dii_sexp_1}", str(dii_sentiment_exsposure[0])),
            ("{dii_sexp_2}", str(dii_sentiment_exsposure[1])),
            ("{dii_sexp_3}", str(dii_sentiment_exsposure[2])),
            ("{dii_sexp_4}", str(dii_sentiment_exsposure[3])),
            ("{dii_sexp_5}", str(dii_sentiment_exsposure[4])),
            ("{dii_sexp_6}", str(dii_sentiment_exsposure[5])),
            ("{dii_sexp_7}", str(dii_sentiment_exsposure[6])),
            ("{dii_sexp_8}", str(dii_sentiment_exsposure[7])),
            ("{dii_sexp_9}", str(dii_sentiment_exsposure[8])),
            ("{dii_sexp_10}", str(dii_sentiment_exsposure[9])),
            ("{dii_sexp_11}", str(dii_sentiment_exsposure[10])),
            ("{dii_sexp_12}", str(dii_sentiment_exsposure[11])),
            ("{dii_sexp_13}", str(dii_sentiment_exsposure[12])),
            ("{dii_sexp_14}", str(dii_sentiment_exsposure[13])),
            ("{dii_sexp_15}", str(dii_sentiment_exsposure[14])),
            ("{dii_sexp_16}", str(dii_sentiment_exsposure[15])),
            ("{dii_sexp_17}", str(dii_sentiment_exsposure[16])),
            ("{dii_sexp_18}", str(dii_sentiment_exsposure[17])),
            ("{dii_sexp_19}", str(dii_sentiment_exsposure[18])),
            ("{dii_sexp_20}", str(dii_sentiment_exsposure[19])),
            ("{dii_sexp_21}", str(dii_sentiment_exsposure[20])),
            ("{dii_sexp_22}", str(dii_sentiment_exsposure[21])),
            ("{dii_sexp_23}", str(dii_sentiment_exsposure[22])),
            ("{dii_sexp_24}", str(dii_sentiment_exsposure[23])),
            ("{dii_sexp_25}", str(dii_sentiment_exsposure[24])),
            ("{dii_sexp_26}", str(dii_sentiment_exsposure[25])),
            ("{dii_sexp_27}", str(dii_sentiment_exsposure[26])),
            ("{dii_sexp_28}", str(dii_sentiment_exsposure[27])),
            ("{dii_sexp_29}", str(dii_sentiment_exsposure[28])),
            ("{dii_sexp_30}", str(dii_sentiment_exsposure[29])),
            ("{dii_sexp_31}", str(dii_sentiment_exsposure[30])),
            ("{dii_sexp_32}", str(dii_sentiment_exsposure[31])),
            ("{dii_sexp_33}", str(dii_sentiment_exsposure[32])),
            ("{dii_sexp_34}", str(dii_sentiment_exsposure[33])),

            # dii sentiment engagement
            ("{dii_seng_1}", str(dii_sentiment_engagement[0])),
            ("{dii_seng_2}", str(dii_sentiment_engagement[1])),
            ("{dii_seng_3}", str(dii_sentiment_engagement[2])),
            ("{dii_seng_4}", str(dii_sentiment_engagement[3])),
            ("{dii_seng_5}", str(dii_sentiment_engagement[4])),
            ("{dii_seng_6}", str(dii_sentiment_engagement[5])),
            ("{dii_seng_7}", str(dii_sentiment_engagement[6])),
            ("{dii_seng_8}", str(dii_sentiment_engagement[7])),
            ("{dii_seng_9}", str(dii_sentiment_engagement[8])),
            ("{dii_seng_10}", str(dii_sentiment_engagement[9])),
            ("{dii_seng_11}", str(dii_sentiment_engagement[10])),
            ("{dii_seng_12}", str(dii_sentiment_engagement[11])),
            ("{dii_seng_13}", str(dii_sentiment_engagement[12])),
            ("{dii_seng_14}", str(dii_sentiment_engagement[13])),
            ("{dii_seng_15}", str(dii_sentiment_engagement[14])),
            ("{dii_seng_16}", str(dii_sentiment_engagement[15])),
            ("{dii_seng_17}", str(dii_sentiment_engagement[16])),
            ("{dii_seng_18}", str(dii_sentiment_engagement[17])),
            ("{dii_seng_19}", str(dii_sentiment_engagement[18])),
            ("{dii_seng_20}", str(dii_sentiment_engagement[19])),
            ("{dii_seng_21}", str(dii_sentiment_engagement[20])),
            ("{dii_seng_22}", str(dii_sentiment_engagement[21])),
            ("{dii_seng_23}", str(dii_sentiment_engagement[22])),
            ("{dii_seng_24}", str(dii_sentiment_engagement[23])),
            ("{dii_seng_25}", str(dii_sentiment_engagement[24])),
            ("{dii_seng_26}", str(dii_sentiment_engagement[25])),
            ("{dii_seng_27}", str(dii_sentiment_engagement[26])),
            ("{dii_seng_28}", str(dii_sentiment_engagement[27])),
            ("{dii_seng_29}", str(dii_sentiment_engagement[28])),
            ("{dii_seng_30}", str(dii_sentiment_engagement[29])),
            ("{dii_seng_31}", str(dii_sentiment_engagement[30])),
            ("{dii_seng_32}", str(dii_sentiment_engagement[31])),
            ("{dii_seng_33}", str(dii_sentiment_engagement[32])),
            ("{dii_seng_34}", str(dii_sentiment_engagement[33])),

            # dii value
            ("{dii_v1}", str(dii_value[0])),
            ("{dii_v2}", str(dii_value[1])),
            ("{dii_v3}", str(dii_value[2])),
            ("{dii_v4}", str(dii_value[3])),
            ("{dii_v5}", str(dii_value[4])),
            ("{dii_v6}", str(dii_value[5])),
            ("{dii_v7}", str(dii_value[6])),
            ("{dii_v8}", str(dii_value[7])),
            ("{dii_v9}", str(dii_value[8])),
            ("{dii_v10}", str(dii_value[9])),
            ("{dii_v11}", str(dii_value[10])),
            ("{dii_v12}", str(dii_value[11])),
            ("{dii_v13}", str(dii_value[12])),
            ("{dii_v14}", str(dii_value[13])),
            ("{dii_v15}", str(dii_value[14])),
            ("{dii_v16}", str(dii_value[15])),
            ("{dii_v17}", str(dii_value[16])),
            ("{dii_v18}", str(dii_value[17])),
            ("{dii_v19}", str(dii_value[18])),
            ("{dii_v20}", str(dii_value[19])),
            ("{dii_v21}", str(dii_value[20])),
            ("{dii_v22}", str(dii_value[21])),
            ("{dii_v23}", str(dii_value[22])),
            ("{dii_v24}", str(dii_value[23])),
            ("{dii_v25}", str(dii_value[24])),
            ("{dii_v26}", str(dii_value[25])),
            ("{dii_v27}", str(dii_value[26])),
            ("{dii_v28}", str(dii_value[27])),
            ("{dii_v29}", str(dii_value[28])),
            ("{dii_v30}", str(dii_value[29])),
            ("{dii_v31}", str(dii_value[30])),
            ("{dii_v32}", str(dii_value[31])),
            ("{dii_v33}", str(dii_value[32])),
            ("{dii_v34}", str(dii_value[33])),

            # ppi section
            ("{ppi_rank_1}", str(ppi_polda_name[0])),
            ("{ppi_rank_2}", str(ppi_polda_name[1])),
            ("{ppi_rank_3}", str(ppi_polda_name[2])),
            ("{ppi_rank_4}", str(ppi_polda_name[3])),
            ("{ppi_rank_5}", str(ppi_polda_name[4])),
            ("{ppi_rank_6}", str(ppi_polda_name[5])),
            ("{ppi_rank_7}", str(ppi_polda_name[6])),
            ("{ppi_rank_8}", str(ppi_polda_name[7])),
            ("{ppi_rank_9}", str(ppi_polda_name[8])),
            ("{ppi_rank_10}", str(ppi_polda_name[9])),
            ("{ppi_rank_11}", str(ppi_polda_name[10])),
            ("{ppi_rank_12}", str(ppi_polda_name[11])),
            ("{ppi_rank_13}", str(ppi_polda_name[12])),
            ("{ppi_rank_14}", str(ppi_polda_name[13])),
            ("{ppi_rank_15}", str(ppi_polda_name[14])),
            ("{ppi_rank_16}", str(ppi_polda_name[15])),
            ("{ppi_rank_17}", str(ppi_polda_name[16])),
            ("{ppi_rank_18}", str(ppi_polda_name[17])),
            ("{ppi_rank_19}", str(ppi_polda_name[18])),
            ("{ppi_rank_20}", str(ppi_polda_name[19])),
            ("{ppi_rank_21}", str(ppi_polda_name[20])),
            ("{ppi_rank_22}", str(ppi_polda_name[21])),
            ("{ppi_rank_23}", str(ppi_polda_name[22])),
            ("{ppi_rank_24}", str(ppi_polda_name[23])),
            ("{ppi_rank_25}", str(ppi_polda_name[24])),
            ("{ppi_rank_26}", str(ppi_polda_name[25])),
            ("{ppi_rank_27}", str(ppi_polda_name[26])),
            ("{ppi_rank_28}", str(ppi_polda_name[27])),
            ("{ppi_rank_29}", str(ppi_polda_name[28])),
            ("{ppi_rank_30}", str(ppi_polda_name[29])),
            ("{ppi_rank_31}", str(ppi_polda_name[30])),
            ("{ppi_rank_32}", str(ppi_polda_name[31])),
            ("{ppi_rank_33}", str(ppi_polda_name[32])),
            ("{ppi_rank_34}", str(ppi_polda_name[33])),

            # ppi exsposure
            ("{ppi_exp_1}", str(ppi_exsposure[0])),
            ("{ppi_exp_2}", str(ppi_exsposure[1])),
            ("{ppi_exp_3}", str(ppi_exsposure[2])),
            ("{ppi_exp_4}", str(ppi_exsposure[3])),
            ("{ppi_exp_5}", str(ppi_exsposure[4])),
            ("{ppi_exp_6}", str(ppi_exsposure[5])),
            ("{ppi_exp_7}", str(ppi_exsposure[6])),
            ("{ppi_exp_8}", str(ppi_exsposure[7])),
            ("{ppi_exp_9}", str(ppi_exsposure[8])),
            ("{ppi_exp_10}", str(ppi_exsposure[9])),
            ("{ppi_exp_11}", str(ppi_exsposure[10])),
            ("{ppi_exp_12}", str(ppi_exsposure[11])),
            ("{ppi_exp_13}", str(ppi_exsposure[12])),
            ("{ppi_exp_14}", str(ppi_exsposure[13])),
            ("{ppi_exp_15}", str(ppi_exsposure[14])),
            ("{ppi_exp_16}", str(ppi_exsposure[15])),
            ("{ppi_exp_17}", str(ppi_exsposure[16])),
            ("{ppi_exp_18}", str(ppi_exsposure[17])),
            ("{ppi_exp_19}", str(ppi_exsposure[18])),
            ("{ppi_exp_20}", str(ppi_exsposure[19])),
            ("{ppi_exp_21}", str(ppi_exsposure[20])),
            ("{ppi_exp_22}", str(ppi_exsposure[21])),
            ("{ppi_exp_23}", str(ppi_exsposure[22])),
            ("{ppi_exp_24}", str(ppi_exsposure[23])),
            ("{ppi_exp_25}", str(ppi_exsposure[24])),
            ("{ppi_exp_26}", str(ppi_exsposure[25])),
            ("{ppi_exp_27}", str(ppi_exsposure[26])),
            ("{ppi_exp_28}", str(ppi_exsposure[27])),
            ("{ppi_exp_29}", str(ppi_exsposure[28])),
            ("{ppi_exp_30}", str(ppi_exsposure[29])),
            ("{ppi_exp_31}", str(ppi_exsposure[30])),
            ("{ppi_exp_32}", str(ppi_exsposure[31])),
            ("{ppi_exp_33}", str(ppi_exsposure[32])),
            ("{ppi_exp_34}", str(ppi_exsposure[33])),

            # ppi engagement
            ("{ppi_eng_1}", str(ppi_engagement[0])),
            ("{ppi_eng_2}", str(ppi_engagement[1])),
            ("{ppi_eng_3}", str(ppi_engagement[2])),
            ("{ppi_eng_4}", str(ppi_engagement[3])),
            ("{ppi_eng_5}", str(ppi_engagement[4])),
            ("{ppi_eng_6}", str(ppi_engagement[5])),
            ("{ppi_eng_7}", str(ppi_engagement[6])),
            ("{ppi_eng_8}", str(ppi_engagement[7])),
            ("{ppi_eng_9}", str(ppi_engagement[8])),
            ("{ppi_eng_10}", str(ppi_engagement[9])),
            ("{ppi_eng_11}", str(ppi_engagement[10])),
            ("{ppi_eng_12}", str(ppi_engagement[11])),
            ("{ppi_eng_13}", str(ppi_engagement[12])),
            ("{ppi_eng_14}", str(ppi_engagement[13])),
            ("{ppi_eng_15}", str(ppi_engagement[14])),
            ("{ppi_eng_16}", str(ppi_engagement[15])),
            ("{ppi_eng_17}", str(ppi_engagement[16])),
            ("{ppi_eng_18}", str(ppi_engagement[17])),
            ("{ppi_eng_19}", str(ppi_engagement[18])),
            ("{ppi_eng_20}", str(ppi_engagement[19])),
            ("{ppi_eng_21}", str(ppi_engagement[20])),
            ("{ppi_eng_22}", str(ppi_engagement[21])),
            ("{ppi_eng_23}", str(ppi_engagement[22])),
            ("{ppi_eng_24}", str(ppi_engagement[23])),
            ("{ppi_eng_25}", str(ppi_engagement[24])),
            ("{ppi_eng_26}", str(ppi_engagement[25])),
            ("{ppi_eng_27}", str(ppi_engagement[26])),
            ("{ppi_eng_28}", str(ppi_engagement[27])),
            ("{ppi_eng_29}", str(ppi_engagement[28])),
            ("{ppi_eng_30}", str(ppi_engagement[29])),
            ("{ppi_eng_31}", str(ppi_engagement[30])),
            ("{ppi_eng_32}", str(ppi_engagement[31])),
            ("{ppi_eng_33}", str(ppi_engagement[32])),
            ("{ppi_eng_34}", str(ppi_engagement[33])),

            # dii sentiment exsposure
            ("{ppi_sexp_1}", str(ppi_sentiment_exsposure[0])),
            ("{ppi_sexp_2}", str(ppi_sentiment_exsposure[1])),
            ("{ppi_sexp_3}", str(ppi_sentiment_exsposure[2])),
            ("{ppi_sexp_4}", str(ppi_sentiment_exsposure[3])),
            ("{ppi_sexp_5}", str(ppi_sentiment_exsposure[4])),
            ("{ppi_sexp_6}", str(ppi_sentiment_exsposure[5])),
            ("{ppi_sexp_7}", str(ppi_sentiment_exsposure[6])),
            ("{ppi_sexp_8}", str(ppi_sentiment_exsposure[7])),
            ("{ppi_sexp_9}", str(ppi_sentiment_exsposure[8])),
            ("{ppi_sexp_10}", str(ppi_sentiment_exsposure[9])),
            ("{ppi_sexp_11}", str(ppi_sentiment_exsposure[10])),
            ("{ppi_sexp_12}", str(ppi_sentiment_exsposure[11])),
            ("{ppi_sexp_13}", str(ppi_sentiment_exsposure[12])),
            ("{ppi_sexp_14}", str(ppi_sentiment_exsposure[13])),
            ("{ppi_sexp_15}", str(ppi_sentiment_exsposure[14])),
            ("{ppi_sexp_16}", str(ppi_sentiment_exsposure[15])),
            ("{ppi_sexp_17}", str(ppi_sentiment_exsposure[16])),
            ("{ppi_sexp_18}", str(ppi_sentiment_exsposure[17])),
            ("{ppi_sexp_19}", str(ppi_sentiment_exsposure[18])),
            ("{ppi_sexp_20}", str(ppi_sentiment_exsposure[19])),
            ("{ppi_sexp_21}", str(ppi_sentiment_exsposure[20])),
            ("{ppi_sexp_22}", str(ppi_sentiment_exsposure[21])),
            ("{ppi_sexp_23}", str(ppi_sentiment_exsposure[22])),
            ("{ppi_sexp_24}", str(ppi_sentiment_exsposure[23])),
            ("{ppi_sexp_25}", str(ppi_sentiment_exsposure[24])),
            ("{ppi_sexp_26}", str(ppi_sentiment_exsposure[25])),
            ("{ppi_sexp_27}", str(ppi_sentiment_exsposure[26])),
            ("{ppi_sexp_28}", str(ppi_sentiment_exsposure[27])),
            ("{ppi_sexp_29}", str(ppi_sentiment_exsposure[28])),
            ("{ppi_sexp_30}", str(ppi_sentiment_exsposure[29])),
            ("{ppi_sexp_31}", str(ppi_sentiment_exsposure[30])),
            ("{ppi_sexp_32}", str(ppi_sentiment_exsposure[31])),
            ("{ppi_sexp_33}", str(ppi_sentiment_exsposure[32])),
            ("{ppi_sexp_34}", str(ppi_sentiment_exsposure[33])),

            # dii sentiment engagement
            ("{ppi_seng_1}", str(ppi_sentiment_engagement[0])),
            ("{ppi_seng_2}", str(ppi_sentiment_engagement[1])),
            ("{ppi_seng_3}", str(ppi_sentiment_engagement[2])),
            ("{ppi_seng_4}", str(ppi_sentiment_engagement[3])),
            ("{ppi_seng_5}", str(ppi_sentiment_engagement[4])),
            ("{ppi_seng_6}", str(ppi_sentiment_engagement[5])),
            ("{ppi_seng_7}", str(ppi_sentiment_engagement[6])),
            ("{ppi_seng_8}", str(ppi_sentiment_engagement[7])),
            ("{ppi_seng_9}", str(ppi_sentiment_engagement[8])),
            ("{ppi_seng_10}", str(ppi_sentiment_engagement[9])),
            ("{ppi_seng_11}", str(ppi_sentiment_engagement[10])),
            ("{ppi_seng_12}", str(ppi_sentiment_engagement[11])),
            ("{ppi_seng_13}", str(ppi_sentiment_engagement[12])),
            ("{ppi_seng_14}", str(ppi_sentiment_engagement[13])),
            ("{ppi_seng_15}", str(ppi_sentiment_engagement[14])),
            ("{ppi_seng_16}", str(ppi_sentiment_engagement[15])),
            ("{ppi_seng_17}", str(ppi_sentiment_engagement[16])),
            ("{ppi_seng_18}", str(ppi_sentiment_engagement[17])),
            ("{ppi_seng_19}", str(ppi_sentiment_engagement[18])),
            ("{ppi_seng_20}", str(ppi_sentiment_engagement[19])),
            ("{ppi_seng_21}", str(ppi_sentiment_engagement[20])),
            ("{ppi_seng_22}", str(ppi_sentiment_engagement[21])),
            ("{ppi_seng_23}", str(ppi_sentiment_engagement[22])),
            ("{ppi_seng_24}", str(ppi_sentiment_engagement[23])),
            ("{ppi_seng_25}", str(ppi_sentiment_engagement[24])),
            ("{ppi_seng_26}", str(ppi_sentiment_engagement[25])),
            ("{ppi_seng_27}", str(ppi_sentiment_engagement[26])),
            ("{ppi_seng_28}", str(ppi_sentiment_engagement[27])),
            ("{ppi_seng_29}", str(ppi_sentiment_engagement[28])),
            ("{ppi_seng_30}", str(ppi_sentiment_engagement[29])),
            ("{ppi_seng_31}", str(ppi_sentiment_engagement[30])),
            ("{ppi_seng_32}", str(ppi_sentiment_engagement[31])),
            ("{ppi_seng_33}", str(ppi_sentiment_engagement[32])),
            ("{ppi_seng_34}", str(ppi_sentiment_engagement[33])),

            # ppi value
            ("{ppi_v1}", str(ppi_value[0])),
            ("{ppi_v2}", str(ppi_value[1])),
            ("{ppi_v3}", str(ppi_value[2])),
            ("{ppi_v4}", str(ppi_value[3])),
            ("{ppi_v5}", str(ppi_value[4])),
            ("{ppi_v6}", str(ppi_value[5])),
            ("{ppi_v7}", str(ppi_value[6])),
            ("{ppi_v8}", str(ppi_value[7])),
            ("{ppi_v9}", str(ppi_value[8])),
            ("{ppi_v10}", str(ppi_value[9])),
            ("{ppi_v11}", str(ppi_value[10])),
            ("{ppi_v12}", str(ppi_value[11])),
            ("{ppi_v13}", str(ppi_value[12])),
            ("{ppi_v14}", str(ppi_value[13])),
            ("{ppi_v15}", str(ppi_value[14])),
            ("{ppi_v16}", str(ppi_value[15])),
            ("{ppi_v17}", str(ppi_value[16])),
            ("{ppi_v18}", str(ppi_value[17])),
            ("{ppi_v19}", str(ppi_value[18])),
            ("{ppi_v20}", str(ppi_value[19])),
            ("{ppi_v21}", str(ppi_value[20])),
            ("{ppi_v22}", str(ppi_value[21])),
            ("{ppi_v23}", str(ppi_value[22])),
            ("{ppi_v24}", str(ppi_value[23])),
            ("{ppi_v25}", str(ppi_value[24])),
            ("{ppi_v26}", str(ppi_value[25])),
            ("{ppi_v27}", str(ppi_value[26])),
            ("{ppi_v28}", str(ppi_value[27])),
            ("{ppi_v29}", str(ppi_value[28])),
            ("{ppi_v30}", str(ppi_value[29])),
            ("{ppi_v31}", str(ppi_value[30])),
            ("{ppi_v32}", str(ppi_value[31])),
            ("{ppi_v33}", str(ppi_value[32])),
            ("{ppi_v34}", str(ppi_value[33])),
        ])

        time = datetime.now().strftime('%Y-%m-%d')
        file_output = f"result/kpi_report_{time}.pptx"
        file_name = f"kpi_report_{time}.pptx"
        prs.save(file_output)
        replacer.write_presentation_to_file(file_output)

        if os.path.exists(save_file):
            os.remove(save_file)
        else:
            print("File not exist")

        bucket = os.getenv("S3_VORTEX_BUCKET_NETWORK")

        s3_client = boto3.client(
            endpoint_url = os.getenv("S3_VORTEX_ADDRESS"),
            service_name = 's3',
            aws_access_key_id= os.getenv("S3_VORTEX_ACCESS_KEY"),
            aws_secret_access_key=os.getenv("S3_VORTEX_SECRET_KEY"),
            config=Config(connect_timeout=90, retries={'max_attempts': 6})
        )
        s3_path = f"report/kpi-ppt/{file_name}"    
        update_query = {"$set": {"status": "success", "s3Path": f"s3://campaign-management/report/kpi-ppt/{file_name}"}}
        collection.update_one(match_id, update_query)
        s3_client.upload_file(file_output , bucket, s3_path)

        if os.path.exists(file_output):
            os.remove(file_output)
        else:
            print("File not exist")

        return f"kpi_report_{time}.pptx"
    
    except:
        raise Exception