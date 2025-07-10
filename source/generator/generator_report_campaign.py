from datetime import datetime
import os
from elasticsearch import Elasticsearch
import pymongo

from pptx import Presentation
from python_pptx_text_replacer import TextReplacer
from .util import *
from generator.query_builder import fetch_all_query
from generator.query_builder import ReportParams
from dotenv import load_dotenv
import boto3
from botocore.client import Config
from config.db import *


load_dotenv()

def generatorRC(param: dict):
    es = ElastichConnection()
    collection = MongoConnecction()

    match_id = collection.find_one({"_id": param.get("id")})

    try:
        print('=============== Generating ReportCampaign ===============')
        prs = Presentation("template/tmp-REPORT-CAMPAIGN.pptx")
        find_chart(prs)

        data = fetch_all_query(es, param)

        tren_postingan_trendline = data.get('tren_postingan_trendline')
        tren_engagement_trendline = data.get('tren_engagement_trendline')
        expose_jumlah_postingan = data.get('expose_jumlah_postingan')
        expose_jumlah_engagement = data.get('expose_jumlah_engagement')
        expose_jumlah_engagement = sorted(expose_jumlah_engagement, key=lambda x: x['doc_count'], reverse=True)
        campaign_performance = data.get('campaign_performance')
        tren_postingan_matrick = data.get('tren_postingan_matrick')
        tren_engagement_matrick = data.get('tren_engagement_matrick')
        tren_engagement_matrick_sorted = sorted(tren_engagement_matrick, key=lambda x: x['doc_count'], reverse=True)

        # Chart section
        generate_line_chart(prs, tren_postingan_trendline, 2, 0)
        generate_line_chart(prs, tren_engagement_trendline, 3, 0)
        generate_chart_string_key(prs, expose_jumlah_postingan, 4, 0)
        generate_chart_string_key(prs, expose_jumlah_engagement, 4, 1)
        
        save_file = ("result/REPORT_CAMPAIGN_chart.pptx")
        prs.save(save_file)


        campaign_performance = sorted(campaign_performance, key=lambda x: (-x['doc_count'], -x['value']))
        campaign_performance_key = retrieve_data_object(campaign_performance, sub_key="key", length_min=10)
        campaign_performance_eksposure = retrieve_data_object(campaign_performance, sub_key="doc_count", length_min=10, default_fill=0)
        campaign_performance_engagement = retrieve_data_object(campaign_performance, sub_key="value", length_min=10, default_fill=0)
        
        tren_postingan_matrick_key = retrieve_data_object(tren_postingan_matrick, sub_key="key", length_min=5)
        if tren_postingan_matrick_key[3] == "":
            tren_postingan_matrick_key[3] = "facebook"

        tren_postingan_matrick_count = retrieve_data_object(tren_postingan_matrick, sub_key="doc_count", length_min=5, default_fill=0)
        avg_data_rank1 = rank_sum(tren_postingan_matrick_count, 0)
        avg_data_rank2 = rank_sum(tren_postingan_matrick_count, 1)

        expose_engagement_key = retrieve_data_object(expose_jumlah_engagement, sub_key='key', length_min=5)
        if expose_engagement_key[3] == "":
            tren_postingan_matrick_key[3] = "facebook"
        elif expose_engagement_key[4] == "":
            expose_engagement_key[4] = "youtube"

        print(expose_engagement_key)

        expose_engagement_count = retrieve_data_object(expose_jumlah_engagement, sub_key='doc_count', length_min=5)

        tren_engagement_matrick_key = retrieve_data_object(tren_engagement_matrick_sorted, sub_key='key', length_min=5)
        tren_engagement_matrick_count = retrieve_data_object(tren_engagement_matrick_sorted, sub_key='doc_count', length_min=5, default_fill=0) 
        sum2 = tren_engagement_matrick_count[0] + tren_engagement_matrick_count[1]

        from_time = param["timeframe"].get("from")
        to_time = param["timeframe"].get("to")

        bulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
        from_time = datetime.utcfromtimestamp(from_time / 1000)
        to_time = datetime.utcfromtimestamp(to_time / 1000)
        date = f"{from_time.day} {bulan[from_time.month - 1]} {from_time.year} - {to_time.day} {bulan[to_time.month - 1]} {to_time.year}"

        # Text Section
        replacer = TextReplacer(save_file, slides='', tables=True, charts=False, textframes=True)
        replacer.replace_text([
            ("{date}", date),

            ("{title_1}", str(campaign_performance_key[0])),
            ("{title_2}", str(campaign_performance_key[1])),
            ("{title_3}", str(campaign_performance_key[2])),
            ("{title_4}", str(campaign_performance_key[3])),
            ("{title_5}", str(campaign_performance_key[4])),
            ("{title_6}", str(campaign_performance_key[5])),

            # ("{subtitle_1}", str(campaign_performance_subkey[0])),
            # ("{subtitle_2}", str(campaign_performance_subkey[1])),
            # ("{subtitle_3}", str(campaign_performance_subkey[2])),
            # ("{subtitle_4}", str(campaign_performance_subkey[3])),
            # ("{subtitle_5}", str(campaign_performance_subkey[4])),
            # ("{subtitle_6}", str(campaign_performance_subkey[5])),

            ("{expo_1}", str(campaign_performance_eksposure[0])),
            ("{expo_2}", str(campaign_performance_eksposure[1])),
            ("{expo_3}", str(campaign_performance_eksposure[2])),
            ("{expo_4}", str(campaign_performance_eksposure[3])),
            ("{expo_5}", str(campaign_performance_eksposure[4])),
            ("{expo_6}", str(campaign_performance_eksposure[5])),

            ("{enga_1}", str(campaign_performance_engagement[0])),
            ("{enga_2}", str(campaign_performance_engagement[1])),
            ("{enga_3}", str(campaign_performance_engagement[2])),
            ("{enga_4}", str(campaign_performance_engagement[3])),
            ("{enga_5}", str(campaign_performance_engagement[4])),
            ("{enga_6}", str(campaign_performance_engagement[5])),

            # ("{nilai_1}", str(campaign_performance_skor[0])),
            # ("{nilai_2}", str(campaign_performance_skor[1])),
            # ("{nilai_3}", str(campaign_performance_skor[2])),
            # ("{nilai_4}", str(campaign_performance_skor[3])),
            # ("{nilai_5}", str(campaign_performance_skor[4])),
            # ("{nilai_6}", str(campaign_performance_skor[5])),

            ("{insta}", str(tren_postingan_matrick_count[0])),
            ("{tiktok}", str(tren_postingan_matrick_count[1])),
            ("{fb}", str(tren_postingan_matrick_count[2])),
            ("{twit}", str(tren_postingan_matrick_count[3])),
            ("{you}", str(tren_postingan_matrick_count[4])),

            ("{rank_1}", str(tren_postingan_matrick_key[0])),
            ("{rank_2}", str(tren_postingan_matrick_key[1])),
            ("{rank_3}", str(tren_postingan_matrick_key[2])),
            ("{rank_4}", str(tren_postingan_matrick_key[3])),

            ("{v_lead_1}", avg_data_rank1),
            ("{v_lead_2}", avg_data_rank2),

            ("{rnk_eng_1}", str(tren_engagement_matrick_key[0])),
            ("{rnk_eng_2}", str(tren_engagement_matrick_key[1])),
            ("{rnk_eng_3}", str(tren_engagement_matrick_key[2])),
            ("{rnk_eng_4}", str(tren_engagement_matrick_key[3])),
            ("{rnk_eng_5}", str(tren_engagement_matrick_key[4])),

            ("{sum_v12}", str(sum2)),

            ("{result_1}", str(expose_engagement_key[0])),
            ("{result_2}", str(expose_engagement_key[1])),
            ("{result_3}", str(expose_engagement_key[2])),
            ("{result_4}", str(expose_engagement_key[3])),
            ("{result_5}", str(expose_engagement_key[4])),

            ("{v_rank_1}", str(tren_postingan_matrick_count[0])),
            ("{v_rank_2}", str(tren_postingan_matrick_count[1])),
            ("{v_rank_3}", str(tren_postingan_matrick_count[2])),
            ("{v_rank_4}", str(tren_postingan_matrick_count[3])),

            ("{insta1}", str(tren_engagement_matrick_count[0])),
            ("{tiktok1}", str(tren_engagement_matrick_count[1])),
            ("{fb1}", str(tren_engagement_matrick_count[2])),
            ("{twit1}", str(tren_engagement_matrick_count[3])),
            ("{you1}", str(tren_engagement_matrick_count[4])),

            ("{influencer_expo_1}", str(campaign_performance_eksposure[0])),
            ("{influencer_expo_2}", str(campaign_performance_eksposure[1])),
            ("{influencer_expo_3}", str(campaign_performance_eksposure[2])),
            ("{influencer_expo_4}", str(campaign_performance_eksposure[3])),
            ("{influencer_expo_5}", str(campaign_performance_eksposure[4])),
            ("{influencer_expo_6}", str(campaign_performance_eksposure[5])),

            ("{i_enga_1}", str(campaign_performance_engagement[0])),
            ("{i_enga_2}", str(campaign_performance_engagement[1])),
            ("{i_enga_3}", str(campaign_performance_engagement[2])),
            ("{i_enga_4}", str(campaign_performance_engagement[3])),
            ("{i_enga_5}", str(campaign_performance_engagement[4])),
            ("{i_enga_6}", str(campaign_performance_engagement[5])),

            ("{judul_1}", str(campaign_performance_key[0])),
            ("{judul_2}", str(campaign_performance_key[1])),
            ("{judul_3}", str(campaign_performance_key[2])),
            ("{judul_4}", str(campaign_performance_key[3])),
            ("{judul_5}", str(campaign_performance_key[4])),
            ("{judul_6}", str(campaign_performance_key[5])),

            # ("{penjelasan_1}", str(campaign_performance_subkey[0])),
            # ("{penjelasan_2}", str(campaign_performance_subkey[1])),
            # ("{penjelasan_3}", str(campaign_performance_subkey[2])),
            # ("{penjelasan_4}", str(campaign_performance_subkey[3])),
            # ("{penjelasan_5}", str(campaign_performance_subkey[4])),
            # ("{penjelasan_6}", str(campaign_performance_subkey[5])),
        ])

        time = datetime.now().strftime('%Y-%m-%d')
        file_output = f"result/campaign_report_{time}.pptx"
        file_name = f"campaign_report_{time}.pptx"
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

        s3_path = f"report/campaign-ppt/{file_name}"    
        update_query = {"$set": {"status": "success", "s3Path": f"s3://campaign-management/report/campaign-ppt/{file_name}"}}
        collection.update_one(match_id, update_query)
        s3_client.upload_file(file_output , bucket, s3_path)

        if os.path.exists(file_output):
            os.remove(file_output)
        else:
            print("File not exist")
    
        return f"campaign_report_{time}.pptx"
    except Exception as e:
        print("error when generating report: ", e)