from .util import *
import json

def get_query(widget):
    # camppaign
    if widget == "tren_postingan_matrick":
        return {"aggs":{"2":{"terms":{"field":"platform","order":{"_count":"desc"},"size":5,"shard_size":25}}},"size":0,"fields":[{"field":"account_category.joind_date","format":"date_time"},{"field":"created_at","format":"date_time"},{"field":"group_created_at","format":"date_time"},{"field":"quoted_created_at","format":"date_time"},{"field":"quoted_user_created_at","format":"date_time"},{"field":"retweeted_status_created_at","format":"date_time"},{"field":"retweeted_user_created_at","format":"date_time"},{"field":"user_created_at","format":"date_time"}],"script_fields":{},"stored_fields":["*"],"runtime_mappings":{},"_source":{"excludes":[]},"query":{"bool":{"must":[],"filter":[{"exists":{"field":"campaign.campaign_id"}},{"bool":{"minimum_should_match":1,"should":[{"match_phrase":{"satuan_wilayah_polri":"polda"}}]}},{"range":{"created_at":{"format":"strict_date_optional_time","gte":"2024-10-03T06:41:55.324Z","lte":"2025-01-03T06:41:55.324Z"}}}],"should":[],"must_not":[]}}}
    elif widget == "tren_postingan_trendline":
        return {"aggs":{"2":{"date_histogram":{"field":"created_at","calendar_interval":"1d","time_zone":"Asia/Bangkok","min_doc_count":1},"aggs":{"3":{"terms":{"field":"platform","order":{"_count":"desc"},"size":5,"shard_size":25}}}}},"size":0,"fields":[{"field":"account_category.joind_date","format":"date_time"},{"field":"created_at","format":"date_time"},{"field":"group_created_at","format":"date_time"},{"field":"quoted_created_at","format":"date_time"},{"field":"quoted_user_created_at","format":"date_time"},{"field":"retweeted_status_created_at","format":"date_time"},{"field":"retweeted_user_created_at","format":"date_time"},{"field":"user_created_at","format":"date_time"}],"script_fields":{},"stored_fields":["*"],"runtime_mappings":{},"_source":{"excludes":[]},"query":{"bool":{"must":[],"filter":[{"exists":{"field":"campaign.campaign_id"}},{"exists":{"field":"jurisdiction_area.polda_code"}},{"range":{"created_at":{"format":"strict_date_optional_time","gte":"2024-09-20T06:27:06.852Z","lte":"2025-01-03T06:27:06.852Z"}}}],"should":[],"must_not":[]}}}
    elif widget == "campaign_performance":
        return {"aggs":{"2":{"terms":{"field":"jurisdiction_area.polda.keyword","order":{"1":"desc"},"size":10,"shard_size":25},"aggs":{"1":{"sum":{"field":"engagement"}}}}},"size":0,"fields":[{"field":"account_category.joind_date","format":"date_time"},{"field":"created_at","format":"date_time"},{"field":"group_created_at","format":"date_time"},{"field":"quoted_created_at","format":"date_time"},{"field":"quoted_user_created_at","format":"date_time"},{"field":"retweeted_status_created_at","format":"date_time"},{"field":"retweeted_user_created_at","format":"date_time"},{"field":"user_created_at","format":"date_time"}],"script_fields":{},"stored_fields":["*"],"runtime_mappings":{},"_source":{"excludes":[]},"query":{"bool":{"must":[],"filter":[{"exists":{"field":"campaign.campaign_id"}},{"range":{"created_at":{"format":"strict_date_optional_time","gte":"2024-09-20T06:20:25.729Z","lte":"2025-01-03T06:20:25.729Z"}}}],"should":[],"must_not":[]}}}
    elif widget == "tren_engagement_matrick":
        return {"aggs":{"2":{"terms":{"field":"platform","order":{"1":"desc"},"size":5,"shard_size":25},"aggs":{"1":{"sum":{"field":"engagement"}}}}},"size":0,"fields":[{"field":"account_category.joind_date","format":"date_time"},{"field":"created_at","format":"date_time"},{"field":"group_created_at","format":"date_time"},{"field":"quoted_created_at","format":"date_time"},{"field":"quoted_user_created_at","format":"date_time"},{"field":"retweeted_status_created_at","format":"date_time"},{"field":"retweeted_user_created_at","format":"date_time"},{"field":"user_created_at","format":"date_time"}],"script_fields":{},"stored_fields":["*"],"runtime_mappings":{},"_source":{"excludes":[]},"query":{"bool":{"must":[],"filter":[{"exists":{"field":"campaign.campaign_id"}},{"bool":{"minimum_should_match":1,"should":[{"match_phrase":{"satuan_wilayah_polri":"polda"}}]}},{"range":{"created_at":{"format":"strict_date_optional_time","gte":"2024-10-03T06:41:55.324Z","lte":"2025-01-03T06:41:55.324Z"}}}],"should":[],"must_not":[]}}}
    elif widget == "tren_engagement_trendline":
        return {"aggs":{"2":{"date_histogram":{"field":"created_at","calendar_interval":"1d","time_zone":"Asia/Bangkok","min_doc_count":1},"aggs":{"3":{"terms":{"field":"platform","order":{"1":"desc"},"size":5,"shard_size":25},"aggs":{"1":{"sum":{"field":"engagement"}}}}}}},"size":0,"fields":[{"field":"account_category.joind_date","format":"date_time"},{"field":"created_at","format":"date_time"},{"field":"group_created_at","format":"date_time"},{"field":"quoted_created_at","format":"date_time"},{"field":"quoted_user_created_at","format":"date_time"},{"field":"retweeted_status_created_at","format":"date_time"},{"field":"retweeted_user_created_at","format":"date_time"},{"field":"user_created_at","format":"date_time"}],"script_fields":{},"stored_fields":["*"],"runtime_mappings":{},"_source":{"excludes":[]},"query":{"bool":{"must":[],"filter":[{"exists":{"field":"campaign.campaign_id"}},{"exists":{"field":"jurisdiction_area.polda_code"}},{"range":{"created_at":{"format":"strict_date_optional_time","gte":"2024-09-20T06:27:06.852Z","lte":"2025-01-03T06:27:06.852Z"}}}],"should":[],"must_not":[]}}}
    elif widget == "expose_jumlah_postingan":
        return {"aggs":{"2":{"terms":{"field":"platform","order":{"_count":"desc"},"size":5,"shard_size":25}}},"size":0,"fields":[{"field":"account_category.joind_date","format":"date_time"},{"field":"created_at","format":"date_time"},{"field":"group_created_at","format":"date_time"},{"field":"quoted_created_at","format":"date_time"},{"field":"quoted_user_created_at","format":"date_time"},{"field":"retweeted_status_created_at","format":"date_time"},{"field":"retweeted_user_created_at","format":"date_time"},{"field":"user_created_at","format":"date_time"}],"script_fields":{},"stored_fields":["*"],"runtime_mappings":{},"_source":{"excludes":[]},"query":{"bool":{"must":[],"filter":[{"exists":{"field":"campaign.campaign_id"}},{"bool":{"minimum_should_match":1,"should":[{"match_phrase":{"satuan_wilayah_polri":"polda"}}]}},{"range":{"created_at":{"format":"strict_date_optional_time","gte":"2024-09-20T06:52:27.595Z","lte":"2025-01-03T06:52:27.595Z"}}}],"should":[],"must_not":[]}}}
    elif widget == "expose_jumlah_engagement":
        return {"aggs":{"2":{"terms":{"field":"platform","order":{"1":"desc"},"size":5,"shard_size":25},"aggs":{"1":{"sum":{"field":"engagement"}}}}},"size":0,"fields":[{"field":"account_category.joind_date","format":"date_time"},{"field":"created_at","format":"date_time"},{"field":"group_created_at","format":"date_time"},{"field":"quoted_created_at","format":"date_time"},{"field":"quoted_user_created_at","format":"date_time"},{"field":"retweeted_status_created_at","format":"date_time"},{"field":"retweeted_user_created_at","format":"date_time"},{"field":"user_created_at","format":"date_time"}],"script_fields":{},"stored_fields":["*"],"runtime_mappings":{},"_source":{"excludes":[]},"query":{"bool":{"must":[],"filter":[{"exists":{"field":"campaign.campaign_id"}},{"bool":{"minimum_should_match":1,"should":[{"match_phrase":{"satuan_wilayah_polri":"polda"}}]}},{"range":{"created_at":{"format":"strict_date_optional_time","gte":"2024-09-20T06:52:27.595Z","lte":"2025-01-03T06:52:27.595Z"}}}],"should":[],"must_not":[]}}}
    
    # kpi
    elif widget == "tren_eksposure_matrick":
        return {"query":{"bool":{"must":[{"exists":{"field":"satuan_wilayah_polri"}},{"range":{"created_at":{"gte":"2024-12-30","lte":"2025-01-05","format":"yyyy-MM-dd"}}}]}},"size":0,"aggs":{"1":{"terms":{"field":"platform","size":10,"order":{"_count":"desc"}}}}}
    elif widget == "tren_eksposure_total":
        return {"query":{"bool":{"must":[{"exists":{"field":"satuan_wilayah_polri"}},{"range":{"created_at":{"gte":"2024-12-30","lte":"2025-01-05","format":"yyyy-MM-dd"}}}]}},"size":0,"aggs":{"Engagement":{"sum":{"field":"engagement"}},"Eksposure":{"value_count":{"field":"id"}}}}
    elif widget == "tren_eksposure_trendline":
        return {"query":{"bool":{"must":[{"exists":{"field":"satuan_wilayah_polri"}},{"range":{"created_at":{"gte":"2024-12-30","lte":"2025-01-05","format":"yyyy-MM-dd"}}}]}},"size":0,"aggs":{"1":{"date_histogram":{"field":"created_at","fixed_interval":"1d","time_zone":"Asia/Jakarta","min_doc_count":1}}}}

    # campaign excel
    elif widget == "polda":
        return {"aggs":{"satwil":{"terms":{"field":"jurisdiction_area.polda.keyword","order":{"_key":"asc"},"size":100},"aggs":{"campaign":{"terms":{"field":"campaign.campaign_name","order":{"_key":"asc"},"size":5,"shard_size":25},"aggs":{"platform":{"terms":{"field":"campaign.campaign_platform","order":{"_key":"asc"},"size":5,"shard_size":25},"aggs":{"task":{"terms":{"field":"campaign.campaign_task","order":{"_key":"asc"},"size":5,"shard_size":25},"aggs":{"created_at":{"date_histogram":{"field":"created_at","calendar_interval":"1w","time_zone":"Asia/Bangkok","min_doc_count":1}}}}}}}}}}},"size":0,"fields":[{"field":"account_category.joind_date","format":"date_time"},{"field":"created_at","format":"date_time"},{"field":"group_created_at","format":"date_time"},{"field":"quoted_created_at","format":"date_time"},{"field":"quoted_user_created_at","format":"date_time"},{"field":"retweeted_status_created_at","format":"date_time"},{"field":"retweeted_user_created_at","format":"date_time"},{"field":"user_created_at","format":"date_time"}],"script_fields":{},"stored_fields":["*"],"runtime_mappings":{},"_source":{"excludes":[]},"query":{"bool":{"must":[],"filter":[{"exists":{"field":"campaign.campaign_id"}},{"exists":{"field":"jurisdiction_area.polda_code"}},{"bool":{"minimum_should_match":1,"should":[{"match_phrase":{"satuan_wilayah_polri":"polda"}}]}},{"range":{"created_at":{"format":"strict_date_optional_time","gte":"2023-10-02T03:55:10.562Z","lte":"2025-01-02T03:55:10.562Z"}}}],"should":[],"must_not":[]}}}
    elif widget == "polres":
        return {"aggs":{"satwil":{"terms":{"field":"jurisdiction_area.polres.keyword","order":{"_key":"asc"},"size":100},"aggs":{"campaign":{"terms":{"field":"campaign.campaign_name","order":{"_key":"asc"},"size":5,"shard_size":25},"aggs":{"platform":{"terms":{"field":"campaign.campaign_platform","order":{"_key":"asc"},"size":5,"shard_size":25},"aggs":{"task":{"terms":{"field":"campaign.campaign_task","order":{"_key":"asc"},"size":5,"shard_size":25},"aggs":{"created_at":{"date_histogram":{"field":"created_at","calendar_interval":"1w","time_zone":"Asia/Bangkok","min_doc_count":1}}}}}}}}}}},"size":0,"fields":[{"field":"account_category.joind_date","format":"date_time"},{"field":"created_at","format":"date_time"},{"field":"group_created_at","format":"date_time"},{"field":"quoted_created_at","format":"date_time"},{"field":"quoted_user_created_at","format":"date_time"},{"field":"retweeted_status_created_at","format":"date_time"},{"field":"retweeted_user_created_at","format":"date_time"},{"field":"user_created_at","format":"date_time"}],"script_fields":{},"stored_fields":["*"],"runtime_mappings":{},"_source":{"excludes":[]},"query":{"bool":{"must":[],"filter":[{"exists":{"field":"campaign.campaign_id"}},{"exists":{"field":"jurisdiction_area.polres_code"}},{"bool":{"minimum_should_match":1,"should":[{"match_phrase":{"satuan_wilayah_polri":"polres"}}]}},{"range":{"created_at":{"format":"strict_date_optional_time","gte":"2023-10-02T03:55:10.562Z","lte":"2025-01-02T03:55:10.562Z"}}}],"should":[],"must_not":[]}}}
    elif widget == "polsek":
        return {"aggs":{"satwil":{"terms":{"field":"jurisdiction_area.polsek.keyword","order":{"_key":"asc"},"size":100},"aggs":{"campaign":{"terms":{"field":"campaign.campaign_name","order":{"_key":"asc"},"size":5,"shard_size":25},"aggs":{"platform":{"terms":{"field":"campaign.campaign_platform","order":{"_key":"asc"},"size":5,"shard_size":25},"aggs":{"task":{"terms":{"field":"campaign.campaign_task","order":{"_key":"asc"},"size":5,"shard_size":25},"aggs":{"created_at":{"date_histogram":{"field":"created_at","calendar_interval":"1w","time_zone":"Asia/Bangkok","min_doc_count":1}}}}}}}}}}},"size":0,"fields":[{"field":"account_category.joind_date","format":"date_time"},{"field":"created_at","format":"date_time"},{"field":"group_created_at","format":"date_time"},{"field":"quoted_created_at","format":"date_time"},{"field":"quoted_user_created_at","format":"date_time"},{"field":"retweeted_status_created_at","format":"date_time"},{"field":"retweeted_user_created_at","format":"date_time"},{"field":"user_created_at","format":"date_time"}],"script_fields":{},"stored_fields":["*"],"runtime_mappings":{},"_source":{"excludes":[]},"query":{"bool":{"must":[],"filter":[{"exists":{"field":"campaign.campaign_id"}},{"exists":{"field":"jurisdiction_area.polsek_code"}},{"bool":{"minimum_should_match":1,"should":[{"match_phrase":{"satuan_wilayah_polri":"polsek"}}]}},{"range":{"created_at":{"format":"strict_date_optional_time","gte":"2023-10-02T03:55:10.562Z","lte":"2025-01-02T03:55:10.562Z"}}}],"should":[],"must_not":[]}}}

    return None

def fetch_all_query(es, params: dict):
    print("ini params: ", params)

    widgets = ["campaign_performance", "tren_postingan_matrick", "tren_postingan_trendline", "tren_engagement_matrick", "tren_engagement_trendline", "expose_jumlah_postingan", "expose_jumlah_engagement", "polda", "polres", "polsek"]
    widget_handlers = {
        # campaign
        "campaign_performance": handle_aggregation,
        "tren_postingan_matrick": handle_unested,
        "tren_postingan_trendline": handle_nested,
        "tren_engagement_matrick": handle_unested,
        "tren_engagement_trendline": handle_nested,
        "expose_jumlah_postingan": handle_unested,
        "expose_jumlah_engagement": handle_unested,
        "polda": process_elastic_data,
        "polres": process_elastic_data,
        "polsek": process_elastic_data,
    }
    widget_data = {}
    
    for widget_name in widgets:
        query = get_query(widget_name)
        handler = widget_handlers.get(widget_name, lambda x: [])


        if params["timeframe"] != None:
            query = update_query_with_timeframe(query, params)

        if hasattr(params, "filters") and params["filters"]:
            for filter_item in params.get("filters"):
                print("filter item: ", filter_item)
                if filter_item:  
                    add_filter = {
                        "match_phrase": {
                            filter_item.get("field"): filter_item.get("value")
                        }
                    }
                    query['query']['bool'].setdefault('filter', []).append(add_filter)
        # print(json.dumps(query, indent=2))
        
        if query:
            index_name = "vortex-v2*"
            result = es.search(index=index_name, body=query)
            widget_data[widget_name] = handler(result)
        else:
            print(f"Query untuk widget {widget_name} tidak ditemukan!")
    
    return widget_data

def fetch_kpi_query(es, params: dict):
    # params = parse_report_params(data)
    print("param query: ", params)


    widgets = ['tren_eksposure_matrick', 'tren_eksposure_total', 'tren_eksposure_trendline']
    widget_handlers = {
        "tren_eksposure_matrick": handle_unested_v2,
        "tren_eksposure_total": handle_simple_data,
        "tren_eksposure_trendline": handle_aggregation_v2
    }
    widget_data = {}

    for widget_name in widgets:
        query = get_query(widget_name)
        handler = widget_handlers.get(widget_name, lambda x: [])
    
        if params["timeframe"] != None:
            query = update_query_with_timeframe_must(query, params)

        if hasattr(params, "filters") and params["filters"]:
            for filter_item in params.get("filters"):
                print("filter item: ", filter_item)
                if filter_item:  
                    add_filter = {
                        "match_phrase": {
                            filter_item.get("field"): filter_item.get("value")
                        }
                    }
                    query['query']['bool'].setdefault('must', []).append(add_filter)
        # print(json.dumps(query, indent=2))
        
        if query:
            index_name = "vortex-v2*"
            result = es.search(index=index_name, body=query)
            widget_data[widget_name] = handler(result)
        else:
            print(f"Query untuk widget {widget_name} tidak ditemukan!")
    
    return widget_data