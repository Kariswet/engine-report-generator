from .util import *

def get_query_sql(widget):
    if widget == "key_performance_indicator":
        return """
            SELECT 
                polda AS "Polda",
                SUM(CASE WHEN kpi_type = 'dpi' AND type_statistic = '' THEN value ELSE 0 END) AS "DPI",
                SUM(CASE WHEN kpi_type = 'dii' AND type_statistic = '' THEN value ELSE 0 END) AS "DII",
                SUM(CASE WHEN kpi_type = 'ppi' AND type_statistic = '' THEN value ELSE 0 END) AS "PPI",
                SUM(CASE WHEN kpi_type = 'overall_score' THEN value ELSE 0 END) AS "KPI"
            FROM kpi_polda_result
            WHERE 
                start_time = '2024-12-30 00:00:00.000'  
                AND end_time = '2025-01-05 23:59:59.000'
                AND satuan_wilayah = 'polda'
            GROUP BY polda
            ORDER BY "KPI" DESC;
        """
    elif widget == "digital_platform_indicator":
        return """
            SELECT 
                polda AS "Polda",
                SUM(CASE WHEN kpi_type = 'dpi' AND type_statistic = 'eksposure' THEN value ELSE 0 END) AS "Exposure",
                SUM(CASE WHEN kpi_type = 'dpi' AND type_statistic = 'engagement' THEN value ELSE 0 END) AS "Engagement",
                SUM(CASE WHEN kpi_type = 'dpi' AND type_statistic = '' THEN value ELSE 0 END) AS "DPI"
            FROM kpi_polda_result
            WHERE 
                start_time = '2024-12-30 00:00:00.000'
                AND end_time = '2025-01-05 23:59:59.000'
                AND satuan_wilayah = 'polda'
            GROUP BY polda
            ORDER BY "DPI" DESC;
        """
    elif widget == "digital_interaction_indicator":
        return """
            SELECT 
                polda AS "Polda",
                SUM(CASE WHEN kpi_type = 'dii' AND type_statistic = 'eksposure' THEN value ELSE 0 END) AS "Exposure",
                SUM(CASE WHEN kpi_type = 'dii' AND type_statistic = 'engagement' THEN value ELSE 0 END) AS "Engagement",
                SUM(CASE WHEN kpi_type = 'dii' AND type_statistic = 'sentiment eksposure' THEN value ELSE 0 END) AS "Sentiment Exposure",
                SUM(CASE WHEN kpi_type = 'dii' AND type_statistic = 'sentiment engagement' THEN value ELSE 0 END) AS "Sentiment Engagement",
                SUM(CASE WHEN kpi_type = 'dii' AND type_statistic = '' THEN value ELSE 0 END) AS "DII"
            FROM kpi_polda_result
            WHERE 
                start_time = '2024-12-30 00:00:00.000'
                AND end_time = '2025-01-05 23:59:59.000'
                AND satuan_wilayah = 'polda'
            GROUP BY polda
            ORDER BY "DII" DESC;
        """
    elif widget == "public_perception_indicator":
        return """
            SELECT 
                polda AS "Polda",
                SUM(CASE WHEN kpi_type = 'ppi' AND type_statistic = 'eksposure' THEN value ELSE 0 END) AS "Exposure",
                SUM(CASE WHEN kpi_type = 'ppi' AND type_statistic = 'engagement' THEN value ELSE 0 END) AS "Engagement",
                SUM(CASE WHEN kpi_type = 'ppi' AND type_statistic = 'sentiment eksposure' THEN value ELSE 0 END) AS "Sentiment Exposure",
                SUM(CASE WHEN kpi_type = 'ppi' AND type_statistic = 'sentiment engagement' THEN value ELSE 0 END) AS "Sentiment Engagement",
                SUM(CASE WHEN kpi_type = 'ppi' AND type_statistic = '' THEN value ELSE 0 END) AS "PPI"
            FROM kpi_polda_result
            WHERE 
                start_time = '2024-12-30 00:00:00.000'
                AND end_time = '2025-01-05 23:59:59.000'
                AND satuan_wilayah = 'polda'
            GROUP BY polda
            ORDER BY "PPI" DESC;
        """
    elif widget == "polda":
        return """
        SELECT 
            polda,
            SUM(CASE WHEN kpi_type = 'dpi' AND type_statistic = '' THEN value ELSE 0 END) AS dpi_score,
            SUM(CASE WHEN kpi_type = 'dii' AND type_statistic = '' THEN value ELSE 0 END) AS dii_score,
            SUM(CASE WHEN kpi_type = 'ppi' AND type_statistic = '' THEN value ELSE 0 END) AS ppi_score,
            SUM(CASE WHEN kpi_type = 'overall_score' THEN value ELSE 0 END) AS overall_score
        FROM kpi_polda_result
        where 1=1 and start_time >= '2024-12-30 00:00:00.000' and end_time <= '2025-01-05 23:59:59.000'
        AND satuan_wilayah = 'polda'
        GROUP BY polda, polres, polsek
        ORDER BY overall_score DESC
        """
    elif widget == "polres":
        return """
        SELECT 
            polda,
            polres,
            SUM(CASE WHEN kpi_type = 'dpi' AND type_statistic = '' THEN value ELSE 0 END) AS dpi_score,
            SUM(CASE WHEN kpi_type = 'dii' AND type_statistic = '' THEN value ELSE 0 END) AS dii_score,
            SUM(CASE WHEN kpi_type = 'ppi' AND type_statistic = '' THEN value ELSE 0 END) AS ppi_score,
            SUM(CASE WHEN kpi_type = 'overall_score' THEN value ELSE 0 END) AS overall_score
        FROM kpi_polda_result
        where 1=1 and start_time >= '2024-12-30 00:00:00.000' and end_time <= '2025-01-05 23:59:59.000'
        AND satuan_wilayah = 'polres'
        GROUP BY polda, polres, polsek
        ORDER BY overall_score DESC
        """
    elif widget == "polsek":
        return """
        SELECT 
            polda,
            polres,
            polsek,
            SUM(CASE WHEN kpi_type = 'dpi' AND type_statistic = '' THEN value ELSE 0 END) AS dpi_score,
            SUM(CASE WHEN kpi_type = 'dii' AND type_statistic = '' THEN value ELSE 0 END) AS dii_score,
            SUM(CASE WHEN kpi_type = 'ppi' AND type_statistic = '' THEN value ELSE 0 END) AS ppi_score,
            SUM(CASE WHEN kpi_type = 'overall_score' THEN value ELSE 0 END) AS overall_score
        FROM kpi_polda_result
        where 1=1 and start_time >= '2024-12-30 00:00:00.000' and end_time <= '2025-01-05 23:59:59.000'
        AND satuan_wilayah = 'polsek'
        GROUP BY polda, polres, polsek
        ORDER BY overall_score DESC
        """

    return None

def fetch_sql_query(conn, data: dict):
    params = parse_report_params(data)
    
    widgets = [
        "key_performance_indicator",
        "digital_platform_indicator",
        "digital_interaction_indicator",
        "public_perception_indicator",
        "polda",
        "polres",
        "polsek"
    ]
    widget_handler = {
        "key_performance_indicator": handle_query_result_kpi,
        "digital_platform_indicator": handle_query_result_dpi,
        "digital_interaction_indicator": handle_query_result_dii,
        "public_perception_indicator": handle_query_result_ppi,
        "polda": process_postgre_data,
        "polres": process_postgre_data,
        "polsek": process_postgre_data
    }
    widget_data = {}


    for widget_name in widgets:
        query = get_query_sql(widget_name)
        handler = widget_handler.get(widget_name, lambda x: [])

        from_time = convert_epoch_to_iso(params.timeframe.from_)
        to_time = convert_epoch_to_iso(params.timeframe.to)

        if params.timeframe != None:
            query = query.replace("BETWEEN '2024-12-30 00:00:00.000' AND '2025-01-05 23:59:59.000'", f"BETWEEN '{from_time}' AND '{to_time}'")

        if query:
            with conn.cursor() as cursor:
                cursor.execute(query)
                result = cursor.fetchall()
                
                widget_data[widget_name] = handler(result)
        else:
            print(f"Query for widget {widget_name} not found")

    return widget_data