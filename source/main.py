import os
import uvicorn
import pymongo
import pika
import json
from fastapi import FastAPI, BackgroundTasks, HTTPException
from fastapi.responses import FileResponse
from generator.generator_report_campaign import generatorRC
from generator.generator_report_kpi import generate_report_kpi
from generator.generator_report_excel import *
from generator.util import *


def generate_report_task_campaign(data: dict):
    return generatorRC(data)

def generate_report_task_kpi(data: dict):
    return generate_report_kpi(data)

def generate_report_task_campaign_excel(data: dict):
    return generate_report_campaign_excel(data)

def generate_report_task_kpi_excel(data: dict):
    return generate_report_kpi_excel(data)


def process_message(message):
    report_type = message.get("reportType")
    try:
        if report_type == "campaign-ppt":
            result = generate_report_task_campaign(message)
            print(f"Report Type campaign-ppt (Campaign PPT) generated: {result}")
            
        elif report_type == "kpi-ppt":
            result = generate_report_task_kpi(message)
            print(f"Report Type kpi-ppt (KPI PPT) generated: {result}")

        elif report_type == "campaign":
            result = generate_report_task_campaign_excel(message)
            print(f"Report Type campaign (Campaign Excel) generated: {result}")

        elif report_type == "kpi":
            result = generate_report_task_kpi_excel(message)
            print(f"Report Type kpi (KPI Excel) generated: {result}")

        else:
            print(f"Unknown report type: {report_type}")
    except Exception as e:
        print(f"Error while generating report for {report_type}: {e}")


def consume_messages():
    RABBITMQ_HOST = os.getenv("RABBITMQ_HOST")
    RABBITMQ_PORT = os.getenv("RABBITMQ_PORT")
    RABBITMQ_USERNAME = os.getenv("RABBITMQ_USERNAME")
    RABBITMQ_PASSWORD = os.getenv("RABBITMQ_PASSWORD")
    RABBITMQ_VHOST = os.getenv("RMQ_VHOST")
    RABBITMQ_EXCHANGE = os.getenv("RABBITMQ_EXCHANGE")
    RABBITMQ_QUEUE = os.getenv("RABBITMQ_QUEUE")
    RABBITMQ_ROUTING_KEY = os.getenv("RABBITMQ_ROUTING_KEY")

    credentials = pika.PlainCredentials(RABBITMQ_USERNAME, RABBITMQ_PASSWORD)
    parameters = pika.ConnectionParameters(RABBITMQ_HOST, RABBITMQ_PORT, RABBITMQ_VHOST, credentials)
    connection = pika.BlockingConnection(parameters)
    channel = connection.channel()

    channel.exchange_declare(exchange=RABBITMQ_EXCHANGE, exchange_type='direct', durable=True)
    channel.queue_declare(queue=RABBITMQ_QUEUE, durable=True)
    channel.queue_bind(exchange=RABBITMQ_EXCHANGE, queue=RABBITMQ_QUEUE, routing_key=RABBITMQ_ROUTING_KEY)

    def callback(ch, method, properties, body):
        print(f"Received message: {body}")

        try:
            message_data= json.loads(body)
            print(message_data)
            process_message(message_data)
            ch.basic_ack(delivery_tag=method.delivery_tag)
        except Exception as e:
            print(f"Error processing message: {e}")
            ch.basic_nack(delivery_tag=method.delivery_tag, requeue=False)

    channel.basic_consume(queue=RABBITMQ_QUEUE, on_message_callback=callback)
    print('Waiting for messages...')
    channel.start_consuming()
   
if __name__ == "__main__":
    consume_messages()
