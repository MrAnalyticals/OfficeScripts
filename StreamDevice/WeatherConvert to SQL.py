import logging
import azure.functions as func
import json


def main(req: func.HttpRequest) -> func.HttpResponse:
    try:
        req_body = req.get_json()
        input_obj = req_body['value'][0]

        time_capture = input_obj['TimeCapture']
        temperature = input_obj['Temperature']
        oxygen_level = input_obj['Oxygen Level']
        nitrogen_level = input_obj['Nitrogen Level']
        carbon_dioxide = input_obj['Carbon Dioxide']

        sql = f"INSERT INTO dbo.[StreamingTemperature] (TimeCapture, Temperature, [Oxygen Level], [Nitrogen Level], [Carbon Dioxide]) VALUES ('{time_capture}', {temperature}, {oxygen_level}, {nitrogen_level}, {carbon_dioxide})"

        return func.HttpResponse(body=sql, status_code=200)

    except Exception as e:
        logging.error(str(e))
        return func.HttpResponse("An error occurred", status_code=400)

