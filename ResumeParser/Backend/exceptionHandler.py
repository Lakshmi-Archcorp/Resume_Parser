import json
from werkzeug.exceptions import HTTPException
from flask import Response


class ExceptionHandler:

    def handleError(e):
        print(e)
        headers = {"Access-Control-Allow-Origin": "*"}
        status = 500
        errorMessage = "InternalServerError"

        if isinstance(e, HTTPException):
            status = 400
            errorMessage = "Bad Request"
        if isinstance(e, ZeroDivisionError):
            status = 400
            errorMessage = "Mathematical Error Divide by Zero"
        
        details = str(e)
        print(details)

        jsonResponse = json.dumps({"message": errorMessage, "detail":details})
        return Response(status = status, response=jsonResponse, mimetype="application/json", headers=headers)
