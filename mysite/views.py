from django.shortcuts import render, HttpResponse
from .tasks import data_processing
from .tasks import test
from celery.result import AsyncResult
from io import BytesIO
import base64
import os

# Create your views here.
def home(request):

	if request.method == "POST":

		# get data and name it as file for convenience 
		file = request.FILES["myFile"]
		filename = file.name

		# Celery does not know how to serialize complex objects such as file objects. 
		# However, this can be solved pretty easily. 
		# What I do is to encode/decode the file to its Base64 string representation. 
		# This allows me to send the file directly through Celery.

		file_bytes = file.read()
		file_bytes_base64 = base64.b64encode(file_bytes)
		file = file_bytes_base64.decode('utf-8') # this is a str

		

		# (...send string through Celery...)
		task = data_processing.delay(file, filename)

		request.session['id'] = task.task_id
		request.session.modified = True

		# task_id = task.task_id

		# wait until task is ready, and return its result
		# status = task.status
		
		# data_processing.delay(file_bytes_base64_str, filename)

		# return render(request, "progress.html", context={'task_id': task.task_id})

		# print('this is the task status: %s'%task.status)
		
		return render(request, "progress.html", context={'task_id': task.task_id})
			

	else:
		return render(request, "index.html")


def download(request):

	# task_id = request.session['id']

	# results = AsyncResult(task_id).get()
	# # results = task.get()

	# zip_filename = 'Results.zip'

	# # convert results in base64 str back into bytes
	# results_bytes_base64 = results.encode('utf-8')
	# results_bytes = base64.b64decode(results_bytes_base64)
	

	# resp = HttpResponse(results_bytes, content_type = 'application/x-zip-compressed')
	# resp['Content-Disposition'] = 'attachment; filename=%s'%zip_filename

	# return resp
	return render(request, "test.html")
	# resp = HttpResponse("status is SUCCESS!")
	# return resp
