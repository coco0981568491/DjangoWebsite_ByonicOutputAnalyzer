from django.shortcuts import render, HttpResponse
from .tasks import data_processing
from celery.result import AsyncResult
from io import BytesIO
import base64
import tempfile

# Create your views here.
def home(request):

	if request.method == "POST":

		# get data and name it as file for convenience 
		file = request.FILES["myFile"]

		# Celery does not know how to serialize complex objects such as file objects. 
		# However, this can be solved pretty easily. 
		# What I do is to encode/decode the file to its Base64 string representation. 
		# This allows me to send the file directly through Celery.

		file_bytes = file.read()
		file_bytes_base64 = base64.b64encode(file_bytes)
		file_bytes_base64_str = file_bytes_base64.decode('utf-8') # this is a str

		# (Celery worker task)
		file_bytes_base64 = file_bytes_base64_str.encode('utf-8')
		file_bytes = base64.b64decode(file_bytes_base64)

		# Write the file to a temporary location, deletion is guaranteed
		with tempfile.TemporaryDirectory() as tmp_dir:
		    tmp_file = os.path.join(tmp_dir, 'something.zip')
		    with open(tmp_file, 'wb') as file:
		        file.write(file_bytes)

		# (...send string through Celery...)
		task = data_processing.delay(file)

		res = AsyncResult(task)

		# check if the task has been finished
		if res.ready(): 

			resp = HttpResponse(res.get(), content_type = 'application/x-zip-compressed')
			resp['Content-Disposition'] = 'attachment; filename=%s'%zip_filename

			return resp

		else:
			return render(request, "progress.html")

	else:
		return render(request, "index.html")


