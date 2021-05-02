from django.shortcuts import render, HttpResponse
from .tasks import data_processing
from celery.result import AsyncResult
from io import BytesIO

# Create your views here.
def home(request):

	if request.method == "POST":

		# get data and name it as file for convenience 
		file = request.FILES["myFile"]

		# send to celery worker
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


