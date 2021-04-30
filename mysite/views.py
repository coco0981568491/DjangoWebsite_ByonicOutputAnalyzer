from django.shortcuts import render, HttpResponse
from django_q.tasks import AsyncTask
from mysite.services import get_results



# Create your views here.
def home(request):
	if request.method == "POST":

		# get data
		file = request.FILES["myFile"]

		# instantiate an async task, will return uuid of the task
		a = AsyncTask("mysite.services.data_processing", file, hook = "mysite.services.get_results")

		# # run it
		a.run()

		# result = a.result(wait=-1)

		# print(result)

		# # Grab ZIP file from in-memory, make response with correct content-type
		# resp = HttpResponse(result.getvalue(), content_type = 'application/x-zip-compressed')
		# resp['Content-Disposition'] = 'attachment; filename=%s'%zip_filename

		# return resp

	else:
		return render(request, "index.html")


