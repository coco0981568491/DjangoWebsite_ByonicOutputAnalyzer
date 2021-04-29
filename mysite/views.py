from django.shortcuts import render, HttpResponse
from django_q.tasks import async_task


# Create your views here.
def home(request):
	if request.method == "POST":

		# get data
		file = request.FILES["myFile"]

		async_task("mysite.services.data_processing", file)

		# Grab ZIP file from in-memory, make response with correct content-type
	    resp = HttpResponse(task.result.getvalue(), content_type = 'application/x-zip-compressed')
	    # ..and correct content-disposition
	    resp['Content-Disposition'] = 'attachment; filename=%s'%zip_filename

	    return resp

	else:
		return render(request, "index.html")

# def upload(request):
# 	return render(request, "fileupload.html")
