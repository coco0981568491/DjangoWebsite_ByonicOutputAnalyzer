from django.shortcuts import render
from django_q.tasks import async_task


# Create your views here.
def home(request):
	if request.method == "POST":

		# get data
		file = request.FILES["myFile"]

		async_task("mysite.services.data_processing", file)
		
	else:
		return render(request, "index.html")

# def upload(request):
# 	return render(request, "fileupload.html")
