from django.shortcuts import render, HttpResponse
from mysite.backend import data_processing

# figure out how to combine rq with django
import redis
from rq import Queue
from . import worker
from worker import *

r = redis.Redis()
q = Queue(connection=conn)


# Create your views here.
def home(request):

	# Folder name in ZIP archive 
	zip_filename = "Results.zip"

	if request.method == "POST":

		# get data and name it as file for convenience 
		file = request.FILES["myFile"]

		# trigger the backend to process the input file
		job = q.enqueue(data_processing, file)

		q_len = len(q)

		if job.result == None:
			return f"Task {job.id} added to the queue at {job.enqueue_at}. {q_len} tasks in the queue"

		else:
			resp = HttpResponse(job.result, content_type = 'application/x-zip-compressed')
			resp['Content-Disposition'] = 'attachment; filename=%s'%zip_filename

			return resp

		
	else:
		return render(request, "index.html")


