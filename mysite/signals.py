from celery.signals import task_success

@task_success.connect
def task_success_handler(sender, result, **kwargs):
    return sender.request.id

