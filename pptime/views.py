from django.shortcuts import render
from django.http import HttpResponse
from django.http import HttpResponseRedirect
from django.urls import reverse
from pptime.forms import PPTimeReportForm
from pptime.lib.pptime_xlsx import *
from pptime.lib.pplib import *

def index(request):
    return HttpResponse("Hello, world. You're at the pptime index. You have to know the right URL to access to the tool.")


def list(request):
    # Handle file upload
    if request.method == 'POST':
        form = PPTimeReportForm(request.POST)
        if form.is_valid():
            user_comment = form.cleaned_data['comment']
            newdoc = make_pptimereport(comment = user_comment)
            #newdoc = make_testreport(comment = user_comment) #Use this for debug since it is much faster
            newdoc.save()

            # Redirect to the document list after POST
            return HttpResponseRedirect(reverse('pptime:list'))
    else:
        form = PPTimeReportForm() # A empty, unbound form

    # Load documents for the list page
    reports = PPTimeReport.objects.all()
    reports = reports.order_by('-creationdate')

    # Render list page with the documents and the form
    return render(request, 'pptime/list.html', {'reports': reports, 'form': form})
