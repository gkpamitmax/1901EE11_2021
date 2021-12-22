from django.urls import path
from .import views

urlpatterns = [
    #home route  function   function name
    path('',  views.main, name='main'),
    path('concise_sheet/',views.concise_sheet,name="concise_sheet"),
    path('send_emails/',views.send_emails,name="send_emails")
]
