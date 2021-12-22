from django.contrib import admin
from django.urls import path
from . import views

urlpatterns = [
   
    path('', views.home,name='home'),
    path('/range_generator', views.range_generator,name='range_generator'),
    path('/generate_all_roll_transcript', views.generate_all_roll_transcript,name='generate_all_roll_transcript'),
]