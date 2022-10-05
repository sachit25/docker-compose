from django.urls import path
from . import views         

urlpatterns = [
    path("home/",views.home,name='home'),
    path("base_file/",views.base_file,name='home'),
    path("form/",views.form,name='home'),
    path("login_page/",views.login_page,name='home'),
    path("navbar/",views.navbar,name='home'),
    path("login/",views. login_pagev2,name=' login_pagev2'),
    path("base/",views.base,name="base"),
    path("index/",views.index,name="index"),


    
]    