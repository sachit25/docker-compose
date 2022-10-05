import re
from django.shortcuts import render
from django.http import HttpResponse
from django.core.files.storage import FileSystemStorage
from .models import File_Upload
def test2(request):
    return HttpResponse("test3--------------")

def index(request):
    return render(request,'index.html')

def base(request):
    return HttpResponse("-------Base template-------")    

# Create your views here.
# def home(request):
#     if request.method == 'POST':
#         file=request.FILES['myfile']
#         if file:
#             print("---file---")
#         print(dir(file),"---type---")
#         file_name=file._name
#         print(file.__sizeof__())
#         print("---post---")
#         print(file,"--file---")
#         return HttpResponse("File Uploaded")
#     if request.method == 'GET':
#         return render(request,'homepage.html')

def login_page(request):
    return render(request,'login_page.html')

def form(request):
    return render(request,'form.html')

def home(request):
    context = {}
    if request.method == 'POST':
        uploaded_file = request.FILES['myfile']
        if uploaded_file:
            print(uploaded_file.name,"------name-----")
            fs = FileSystemStorage()
            name = fs.save(uploaded_file.name, uploaded_file)
            context['url'] = fs.url(name)
            return render(request, 'base.html', context)
        else:
            return render(request,'base.html')    
    
    if request.method == 'GET':
        print("--get request----")
        return render(request,'base.html')

def base_file(request):
    return render(request,"upload.html")   

def login_pagev2(request):
    if request.method == 'POST':
       
        eml=request.POST.get('email')
        psw = request.POST.get('password')
        print(eml,psw,"--2--")
        return render(request,"login_pagev2.html",context={'error':'error'})
    if request.method == 'GET':
        return render(request,"login_pagev2.html")
    
def navbar(request):
    return render(request,'navbar.html')         

def middleware_custom(get_response):
    def func1(request):
        response=get_response(request)
        return response
    return middleware_custom       
        