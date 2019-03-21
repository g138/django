# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import magic
from django.shortcuts import render
from wand.image import Image as wi
import subprocess
#from zipfile import *
import zipfile
import os
import StringIO
from django.http import HttpResponse
import xml.etree.cElementTree as ET
from lxml import etree
import datetime
import pytz
import shutil
from django.core.files.storage import FileSystemStorage
# from django.core.files.storage import FileSystemStorage
# Create your views here.

def home(request):
    return render(request, 'upload/home.html')


    

def file_file(request):
    context = {}
    if request.method == 'POST':
        uploaded_file = request.FILES['document']
        mime = magic.from_buffer(uploaded_file.read(),mime=True)
        
        
        fs = FileSystemStorage()
        name = fs.save(uploaded_file.name, uploaded_file)
        context['url'] = fs.url(name)
        cur_dir = os.getcwd()
        print(cur_dir)
        command = 'cp '+cur_dir+'/media/'+uploaded_file.name+' '+cur_dir
        print('===============>',command)
        subprocess.call(command,shell=True)


        
        if (mime == 'application/vnd.openxmlformats-officedocument.presentationml.presentation' or mime == 'application/vnd.ms-powerpoint'):

            command_name = 'unoconv -f pdf ' + uploaded_file.name
            subprocess.call(command_name, shell=True)
            file_name_mime = uploaded_file.name

            if (mime == 'application/vnd.openxmlformats-officedocument.presentationml.presentation'):
                file_name_without_mime = file_name_mime[:-5]
            if(mime == 'application/vnd.ms-powerpoint'):
                file_name_without_mime = file_name_mime[:-4]
            
            date_time = datetime.datetime.now(tz=pytz.UTC)
            date_time = str(date_time)
            attr_qname = etree.QName("http://www.w3.org/2001/XMLSchema-instance","schemaLocation")
            nsmap = {"xsi":"http://www.w3.org/2001/XMLSchema-instance","xsd":"http://www.w3.org/2001/XMLSchema"}
            root1 = etree.Element("SWQuestionnaire",nsmap=nsmap)
            etree.SubElement(root1, "mTitle").text = "5 Moments Lesson"
            etree.SubElement(root1, "mCreationDate").text = date_time
            etree.SubElement(root1, "mCultureCode").text = "en-IE"
            tree = etree.ElementTree(root1)
            
            tree.write("SWLesson.xml",encoding='utf-8', xml_declaration=True)
            
            etree.SubElement(root1, "m_lsSessions").text =""
            tree = etree.ElementTree(root1)
            tree.write("SWLessonResponses.xml",encoding='utf-8', xml_declaration=True)


            file_name_pdf = file_name_without_mime + '.pdf'
            pdf = wi(filename=file_name_pdf, resolution=300)
            pdfImage = pdf.convert('jpeg')
            i=1
            filelist = []
            for img in pdfImage.sequence:
                page = wi(image=img)
                page.save(filename='Slide'+str(i)+'.jpg')
                filelist.append('Slide'+str(i)+'.jpg')
                i += 1
            filelist.append('SWLesson.xml')
            filelist.append('SWLessonResponses.xml')
            print(filelist)


            zip_subdir = file_name_without_mime
            zip_filename = "%s.zip"%zip_subdir
            s = StringIO.StringIO()
            zip_file = zipfile.ZipFile(s,"w")
            for fpath in filelist:
                fdir, fname = os.path.split(fpath)
                zip_path = os.path.join(zip_subdir, fname)
                zip_file.write(fpath, zip_path)
            zip_file.close()
            response = HttpResponse(s.getvalue(), content_type='application/zip')
            response['Content-Disposition'] = 'attachment;filename=%s' %zip_filename
            for i in filelist:
                os.remove(i)
            os.remove(uploaded_file.name)
            os.remove(file_name_pdf)
            shutil.rmtree('media')
            return response
            context={
                'mime':'File Type : Unexpected!!'
            }
            
    return render(request, 'upload/file.html', context)