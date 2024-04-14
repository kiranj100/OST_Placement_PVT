import re
import pdfplumber
import xlwt
from django.http import HttpResponse, HttpResponseBadRequest
from django.shortcuts import render, redirect
from .forms import UploadFileForm
from .models import UploadedFile


def upload_file(request):
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            uploaded_file = form.save(commit=False)
            if uploaded_file.pdf_file.name.lower().endswith('.pdf'):
                uploaded_file.save()
                return redirect('convert_to_excel')
            else:
                return render(request, template_name="error.html")
    else:
        form = UploadFileForm()
    return render(request, 'upload.html', context= {'form': form})


def extract_candidate_info(text):

    name_regex = r'([A-Z][a-z]+(?: [A-Z][a-z]+)*)'
    email_regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    phone_regex = r'\b\d{3}[-.\s]?\d{3}[-.\s]?\d{4}\b'

    names = re.findall(name_regex, text)
    emails = re.findall(email_regex, text)
    phone_numbers = re.findall(phone_regex, text)

    return names, emails, phone_numbers


def convert_to_excel(request):
    uploaded_file = UploadedFile.objects.last()
    pdf_path = uploaded_file.pdf_file.path
    excel_path = pdf_path.replace('.pdf', '_extracted.xls')

    with pdfplumber.open(pdf_path) as pdf:
        all_candidate_info = []

        for page in pdf.pages:
            text = page.extract_text()

            names, emails, phone_numbers = extract_candidate_info(text)

            candidate_info = zip(names, emails, phone_numbers)
            all_candidate_info.extend(candidate_info)


    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('Sheet 1')
    worksheet.write(0, 0, 'Candidate Name')
    worksheet.write(0, 1, 'Emails')
    worksheet.write(0, 2, 'Phone Numbers')

    row = 1
    for candidate in all_candidate_info:
        worksheet.write(row, 0, candidate[0])
        worksheet.write(row, 1, candidate[1])
        worksheet.write(row, 2, candidate[2])
        row += 1

    workbook.save(excel_path)

    response = HttpResponse(open(excel_path, 'rb'), content_type='application/ms-excel')
    response[
        'Content-Disposition'] = f'attachment; filename="{uploaded_file.pdf_file.name.replace(".pdf", "_extracted.xls")}"'
    return response
