import PyPDF2
import sys
import os

#-------------------------------------------------------------
# Check the operating system and create directories on desktop
# TODO: create directories for darwin systems
#-------------------------------------------------------------
if 'win' in sys.platform:
    os.makedirs(r"C:\Users\Kingsley\Desktop\Consents\Template\\", exist_ok=True)
    os.makedirs(r"C:\users\Kingsley\Desktop\Consents\Scanned\\", exist_ok=True)
    os.makedirs(r"C:\Users\Kingsley\Desktop\Consents\Merged\\", exist_ok=True)
    
    template_path = r"C:\Users\Kingsley\Desktop\Consents\Template\template.pdf"
    scanned_path = r"C:\users\Kingsley\Desktop\Consents\Scanned\\"
    merged_path = r"C:\Users\Kingsley\Desktop\Consents\Merged\\"
elif 'darwin' in sys.platform:
    pass
#-------------------------------------------------------------
# Validating location of files
#-------------------------------------------------------------
print("Please make sure that you have correctly placed the template PDF in the "
      "template folder and the scanned PDFs in the Scanned folder.")
merge = input("\nContinue to merge files? Yes or No ( press Y or N ): ")
print()
#-------------------------------------------------------------
# Looping through list of files and reading the first page
# of each scanned file and then writing that page to a new PDF Object
#-------------------------------------------------------------
if merge.lower() == "y":
    scanned_files = []
    for filename in os.listdir(scanned_path):
        if filename.endswith('.pdf'):
            scanned_files.append(filename)
    print(f'You are merging {len(scanned_files)} file(s)')
    print()
    if len(scanned_files) == 0:
        print(f"Merge {len(scanned_files)} file(s) not possible.")
        print("Please check to make sure your folders are not empty and try again")
        print("Exiting...")
        sys.exit()
    for index, filename in enumerate(scanned_files, start=1):
        print(f"{index}...merging {filename} with template")
        scanned_file = open(scanned_path+filename, 'rb')    
        scanned_reader = PyPDF2.PdfFileReader(scanned_path+filename)
        pdfwriter = PyPDF2.PdfFileWriter()
        pdfwriter.addPage(scanned_reader.getPage(0))
    
#-------------------------------------------------------------
# Reading the template file and appending pages 2 to 4 to end
# the PDF object created above
#-------------------------------------------------------------
        template = open(template_path, "rb")
        template_reader = PyPDF2.PdfFileReader(template)
        for pagenumber in range(1, 4):
            pageobject = template_reader.getPage(pagenumber)
            pdfwriter.addPage(pageobject)
#-------------------------------------------------------------
# Reading scanned file and appending all pages except first page
# to the PDF object
#-------------------------------------------------------------
        for pagenumber in range(1, scanned_reader.numPages):
            pageobject = scanned_reader.getPage(pagenumber)
            pdfwriter.addPage(pageobject)
#-------------------------------------------------------------
# Reading the last page of the template file and appending it
# to the PDF object
#-------------------------------------------------------------
        pdfwriter.addPage(template_reader.getPage(6))

#-------------------------------------------------------------
# Creating a new PDF file (with name of original scanned file)
# and writing the contents of the PDF object to the PDF file
#-------------------------------------------------------------
        merged_file = open(merged_path+filename, 'wb')
        pdfwriter.write(merged_file)
#-------------------------------------------------------------
# Necessary housekeeping
#-------------------------------------------------------------
    scanned_file.close()
    template.close()
    merged_file.close()
#-------------------------------------------------------------
    print(f"Merging complete")
else:
    print('OOPS! Exiting >>>')
    sys.exit()

#-------------------------------------------------------------
'''
PDF Merger for METS-KNUST, Department of Physiology, School
of Medical Sciences.
Scripted by: Kingsley Apusiga
Version No.: 001
Version Date: 01-09-2021 (1st September 2021)
Copyright: All Rights Reserved
'''





