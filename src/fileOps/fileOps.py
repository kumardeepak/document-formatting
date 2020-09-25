import os
from zipfile import ZipFile
import uuid

# class FileOperations(object):
#     def __init__(self):
#         self.filename   = None
#         self.filepath   = None

#     def open(self, filepath, working_dir):
#         self.filename   = os.path.basename(filepath)
#         self.filepath   = filepath

#         extract_dir    = os.path.join(working_dir, self.filename)
    
#         with ZipFile(filepath, 'r') as file:
#             file.extractall(path=extract_dir)
#             filenames = file.namelist()
        
#         return extract_dir, filenames

#     def file_xml(filepath):


def extract_docx(filepath, working_dir):
    filename       = os.path.splitext(os.path.basename(filepath))[0]
    extract_dir    = os.path.join(working_dir, filename)
    
    with ZipFile(filepath, 'r') as file:
        file.extractall(path=extract_dir)
        filenames = file.namelist()
    
    return extract_dir, filenames

def save_docx(extracted_dir, filenames, output_filename):
    with ZipFile(output_filename, 'w') as docx:
        for filename in filenames: 
            docx.write(os.path.join(extracted_dir, filename), filename)