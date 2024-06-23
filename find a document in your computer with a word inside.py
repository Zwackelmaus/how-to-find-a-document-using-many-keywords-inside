import os
from docx import Document

def search_docx_files(directory, keywords):
    found_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.docx'):
                file_path = os.path.join(root, file)
                try:
                    document = Document(file_path)
                    print(f"Checking file: {file_path}")
                    text = ""
                    for paragraph in document.paragraphs:
                        text += " " + paragraph.text.lower()
                    if all(keyword.lower() in text for keyword in keywords):
                        found_files.append(file_path)
                    else:
                        print(f"Keywords not found in file: {file_path}")
                except Exception as e:
                    print(f"Error reading file: {file_path}")
                    print(f"Error message: {str(e)}")
    
    return found_files

# Specify the directory and keywords
directory = 'where you think you may save the document'
keywords = ['keyword1', 'keyword2']

# Search for docx files
result = search_docx_files(directory, keywords)

# Display the found files
if result:
    print(f"Found {len(result)} files containing the keywords '{', '.join(keywords)}':")
    for file_path in result:
        print(file_path)
else:
    print(f"No files found containing the keywords '{', '.join(keywords)}' in the specified directory.")
