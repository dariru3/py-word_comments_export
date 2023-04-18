import zipfile
from xml.etree.ElementTree import iterparse

def extract_comments(docx_file):
    comments = []

    # Open the Word document as a ZIP archive and read the comments file
    with zipfile.ZipFile(docx_file) as zfile:
        for filename in zfile.namelist():
            if filename == "word/comments.xml":
                with zfile.open(filename) as comments_file:
                    # Parse the comments.xml file and extract comments
                    for _, element in iterparse(comments_file):
                        if element.tag.endswith("comment"):
                            comment_id = element.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id")
                            author_name = element.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author")
                            comment_text = "".join(element.itertext()).strip()
                            comments.append((comment_id, author_name, comment_text))
                            element.clear()
    return comments

docx_file = "comments.docx"
output_file = "output_comments.txt"

comments = extract_comments(docx_file)

with open(output_file, 'w', encoding='utf-8') as f:
    for comment_id, author_name, comment_text in comments:
        f.write(f"Comment ID: {comment_id}, Author: {author_name}, Comment Text: {comment_text}\n")