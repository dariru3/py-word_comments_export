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
                            comment_text = "".join(element.itertext()).strip()
                            comments.append((comment_id, comment_text))
                            element.clear()
    return comments

docx_file = "comments.docx"
comments = extract_comments(docx_file)

for comment_id, comment_text in comments:
    print(f"Comment ID: {comment_id}, Comment Text: {comment_text}")
