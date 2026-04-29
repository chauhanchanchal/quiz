import requests
import json
import os

url = "http://127.0.0.1:5000/generate-exam"

header = {
    "institution": "UIT",
    "degree": "B.Tech",
    "title": "Mock Exam",
    "subject": "CS",
    "date": "2026-05-01",
    "time": "2 Hours",
    "totalMarks": "50"
}

sections = [
    {
        "name": "Section A",
        "type": "MCQ",
        "count": "5",
        "marks": "2",
        "difficulty": "Medium"
    }
]

file_path = "c:/dev/quiz/static/answer_pdfs/bb756e4b-d7ef-4d99-843c-958dde866812.pdf"

print("Generating Question Paper...")
res_paper = requests.post(
    url,
    files={'file': ('bb756e4b.pdf', open(file_path, 'rb'), 'application/pdf')},
    data={
        'header': json.dumps(header),
        'sections': json.dumps(sections),
        'mode': 'paper'
    }
)

if res_paper.status_code == 200:
    print("Paper Success:", res_paper.json())
else:
    print("Paper Error:", res_paper.text)

print("\nGenerating Answer Key...")
res_key = requests.post(
    url,
    files={'file': ('bb756e4b.pdf', open(file_path, 'rb'), 'application/pdf')},
    data={
        'header': json.dumps(header),
        'sections': json.dumps(sections),
        'mode': 'key'
    }
)

if res_key.status_code == 200:
    print("Key Success:", res_key.json())
else:
    print("Key Error:", res_key.text)
