from jinja2 import Template
from weasyprint import HTML
import pandas as pd
import os
import mail_handler

BASE_URL = os.path.dirname(os.path.abspath( __file__ ))
OPTIONS = {
    'encoding': "UTF-8",
    'default_css': '@page { size: A4; margin: 0; }',
}


event_info = {
    "name": "Hackathon",
    "date": "29.11.2025"
}

def get_participants(table_path: str) -> list[dict]:
    df = pd.read_excel(table_path)
    participants_data = []

    for index, row in df.iterrows():
        participant = {
            'name': row['Имя'],
            'surname': row['Фамилия'],
            'email': row['Email'],
            'language': row['Язык'],
            'role': row['Роль'],
            'place': row['Место']
        }
        participants_data.append(participant)

    return participants_data

def render_pdf(event_data: dict, personal_data: dict):
    templates = {
        "Спикер": "appreciation_letter",
        "Победитель": "winner_diploma",
        "Призёр": "prize_diploma",
        "Участник": "certificate"
    }

    if personal_data["role"] != "Победитель":
        template_name = f"{templates[personal_data['role']]}_{personal_data['language']}"
    elif personal_data["place"] != "-":
        template_name = f"winner_place_diploma_{personal_data['language']}"
    else:
        template_name = f"winner_diploma_{personal_data['language']}"

    with open(f"templates/{template_name}.html", 'r', encoding="utf-8") as f:
        template_content = f.read()
    template = Template(template_content)

    if "winner_place_diploma" in template_name:
        rendered_html = template.render(event_name=event_data["name"],
                                        name=personal_data["name"],
                                        surname=personal_data["surname"],
                                        grade=personal_data["place"],
                                        date=event_data["date"])
    else:
        rendered_html = template.render(event_name=event_data["name"],
                                        surname=personal_data["surname"],
                                        name=personal_data["name"],
                                        date=event_data["date"])

    pdf_dir = event_data["name"]
    if not os.path.exists(pdf_dir):
        os.makedirs(pdf_dir)

    pdf_path = f"{pdf_dir}/{personal_data['name']}_{personal_data['surname']}.pdf"

    HTML(string=rendered_html, base_url=BASE_URL).write_pdf(pdf_path)

    print(f"PDF успешно создан: {pdf_path}")
    return pdf_path

participants = get_participants(os.path.join(BASE_URL, "example_data.xlsx"))

for participant in participants:
    pdf_path = render_pdf(event_info, participant)
    mail_handler.send_gmail(
        participant["email"],
        participant["name"],
        participant["language"].lower(),
        event_info["name"],
        pdf_path
    )