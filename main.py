from jinja2 import Template
from weasyprint import HTML
import os

BASE_URL = os.path.dirname(os.path.abspath( __file__ ))

options = {
    'encoding': "UTF-8",
    'default_css': '@page { size: A4; margin: 0; }',
}

personal_data = {
    "name": "NAME",
    "surname": "SURNAME",
    "email": "EMAIL",
    "language": "LANGUAGE",
    "role": "Призёр",
    "place": "-"
}

event_info = {
    "name": "TITLE",
    "date": "DATE"
}


def render_pdf(event_info: dict, person_data: dict):
    templates = {
        "Спикер": "appreciation_letter",
        "Победитель": "winner_diploma",
        "Призёр": "prize_diploma",
        "Участник": "certificate"
    }

    if personal_data["role"] != "Победитель":
        template_name = templates[person_data["role"]]
    elif person_data["place"] != "-":
        template_name = "winner_place_diploma"
    else:
        template_name = "winner_diploma"

    with open(f"templates/{template_name}.html", 'r', encoding="utf-8") as f:
        template_content = f.read()
    template = Template(template_content)


    if template_name == "winner_place_diploma":
        rendered_html = template.render(event_name=event_info["name"],
                                        name=person_data["name"],
                                        surname=person_data["surname"],
                                        grade=person_data["place"],
                                        date=event_info["date"])
    else:
        rendered_html = template.render(event_name=event_info["name"],
                                        surname=person_data["surname"],
                                        name=person_data["name"],
                                        date=event_info["date"])

    pdf_dir = event_info["name"]
    if not os.path.exists(pdf_dir):
        os.makedirs(pdf_dir)

    pdf_path = f"{pdf_dir}/{person_data["name"]}_{person_data["surname"]}.pdf"


    HTML(string=rendered_html, base_url=BASE_URL).write_pdf(pdf_path)

    print(f"PDF успешно создан: {pdf_path}")


render_pdf(event_info, personal_data)
