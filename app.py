from dataclasses import dataclass, field, fields
from io import BytesIO
from pathlib import Path

from flask import Flask, render_template, request, send_file
import openpyxl
from openpyxl.utils import column_index_from_string


TEMPLATE_PATH = Path(__file__).with_name("empty.xlsx")


def write_spaced(ws, start_cell, text, step=4):
    """Write text characters with spacing starting at a given cell."""
    col = ''.join(filter(str.isalpha, start_cell))
    row = int(''.join(filter(str.isdigit, start_cell)))
    start = column_index_from_string(col)
    for i, ch in enumerate((text or "").upper()):
        ws.cell(row=row, column=start + i * step, value=ch)


@dataclass
class AddressData:
    subject_rf: str = ""
    settlement: str = ""
    locality: str = ""
    street: str = ""
    house: str = ""
    apartment: str = ""


def _addr_default():
    return AddressData()


@dataclass
class PersonData:
    surname_ru: str = ""
    surname_lat: str = ""
    name_ru: str = ""
    name_lat: str = ""
    patronymic_ru: str = ""
    patronymic_lat: str = ""
    citizenship: str = ""
    birth_day: str = ""
    birth_month: str = ""
    birth_year: str = ""
    sex: str = "М"
    birth_place: str = ""
    doc_type: str = ""
    doc_series: str = ""
    doc_number: str = ""
    issue_day: str = ""
    issue_month: str = ""
    issue_year: str = ""
    expiry_day: str = ""
    expiry_month: str = ""
    expiry_year: str = ""
    arrival_day: str = ""
    arrival_month: str = ""
    arrival_year: str = ""
    stay_day: str = ""
    stay_month: str = ""
    stay_year: str = ""
    migration_series: str = ""
    migration_number: str = ""
    prev_address: AddressData = field(default_factory=_addr_default)
    reg_address: AddressData = field(default_factory=_addr_default)


@dataclass
class HostData:
    surname: str = ""
    name: str = ""
    patronymic: str = ""
    doc_type: str = ""
    doc_series: str = ""
    doc_number: str = ""
    issue_day: str = ""
    issue_month: str = ""
    issue_year: str = ""
    residence: AddressData = field(default_factory=_addr_default)


def save_to_excel(person: PersonData, host: HostData, output):
    """Fill the Excel template with provided data and write to output."""
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    s1, s2, s3, s4 = wb["стр.1"], wb["стр.2"], wb["стр.3"], wb["стр.4"]

    write_spaced(s1, "N8", person.surname_ru)
    write_spaced(s1, "AL10", person.surname_lat)
    write_spaced(s1, "N12", person.name_ru)
    write_spaced(s1, "AH14", person.name_lat)
    write_spaced(s1, "AD16", person.patronymic_ru)
    write_spaced(s1, "AL18", person.patronymic_lat)
    write_spaced(s1, "V22", person.citizenship)

    write_spaced(s1, "AD25", person.birth_day)
    write_spaced(s1, "AT25", person.birth_month)
    write_spaced(s1, "BF25", person.birth_year)

    if person.sex == "М":
        s1["CL25"], s1["DB25"] = "X", ""
        s3["CX40"], s3["DN40"] = "X", ""
    else:
        s1["CL25"], s1["DB25"] = "", "X"
        s3["CX40"], s3["DN40"] = "", "X"

    write_spaced(s1, "Z27", person.birth_place)

    write_spaced(s1, "J33", person.doc_type)
    write_spaced(s1, "J35", person.doc_series)
    write_spaced(s1, "AP35", person.doc_number)
    write_spaced(s1, "I37", person.issue_day)
    write_spaced(s1, "Z37", person.issue_month)
    write_spaced(s1, "AL37", person.issue_year)
    write_spaced(s1, "BN37", person.expiry_day)
    write_spaced(s1, "CD37", person.expiry_month)
    write_spaced(s1, "CP37", person.expiry_year)

    write_spaced(s1, "I64", person.arrival_day)
    write_spaced(s1, "Z64", person.arrival_month)
    write_spaced(s1, "AL64", person.arrival_year)
    write_spaced(s1, "BN64", person.stay_day)
    write_spaced(s1, "CD64", person.stay_month)
    write_spaced(s1, "CP64", person.stay_year)

    write_spaced(s1, "AH66", person.migration_series)
    write_spaced(s1, "BB66", person.migration_number)

    write_spaced(s2, "B4", person.prev_address.subject_rf)
    write_spaced(s2, "B6", person.prev_address.settlement)
    write_spaced(s2, "B8", person.prev_address.locality)
    write_spaced(s2, "B10", person.prev_address.street)

    if person.prev_address.house:
        s2["B12"] = "ДОМ"
        s2["AT12"] = person.prev_address.house
    else:
        s2["B12"] = s2["AT12"] = ""

    if person.prev_address.apartment:
        s2["B14"] = "КВАРТИРА"
        s2["AT14"] = person.prev_address.apartment
    else:
        s2["B14"] = s2["AT14"] = ""

    write_spaced(s2, "B17", person.reg_address.subject_rf)
    write_spaced(s2, "B19", person.reg_address.settlement)
    write_spaced(s2, "B21", person.reg_address.locality)
    write_spaced(s2, "B23", person.reg_address.street)

    s2["B25"] = "ДОМ"
    s2["AT25"] = person.reg_address.house
    if person.reg_address.apartment:
        s2["B27"] = "КВАРТИРА"
        s2["AT27"] = person.reg_address.apartment
    else:
        s2["B27"] = s2["AT27"] = ""

    write_spaced(s3, "N5", host.surname)
    write_spaced(s3, "N7", host.name)
    write_spaced(s3, "AH9", host.patronymic)
    write_spaced(s3, "J11", host.doc_type)
    write_spaced(s3, "BF11", host.doc_series)
    write_spaced(s3, "BZ11", host.doc_number)
    write_spaced(s3, "I13", host.issue_day)
    write_spaced(s3, "Z13", host.issue_month)
    write_spaced(s3, "AL13", host.issue_year)

    write_spaced(s3, "B16", host.residence.subject_rf)
    write_spaced(s3, "B18", host.residence.settlement)
    write_spaced(s3, "B20", host.residence.locality)
    write_spaced(s3, "B22", host.residence.street)
    s3["B24"] = "ДОМ"
    s3["AT24"] = host.residence.house
    if host.residence.apartment:
        s3["B26"] = "КВАРТИРА"
        s3["AT26"] = host.residence.apartment
    else:
        s3["B26"] = s3["AT26"] = ""

    write_spaced(s3, "N31", person.surname_ru)
    write_spaced(s3, "N33", person.name_ru)
    write_spaced(s3, "AH35", person.patronymic_ru)
    write_spaced(s3, "R37", person.citizenship)
    write_spaced(s3, "AD40", person.birth_day)
    write_spaced(s3, "AX40", person.birth_month)
    write_spaced(s3, "BN40", person.birth_year)

    write_spaced(s3, "J42", person.doc_type)
    write_spaced(s3, "BF42", person.doc_series)
    write_spaced(s3, "CH42", person.doc_number)
    write_spaced(s3, "I44", person.issue_day)
    write_spaced(s3, "Z44", person.issue_month)
    write_spaced(s3, "AL44", person.issue_year)
    write_spaced(s3, "BN44", person.expiry_day)
    write_spaced(s3, "CD44", person.expiry_month)
    write_spaced(s3, "CP44", person.expiry_year)

    write_spaced(s3, "B47", person.reg_address.subject_rf)
    write_spaced(s3, "B49", person.reg_address.settlement)
    write_spaced(s3, "B51", person.reg_address.locality)
    write_spaced(s3, "B53", person.reg_address.street)
    s3["B55"] = "ДОМ"
    s3["AT55"] = person.reg_address.house
    if person.reg_address.apartment:
        s3["B57"] = "КВАРТИРА"
        s3["AT57"] = person.reg_address.apartment
    else:
        s3["B57"] = s3["AT57"] = ""

    write_spaced(s4, "N33", host.surname)
    write_spaced(s4, "N35", host.name)
    write_spaced(s4, "AH37", host.patronymic)

    wb.save(output)
    output.seek(0)
    return output


app = Flask(__name__)


ADDRESS_FIELDS = [f.name for f in fields(AddressData)]
PERSON_FIELDS = [f.name for f in fields(PersonData) if f.type is str]
HOST_FIELDS = [f.name for f in fields(HostData) if f.type is str]

ADDRESS_LABELS = {
    "subject_rf": "Субъект РФ",
    "settlement": "Район",
    "locality": "Населённый пункт",
    "street": "Улица",
    "house": "Дом",
    "apartment": "Квартира",
}

PERSON_LABELS = {
    "surname_ru": "Фамилия",
    "surname_lat": "Фамилия (лат.)",
    "name_ru": "Имя",
    "name_lat": "Имя (лат.)",
    "patronymic_ru": "Отчество",
    "patronymic_lat": "Отчество (лат.)",
    "citizenship": "Гражданство",
    "birth_day": "Дата рождения (день)",
    "birth_month": "Дата рождения (месяц)",
    "birth_year": "Дата рождения (год)",
    "sex": "Пол",
    "birth_place": "Место рождения",
    "doc_type": "Документ: вид",
    "doc_series": "Серия",
    "doc_number": "Номер",
    "issue_day": "Дата выдачи (день)",
    "issue_month": "Дата выдачи (месяц)",
    "issue_year": "Дата выдачи (год)",
    "expiry_day": "Срок действия (день)",
    "expiry_month": "Срок действия (месяц)",
    "expiry_year": "Срок действия (год)",
    "arrival_day": "Дата прибытия (день)",
    "arrival_month": "Дата прибытия (месяц)",
    "arrival_year": "Дата прибытия (год)",
    "stay_day": "Дата убытия (день)",
    "stay_month": "Дата убытия (месяц)",
    "stay_year": "Дата убытия (год)",
    "migration_series": "Серия миграционной карты",
    "migration_number": "Номер миграционной карты",
}

HOST_LABELS = {
    "surname": "Фамилия",
    "name": "Имя",
    "patronymic": "Отчество",
    "doc_type": "Документ: вид",
    "doc_series": "Серия",
    "doc_number": "Номер",
    "issue_day": "Дата выдачи (день)",
    "issue_month": "Дата выдачи (месяц)",
    "issue_year": "Дата выдачи (год)",
}


@app.route("/")
def index():
    return render_template(
        "index.html",
        person_labels=PERSON_LABELS,
        host_labels=HOST_LABELS,
        address_labels=ADDRESS_LABELS,
    )


@app.route("/generate", methods=["POST"])
def generate():
    person_data = {name: request.form.get(f"person_{name}", "") for name in PERSON_FIELDS}
    prev_data = {name: request.form.get(f"person_prev_address_{name}", "") for name in ADDRESS_FIELDS}
    reg_data = {name: request.form.get(f"person_reg_address_{name}", "") for name in ADDRESS_FIELDS}
    host_data = {name: request.form.get(f"host_{name}", "") for name in HOST_FIELDS}
    residence_data = {name: request.form.get(f"host_residence_{name}", "") for name in ADDRESS_FIELDS}

    person = PersonData(
        **person_data,
        prev_address=AddressData(**prev_data),
        reg_address=AddressData(**reg_data),
    )
    host = HostData(**host_data, residence=AddressData(**residence_data))

    output = BytesIO()
    save_to_excel(person, host, output)

    filename = f"{person.surname_ru}_{person.name_ru}.xlsx".upper().replace(" ", "_")
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run()
