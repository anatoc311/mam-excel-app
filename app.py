import sys, json
from pathlib import Path
from dataclasses import dataclass, field, asdict

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget, QScrollArea,
    QFormLayout, QLineEdit, QGroupBox, QVBoxLayout, QHBoxLayout,
    QLabel, QRadioButton, QButtonGroup, QPushButton, QMessageBox,
    QFileDialog
)
from PyQt6.QtGui import QIntValidator
import openpyxl
from openpyxl.utils import column_index_from_string


TEMPLATE_PATH = Path(__file__).with_name("empty.xlsx")

# default dirs for dialogs
EXCEL_DIR = Path(r"C:\Users\Anton\Downloads\Мамино")
JSON_DIR  = EXCEL_DIR / "Файлы с данными"


def upper_line(maxlen=None):
    le = QLineEdit()
    if maxlen:
        le.setMaxLength(maxlen)
    le.editingFinished.connect(lambda le=le: le.setText(le.text().upper()))
    return le


def int_line(min_v, max_v):
    le = QLineEdit()
    le.setValidator(QIntValidator(min_v, max_v))
    le.setMaxLength(len(str(max_v)))
    return le


def write_spaced(ws, start_cell, text, step=4):
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


class AddressWidget(QGroupBox):
    def __init__(self, title):
        super().__init__(title)
        self._subject = upper_line()
        self._settlement = upper_line()
        self._locality = upper_line()
        self._street = upper_line()
        self._house = upper_line()
        self._apartment = upper_line()
        l = QFormLayout()
        l.addRow("Субъект РФ", self._subject)
        l.addRow("Район", self._settlement)
        l.addRow("Населённый пункт", self._locality)
        l.addRow("Улица", self._street)
        l.addRow("Дом ", self._house)
        l.addRow("Квартира ", self._apartment)
        self.setLayout(l)

    def get_data(self):
        return AddressData(
            subject_rf=self._subject.text(),
            settlement=self._settlement.text(),
            locality=self._locality.text(),
            street=self._street.text(),
            house=self._house.text(),
            apartment=self._apartment.text()
        )

    def set_data(self, d: AddressData):
        self._subject.setText(d.subject_rf)
        self._settlement.setText(d.settlement)
        self._locality.setText(d.locality)
        self._street.setText(d.street)
        self._house.setText(d.house)
        self._apartment.setText(d.apartment)


class PersonTab(QWidget):
    def __init__(self):
        super().__init__()
        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        container = QWidget(); f = QFormLayout(container)

        self._surname_ru = upper_line()
        self._surname_lat = upper_line()
        self._name_ru = upper_line()
        self._name_lat = upper_line()
        self._patronymic_ru = upper_line()
        self._patronymic_lat = upper_line()

        self._citizenship = upper_line()
        self._birth_day = int_line(1, 31)
        self._birth_month = int_line(1, 12)
        self._birth_year = int_line(1900, 2100)

        male = QRadioButton("Мужской"); female = QRadioButton("Женский")
        male.setChecked(True)
        self._sex_group = QButtonGroup()
        self._sex_group.addButton(male, 0); self._sex_group.addButton(female, 1)
        sex_box = QHBoxLayout(); sex_box.addWidget(male); sex_box.addWidget(female)
        sex_widget = QWidget(); sex_widget.setLayout(sex_box)

        self._birth_place = upper_line()
        self._doc_type = upper_line()
        self._doc_series = upper_line()
        self._doc_number = upper_line()
        self._issue_day = int_line(1, 31)
        self._issue_month = int_line(1, 12)
        self._issue_year = int_line(1900, 2100)
        self._exp_day = int_line(1, 31)
        self._exp_month = int_line(1, 12)
        self._exp_year = int_line(1900, 2100)

        self._mig_series = upper_line(10)
        self._mig_number = upper_line(20)

        self._arrival_day = int_line(1, 31)
        self._arrival_month = int_line(1, 12)
        self._arrival_year = int_line(1900, 2100)
        self._stay_day = int_line(1, 31)
        self._stay_month = int_line(1, 12)
        self._stay_year = int_line(1900, 2100)

        f.addRow("Фамилия", self._surname_ru)
        f.addRow("Фамилия (лат.)", self._surname_lat)
        f.addRow("Имя", self._name_ru)
        f.addRow("Имя (лат.)", self._name_lat)
        f.addRow("Отчество", self._patronymic_ru)
        f.addRow("Отчество (лат.)", self._patronymic_lat)
        f.addRow("Гражданство", self._citizenship)

        bb = QHBoxLayout()
        bb.addWidget(QLabel("Д")); bb.addWidget(self._birth_day)
        bb.addWidget(QLabel("М")); bb.addWidget(self._birth_month)
        bb.addWidget(QLabel("Г")); bb.addWidget(self._birth_year)
        bw = QWidget(); bw.setLayout(bb)
        f.addRow("Дата рождения", bw)

        f.addRow("Пол", sex_widget)
        f.addRow("Место рождения", self._birth_place)
        f.addRow("Документ: вид", self._doc_type)
        f.addRow("серия", self._doc_series)
        f.addRow("номер", self._doc_number)

        ib = QHBoxLayout()
        ib.addWidget(QLabel("Д")); ib.addWidget(self._issue_day)
        ib.addWidget(QLabel("М")); ib.addWidget(self._issue_month)
        ib.addWidget(QLabel("Г")); ib.addWidget(self._issue_year)
        iw = QWidget(); iw.setLayout(ib)
        f.addRow("Дата выдачи", iw)

        eb = QHBoxLayout()
        eb.addWidget(QLabel("Д")); eb.addWidget(self._exp_day)
        eb.addWidget(QLabel("М")); eb.addWidget(self._exp_month)
        eb.addWidget(QLabel("Г")); eb.addWidget(self._exp_year)
        ew = QWidget(); ew.setLayout(eb)
        f.addRow("Срок действия", ew)

        f.addRow("Серия мигр. карты", self._mig_series)
        f.addRow("Номер мигр. карты", self._mig_number)

        ab = QHBoxLayout()
        ab.addWidget(QLabel("Д")); ab.addWidget(self._arrival_day)
        ab.addWidget(QLabel("М")); ab.addWidget(self._arrival_month)
        ab.addWidget(QLabel("Г")); ab.addWidget(self._arrival_year)
        aw = QWidget(); aw.setLayout(ab)
        f.addRow("Дата заезда", aw)

        sb = QHBoxLayout()
        sb.addWidget(QLabel("Д")); sb.addWidget(self._stay_day)
        sb.addWidget(QLabel("М")); sb.addWidget(self._stay_month)
        sb.addWidget(QLabel("Г")); sb.addWidget(self._stay_year)
        sw = QWidget(); sw.setLayout(sb)
        f.addRow("Срок пребывания до", sw)

        self._prev_addr = AddressWidget("Адрес прежнего места пребывания")
        self._reg_addr = AddressWidget("Адрес места пребывания")
        f.addRow(self._prev_addr); f.addRow(self._reg_addr)

        scroll.setWidget(container)
        l = QVBoxLayout(self); l.addWidget(scroll)

    def get_data(self):
        sex = "М" if self._sex_group.checkedId() == 0 else "Ж"
        return PersonData(
            surname_ru=self._surname_ru.text(), surname_lat=self._surname_lat.text(),
            name_ru=self._name_ru.text(), name_lat=self._name_lat.text(),
            patronymic_ru=self._patronymic_ru.text(), patronymic_lat=self._patronymic_lat.text(),
            citizenship=self._citizenship.text(),
            birth_day=self._birth_day.text(), birth_month=self._birth_month.text(), birth_year=self._birth_year.text(),
            sex=sex, birth_place=self._birth_place.text(),
            doc_type=self._doc_type.text(), doc_series=self._doc_series.text(), doc_number=self._doc_number.text(),
            issue_day=self._issue_day.text(), issue_month=self._issue_month.text(), issue_year=self._issue_year.text(),
            expiry_day=self._exp_day.text(), expiry_month=self._exp_month.text(), expiry_year=self._exp_year.text(),
            arrival_day=self._arrival_day.text(), arrival_month=self._arrival_month.text(), arrival_year=self._arrival_year.text(),
            stay_day=self._stay_day.text(), stay_month=self._stay_month.text(), stay_year=self._stay_year.text(),
            migration_series=self._mig_series.text(), migration_number=self._mig_number.text(),
            prev_address=self._prev_addr.get_data(), reg_address=self._reg_addr.get_data()
        )

    def set_data(self, d: PersonData):
        self._surname_ru.setText(d.surname_ru); self._surname_lat.setText(d.surname_lat)
        self._name_ru.setText(d.name_ru); self._name_lat.setText(d.name_lat)
        self._patronymic_ru.setText(d.patronymic_ru); self._patronymic_lat.setText(d.patronymic_lat)
        self._citizenship.setText(d.citizenship)
        self._birth_day.setText(d.birth_day); self._birth_month.setText(d.birth_month); self._birth_year.setText(d.birth_year)
        (self._sex_group.buttons()[0] if d.sex == "М" else self._sex_group.buttons()[1]).setChecked(True)
        self._birth_place.setText(d.birth_place)
        self._doc_type.setText(d.doc_type); self._doc_series.setText(d.doc_series); self._doc_number.setText(d.doc_number)
        self._issue_day.setText(d.issue_day); self._issue_month.setText(d.issue_month); self._issue_year.setText(d.issue_year)
        self._exp_day.setText(d.expiry_day); self._exp_month.setText(d.expiry_month); self._exp_year.setText(d.expiry_year)
        self._arrival_day.setText(d.arrival_day); self._arrival_month.setText(d.arrival_month); self._arrival_year.setText(d.arrival_year)
        self._stay_day.setText(d.stay_day); self._stay_month.setText(d.stay_month); self._stay_year.setText(d.stay_year)
        self._mig_series.setText(d.migration_series); self._mig_number.setText(d.migration_number)
        self._prev_addr.set_data(d.prev_address); self._reg_addr.set_data(d.reg_address)


class HostTab(QWidget):
    def __init__(self):
        super().__init__()
        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        container = QWidget(); f = QFormLayout(container)
        self._surname = upper_line()
        self._name = upper_line()
        self._patronymic = upper_line()
        self._doc_type = upper_line()
        self._doc_series = upper_line()
        self._doc_number = upper_line()
        self._issue_day = int_line(1, 31)
        self._issue_month = int_line(1, 12)
        self._issue_year = int_line(1900, 2100)

        f.addRow("Фамилия", self._surname); f.addRow("Имя", self._name); f.addRow("Отчество", self._patronymic)
        f.addRow("Документ: вид", self._doc_type); f.addRow("серия", self._doc_series); f.addRow("номер", self._doc_number)

        ib = QHBoxLayout()
        ib.addWidget(QLabel("Д")); ib.addWidget(self._issue_day)
        ib.addWidget(QLabel("М")); ib.addWidget(self._issue_month)
        ib.addWidget(QLabel("Г")); ib.addWidget(self._issue_year)
        iw = QWidget(); iw.setLayout(ib)
        f.addRow("Дата выдачи", iw)

        self._res_addr = AddressWidget("Адрес проживания принимающей стороны")
        f.addRow(self._res_addr)

        scroll.setWidget(container); l = QVBoxLayout(self); l.addWidget(scroll)

    def get_data(self):
        return HostData(
            surname=self._surname.text(), name=self._name.text(), patronymic=self._patronymic.text(),
            doc_type=self._doc_type.text(), doc_series=self._doc_series.text(), doc_number=self._doc_number.text(),
            issue_day=self._issue_day.text(), issue_month=self._issue_month.text(), issue_year=self._issue_year.text(),
            residence=self._res_addr.get_data()
        )

    def set_data(self, d: HostData):
        self._surname.setText(d.surname); self._name.setText(d.name); self._patronymic.setText(d.patronymic)
        self._doc_type.setText(d.doc_type); self._doc_series.setText(d.doc_series); self._doc_number.setText(d.doc_number)
        self._issue_day.setText(d.issue_day); self._issue_month.setText(d.issue_month); self._issue_year.setText(d.issue_year)
        self._res_addr.set_data(d.residence)


def save_to_excel(person, host, output_path: Path):
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    s1, s2, s3, s4 = wb["стр.1"], wb["стр.2"], wb["стр.3"], wb["стр.4"]

    write_spaced(s1, "N8",  person.surname_ru)
    write_spaced(s1, "AL10", person.surname_lat)
    write_spaced(s1, "N12",  person.name_ru)
    write_spaced(s1, "AH14", person.name_lat)
    write_spaced(s1, "AD16", person.patronymic_ru)
    write_spaced(s1, "AL18", person.patronymic_lat)
    write_spaced(s1, "V22",  person.citizenship)

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

    write_spaced(s1, "J33",  person.doc_type)
    write_spaced(s1, "J35",  person.doc_series)
    write_spaced(s1, "AP35", person.doc_number)
    write_spaced(s1, "I37",  person.issue_day)
    write_spaced(s1, "Z37",  person.issue_month)
    write_spaced(s1, "AL37", person.issue_year)
    write_spaced(s1, "BN37", person.expiry_day)
    write_spaced(s1, "CD37", person.expiry_month)
    write_spaced(s1, "CP37", person.expiry_year)

    write_spaced(s1, "I64",  person.arrival_day)
    write_spaced(s1, "Z64",  person.arrival_month)
    write_spaced(s1, "AL64", person.arrival_year)
    write_spaced(s1, "BN64", person.stay_day)
    write_spaced(s1, "CD64", person.stay_month)
    write_spaced(s1, "CP64", person.stay_year)

    write_spaced(s1, "AH66", person.migration_series)
    write_spaced(s1, "BB66", person.migration_number)

    write_spaced(s2, "B4",  person.prev_address.subject_rf)
    write_spaced(s2, "B6",  person.prev_address.settlement)
    write_spaced(s2, "B8",  person.prev_address.locality)
    write_spaced(s2, "B10", person.prev_address.street)

    if person.prev_address.house:
        s2["B12"]  = "ДОМ"
        s2["AT12"] = person.prev_address.house
    else:
        s2["B12"] = s2["AT12"] = ""

    if person.prev_address.apartment:
        s2["B14"]  = "КВАРТИРА"
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
        s2["B27"]  = "КВАРТИРА"
        s2["AT27"] = person.reg_address.apartment
    else:
        s2["B27"] = s2["AT27"] = ""

    write_spaced(s3, "N5",  host.surname)
    write_spaced(s3, "N7",  host.name)
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
    s3["B24"]  = "ДОМ"
    s3["AT24"] = host.residence.house
    if host.residence.apartment:
        s3["B26"]  = "КВАРТИРА"
        s3["AT26"] = host.residence.apartment
    else:
        s3["B26"] = s3["AT26"] = ""

    write_spaced(s3, "N31", person.surname_ru)
    write_spaced(s3, "N33", person.name_ru)
    write_spaced(s3, "AH35", person.patronymic_ru)
    write_spaced(s3, "R37",  person.citizenship)
    write_spaced(s3, "AD40", person.birth_day)
    write_spaced(s3, "AX40", person.birth_month)
    write_spaced(s3, "BN40", person.birth_year)

    write_spaced(s3, "J42",  person.doc_type)
    write_spaced(s3, "BF42", person.doc_series)
    write_spaced(s3, "CH42", person.doc_number)
    write_spaced(s3, "I44",  person.issue_day)
    write_spaced(s3, "Z44",  person.issue_month)
    write_spaced(s3, "AL44", person.issue_year)
    write_spaced(s3, "BN44", person.expiry_day)
    write_spaced(s3, "CD44", person.expiry_month)
    write_spaced(s3, "CP44", person.expiry_year)

    write_spaced(s3, "B47", person.reg_address.subject_rf)
    write_spaced(s3, "B49", person.reg_address.settlement)
    write_spaced(s3, "B51", person.reg_address.locality)
    write_spaced(s3, "B53", person.reg_address.street)
    s3["B55"]  = "ДОМ"
    s3["AT55"] = person.reg_address.house
    if person.reg_address.apartment:
        s3["B57"]  = "КВАРТИРА"
        s3["AT57"] = person.reg_address.apartment
    else:
        s3["B57"] = s3["AT57"] = ""

    write_spaced(s4, "N33", host.surname)
    write_spaced(s4, "N35", host.name)
    write_spaced(s4, "AH37", host.patronymic)

    wb.save(output_path)
    return output_path


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Регистрационные сведения")

        self._person_tab = PersonTab(); self._host_tab = HostTab()
        tabs = QTabWidget(); tabs.addTab(self._person_tab, "Прописываемый"); tabs.addTab(self._host_tab, "Принимающий")

        save_btn_excel = QPushButton("Сохранить в Excel"); save_btn_excel.clicked.connect(self._handle_save_excel)
        save_btn_data  = QPushButton("Сохранить данные"); save_btn_data.clicked.connect(self._handle_save_data)
        load_btn_data  = QPushButton("Загрузить данные"); load_btn_data.clicked.connect(self._handle_load_data)

        row = QHBoxLayout(); row.addStretch(); row.addWidget(save_btn_data); row.addWidget(load_btn_data)
        central = QWidget(); lay = QVBoxLayout(central)
        lay.addLayout(row); lay.addWidget(tabs); lay.addWidget(save_btn_excel)

        self.setCentralWidget(central); self.resize(800, 600)

    def _handle_save_excel(self):
        person = self._person_tab.get_data()
        host   = self._host_tab.get_data()
        default_name = f"{person.surname_ru}_{person.name_ru}".upper().replace(" ", "_") + ".xlsx"
        EXCEL_DIR.mkdir(parents=True, exist_ok=True)
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить Excel-файл", str(EXCEL_DIR / default_name), "Excel (*.xlsx)"
        )
        if not file_path:
            return
        save_to_excel(person, host, Path(file_path))
        QMessageBox.information(self, "Готово", f"Файл сохранён:\n{file_path}")

    def _handle_save_data(self):
        person = self._person_tab.get_data()
        host   = self._host_tab.get_data()
        default_name = f"{person.surname_ru}_{person.name_ru}".upper().replace(" ", "_") + ".json"
        JSON_DIR.mkdir(parents=True, exist_ok=True)
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить данные формы", str(JSON_DIR / default_name), "JSON (*.json)"
        )
        if not file_path:
            return
        data = {"person": asdict(person), "host": asdict(host)}
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        QMessageBox.information(self, "Готово", f"Данные сохранены:\n{file_path}")

    def _handle_load_data(self):
        JSON_DIR.mkdir(parents=True, exist_ok=True)
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Открыть данные формы", str(JSON_DIR), "JSON (*.json)"
        )
        if not file_path:
            return
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            p = data["person"]; h = data["host"]
            person = PersonData(
                **{k: v for k, v in p.items() if k not in ("prev_address", "reg_address")},
                prev_address=AddressData(**p["prev_address"]),
                reg_address=AddressData(**p["reg_address"]),
            )
            host = HostData(
                **{k: v for k, v in h.items() if k != "residence"},
                residence=AddressData(**h["residence"]),
            )
            self._person_tab.set_data(person)
            self._host_tab.set_data(host)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить файл:\n{e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = MainWindow(); w.show()
    sys.exit(app.exec())
