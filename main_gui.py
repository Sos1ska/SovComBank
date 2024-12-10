import json, os, requests, bs4, getpass
from tkinter import *
from tkinter import ttk
from ctypes import windll

LOG_GUI = []
LOG_OPERATOR = 0
LOG_DATA = []
LOG_CR = []
LOG_BONUS = 0
RAM = []
RAM_SECOND_SLOT = []
RAM_THIRD_SLOT = []

# Название файлов и также расценки на договора
class data:
    name_98 = "98_OperatorAggregates.xml"
    dash = "dash.json"
    bonus = "agree.xml"
    FP = {
        "База":180,
        "Дисконт":150,
        "Повышенный":210,
        "Без ФЗ":50
    }
    TP = {
        "0.00":0,
        "1999.00":40,
        "3999.00":60,
        "4999.00":100,
        "5999.00":120,
        "7999.00":140,
        "9999.00":150,
        "14999.00":200,
        "19999.00":250,
        "24999.00":450,
        "29999.00":500,
        "39999.00":600,
        "49999.00":700,
        "59999.00":750,
        "69999.00":800
    }

# Класс с основой системой подрузки данных + подсчёт
class TruAR(data):
    #username = "volkovva"
    # Открытие нужного отчёта и передача данных системе
    def __init__(self, groupName):
        self.username = "homer" if groupName == "test" else groupName
    def _open_report(self, report):
        def find(file, path) -> bool:
            try:
                _files = os.listdir(path)
                if file in _files : return True
                else : return False
            except FileNotFoundError:
                #log.append('File Not Found')
                return False
        match report:
            case "98":
                if find(super().name_98, "C:\\Users\\%s\\Desktop\\Auto\\" % (self.username)) == True:
                    with open("C:\\Users\\%s\\Desktop\\Auto\\%s" % (self.username, super().name_98), "r", encoding="utf-8") as file : get_data = file.read()
                    return get_data
                else:
                    os.system("""powershell -Command "& {[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); [System.Windows.Forms.MessageBox]::Show('FileNotFound in path %s', 'TruAR', 'OK', 'Information')}""""" % ("C:\\Users\\%s\\Desktop\\Auto\\" % (self.username) + super().name_98))
                    os._exit(0)
            case "dash":
                if find(super().dash, "C:\\Users\\%s\\Desktop\\Auto\\" % (self.username)) == True:
                    with open("C:\\Users\\%s\\Desktop\\Auto\\%s" % (self.username, super().dash), "r", encoding="utf-8") as file : get_data = json.load(file)
                    return get_data
                else:
                    os.system("""powershell -Command "& {[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); [System.Windows.Forms.MessageBox]::Show('FileNotFound in path %s', 'TruAR', 'OK', 'Information')}""""" % ("C:\\Users\\%s\\Desktop\\Auto\\" % (self.username) + super().dash))
                    os._exit(0)
            case "agreements":
                if find(super().bonus, "C:\\Users\\%s\\Desktop\\Auto\\" % (self.username)) == True:
                    with open("C:\\Users\\%s\\Desktop\\Auto\\%s" % (self.username, super().bonus), "r", encoding="utf-8") as file : get_data = file.read()
                    return get_data
                else:
                    os.system("""powershell -Command "& {[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); [System.Windows.Forms.MessageBox]::Show('FileNotFound in path %s', 'TruAR', 'OK', 'Information')}""""" % ("C:\\Users\\%s\\Desktop\\Auto\\" % (self.username) + super().bonus))
                    os._exit(0)
    # Парсинг данных из модуля _open_report
    def get_datas(self, get_group, report) -> True | False:
        global LOG_OPERATOR
        global LOG_DATA
        get_data = self._open_report(report)
        get_group = self.username if self.username != "homer" else "volkovva"
        if report == "98" or report == "agreements" : data = bs4.BeautifulSoup(get_data, 'xml')
        elif report == "dash" : data = get_data
        with open("groups.json", "r", encoding="utf-8") as file_groups : data_group = json.load(file_groups)
        group = data_group["%s" % (get_group)]
        if report == "98" or report == "dash":
            for operator in group:
                if report == "98": 
                    search = data.find('AuthorName', {'AuthorName': operator["name"]})
                    try:
                        print('Сотрудник: %s\nОбработал: %s\nСогласен: %s\nПодписал: %s\nОбщая сумма: %s\n-------' % (
                            operator["name"], search.find('GroupName').get('FlgWorked'), search.find('GroupName').get('FlgSold'),
                            search.find('GroupName').get('FlgLoan'), search.find('GroupName').get('LoanAmount')
                        ))

                        LOG_OPERATOR = LOG_OPERATOR + 1
                    except AttributeError : print('Сотрудник: %s не найден в отчёте!\n-------' % (operator["name"]))
                elif report == "dash":
                    try:
                        for search in data:
                            if operator["name"] == search["Оператор"]:
                                print('Сотрудник: %s\nКол-во догов: %s\nОбщая сумма за день: %s\n-------' % (
                                    operator["name"], search["Количество"], search["Сумма"]
                                ))
                                LOG_OPERATOR = LOG_OPERATOR + 1
                            else : pass
                    except AttributeError : print('Сотрудник: %s не найден в отчёте!\n-------' % (operator["name"]))
            print('Кол-во сотрудников в отчёте: %s' % (LOG_OPERATOR))
            del get_data, data, group, file_groups, data_group, search, operator
        elif report == "agreements":
            for operator in group:
                search = data.find_all("Details3", attrs={"Сотрудник_выдачи2": operator["name"]})
                for info in search:
                    try:
                        #free_money = "ФЗ: Есть" if int(info.get("Отказ_от_ФЗ")) > 0 else "ФЗ: Отказ"
                        free_money = info.get("Тип_ФЗ2")
                        print("Сотрудник: %s --- Сумма: %s --- %s --- Клиент: %s --- ФЗ: %s --- ТП: %s" % (
                            operator["name"], info.get("Сумма_кредита2"), free_money, info.get("ФИО_Клиента3"), 
                            info.get("comis_insur_sum2"), info.get("comis_gold_card2")
                        ))
                        train = (operator["name"], info.get("Сумма_кредита2"), free_money, info.get("ФИО_Клиента3"), info.get("comis_insur_sum2"), info.get("comis_gold_card2"))
                        print(train)
                        RAM_THIRD_SLOT.append(train)
                        LOG_OPERATOR = LOG_OPERATOR + 1
                    except AttributeError : print('Сотрудник: %s не найден в отчёте!\n-------' % (operator["name"]))
            return True
    # Расчёт, загружает нужные файлы автоматом
    def calc(self, method="bonus" or "CR") -> True | False:
        global LOG_OPERATOR
        global LOG_DATA
        global LOG_BONUS
        global LOG_CR
        global RAM
        global RAM_SECOND_SLOT
        get_group = self.username if self.username != "homer" else "volkovva"
        match method:
            case "CR":
                report_98 = self._open_report("98")
                with open("groups.json", "r", encoding="utf-8") as file_groups : data_group = json.load(file_groups)
                group = data_group["%s" % (get_group)]
                data = bs4.BeautifulSoup(report_98, 'xml')
                for operator in group:
                    search = data.find('AuthorName', {'AuthorName': operator["name"]})
                    try:
                        name = operator["name"]
                        try : appl = int(search.find('GroupName').get('FlgWorked'))
                        except TypeError : appl = 0
                        try : OK = int(search.find('GroupName').get('FlgSold'))
                        except TypeError : OK = 0
                        try : signed = int(search.find('GroupName').get('FlgLoan'))
                        except : signed = 0
                        try : CR_OK = OK/appl*100
                        except ZeroDivisionError : CR_OK = 0
                        try : CR_SIGN = signed/OK*100
                        except ZeroDivisionError : CR_SIGN = 0 
                        _operator = {
                            "name": name,
                            "appl": appl,
                            "OK": OK,
                            "CR OK": str(CR_OK)[:5]+"%",
                            "signed": signed,
                            "CR SIGN": str(CR_SIGN)[:5]+"%"
                        }
                        print(_operator)
                        LOG_CR.append(_operator)
                        train = (_operator["name"], _operator["appl"], _operator["OK"], _operator["signed"], _operator["CR OK"], _operator["CR SIGN"])
                        RAM.append(train)
                        LOG_OPERATOR = LOG_OPERATOR + 1
                    except AttributeError : print('Сотрудник: %s не найден в отчёте!\n-------' % (operator["name"]))
                print('Кол-во сотрудников в отчёте: %s' % (LOG_OPERATOR))
                return True
            case "bonus":
                report_bonus = self._open_report("agreements")
                data = bs4.BeautifulSoup(report_bonus, 'xml')
                with open("groups.json", "r", encoding="utf-8") as file_groups : data_group = json.load(file_groups)
                group = data_group["%s" % (get_group)]
                for operator in group:
                    search = data.find_all("Details3", attrs={"Сотрудник_выдачи2": operator["name"]})
                    for info in search:
                        try:
                            LOG_DATA[0]
                        except IndexError:
                            LOG_DATA.append(operator["name"])
                        try:
                            if operator["name"] == LOG_DATA[0]:
                                LOG_BONUS = LOG_BONUS + super().FP[info.get("Тип_ФЗ2")] + super().TP[str(info.get("comis_gold_card2"))]
                            else:
                                print("Премия сотрудник\n\tСотрудник: %s\n\tПремия: %s $" % (LOG_DATA[0], LOG_BONUS))
                                train = (LOG_DATA[0], LOG_BONUS)
                                RAM_SECOND_SLOT.append(train)
                                LOG_DATA.clear()
                                LOG_BONUS = 0
                        except AttributeError : pass
                return True

class GUIData():
    def _labels(self, binding):
        pass
    def _buttons(self, binding):
        global LOG_DATA
        global LOG_GUI
        def _command():
            if TruAR("test").calc("CR") == True:
                if TruAR("test").calc("bonus") == True:
                    if TruAR("test").get_datas("test", "agreements") == True:
                        print(RAM, "\n", RAM_SECOND_SLOT)
                        txt = Label(binding, text="Конверсия")
                        txt.place(relx=0.2, rely=0.02)
                        txt2 = Label(binding, text="Премия за месяц")
                        txt2.place(relx=0.7, rely=0.02)
                        txt3 = Label(binding, text="Договора")
                        txt3.place(relx=0.2, rely=0.45)
                        treeCR = RAM
                        columns = ("Сотрудник", "Обработано", "Одобрено", "Подписано", "CR Одобрения", "CR Подписания")
                        tree = ttk.Treeview(binding, columns=columns, show="headings")
                        tree.heading("Сотрудник", text="Сотрудник")
                        tree.heading("Обработано", text="Обработано")
                        tree.heading("Одобрено", text="Одобрено")
                        tree.heading("Подписано", text="Подписано")
                        tree.heading("CR Одобрения", text="CR Одобрения")
                        tree.heading("CR Подписания", text="CR Подписания")
                        tree.column("#1", stretch=NO, width=250)
                        tree.column("#2", stretch=NO, width=100)
                        tree.column("#3", stretch=NO, width=100)
                        tree.column("#4", stretch=NO, width=100)
                        tree.column("#5", stretch=NO, width=100)
                        tree.column("#6", stretch=NO, width=100)
                        print(RAM)
                        print(len(RAM))
                        for item in range(0, len(treeCR)):
                            tree.insert("", END, values=treeCR[item])
                        tree.place(relx=0.005, rely=0.04)
                        treeBonus = RAM_SECOND_SLOT
                        columns_b = ("Сотрудник", "Премия")
                        treeB = ttk.Treeview(binding, columns=columns_b, show="headings")
                        treeB.heading("Сотрудник", text="Сотрудник")
                        treeB.heading("Премия", text="Премия")
                        treeB.column("#1", stretch=NO, width=250)
                        treeB.column("#2", stretch=NO, width=75)
                        print(RAM_SECOND_SLOT)
                        print(len(RAM_SECOND_SLOT))
                        for item_b in range(0, len(treeBonus)):
                            treeB.insert("", END, values=treeBonus[item_b])
                        treeB.place(relx=0.635, rely=0.04)

                        treeAgree = RAM_THIRD_SLOT
                        columns_A = ("Сотрудник", "Сумма кредита", "Тип ФЗ", "Клиент", "ФЗ", "ТП")
                        treeA = ttk.Treeview(binding, columns=columns_A, show="headings")
                        treeA.heading("Сотрудник", text="Сотрудник")
                        treeA.heading("Сумма кредита", text="Сумма кредита")
                        treeA.heading("Тип ФЗ", text="Тип ФЗ")
                        treeA.heading("Клиент", text="Клиент")
                        treeA.heading("ФЗ", text="ФЗ")
                        treeA.heading("ТП", text="ТП")
                        treeA.column("#1", stretch=NO, width=200)
                        treeA.column("#2", stretch=NO, width=200)
                        treeA.column("#3", stretch=NO, width=200)
                        treeA.column("#4", stretch=NO, width=200)
                        treeA.column("#5", stretch=NO, width=200)
                        treeA.column("#6", stretch=NO, width=200)
                        print(RAM_THIRD_SLOT)
                        print(len(RAM_THIRD_SLOT))
                        for item_a in range(0, len(treeAgree)):
                            treeA.insert("", END, values=treeAgree[item_a])
                        treeA.place(relx=0.005, rely=0.5)
            """placement = 0.004
            for item in range(0, len(RAM)):
                #print(RAM[item])
                #print(RAM[])
                info = Label(binding, text=RAM[item])
                info.place(relx=0.02, rely=placement)
                placement = placement + 0.0225"""

        PayLoadReports = ttk.Button(binding, text="Загрузить", command=_command)
        PayLoadReports.place(relx=0.75, rely=0.75)

class TruARGUI(GUIData):
    __main__ = Tk()
    def start(self):
        windll.shcore.SetProcessDpiAwareness(1)
        self.__main__.title("TruAR Beta")
        #self.__main__.geometry("1700x900")
        self.__main__.geometry("1250x900")
        self.__main__.resizable(False, False)
        super()._buttons(self.__main__)
        MainMenu = Menu(self.__main__)
        self.__main__.config(menu=MainMenu)
        MainMenu.add_command(label="Справка")
        self.__main__.mainloop()

if __name__ == "__main__":
    TruARGUI().start()
