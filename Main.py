import json, os, requests, bs4, getpass, subprocess, time

LOG_GUI = []
LOG_OPERATOR = 0
LOG_DATA = []
LOG_CR = []
LOG_BONUS = 0

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
    username = getpass.getuser()
    #username = "volkovva"
    def __init__(self):
        self.calc("volkovva")
        #self.calc(self.username)
        #self.get_datas("volkovva", "agreements")
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
    def get_datas(self, get_group, report):
        global LOG_OPERATOR
        global LOG_DATA
        get_data = self._open_report(report)
        if report == "98" or report == "agreements" : data = bs4.BeautifulSoup(get_data, 'xml')
        elif report == "dash" : data = get_data
        with open("groups.json", "r", encoding="utf-8") as file_groups : data_group = json.load(file_groups)
        group = data_group["%s" % (get_group)]
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
            elif report == "agreements":
                search = data.find_all("Details3", attrs={"Сотрудник_выдачи2": operator["name"]})
                for info in search:
                    try:
                        #free_money = "ФЗ: Есть" if int(info.get("Отказ_от_ФЗ")) > 0 else "ФЗ: Отказ"
                        free_money = info.get("Тип_ФЗ2")
                        print("Сотрудник: %s --- Сумма: %s --- %s --- Клиент: %s --- ФЗ: %s --- ТП: %s" % (
                            operator["name"], info.get("Сумма_кредита2"), free_money, info.get("ФИО_Клиента3"), 
                            info.get("comis_insur_sum2"), info.get("comis_gold_card2")
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
    def calc(self, get_group):
        global LOG_OPERATOR
        global LOG_DATA
        global LOG_BONUS
        global LOG_CR
        report_98 = self._open_report("98")
        data = bs4.BeautifulSoup(report_98, 'xml')
        with open("groups.json", "r", encoding="utf-8") as file_groups : data_group = json.load(file_groups)
        group = data_group["%s" % (get_group)]
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
                LOG_OPERATOR = LOG_OPERATOR + 1
            except AttributeError : print('Сотрудник: %s не найден в отчёте!\n-------' % (operator["name"]))
        print('Кол-во сотрудников в отчёте: %s' % (LOG_OPERATOR))
        time.sleep(2)
        del report_98, data, file_groups, data_group, group
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
                        print("Премия сотрудник\n\tСотрудник: %s\n\tПремия: %s" % (LOG_DATA[0], LOG_BONUS))
                        LOG_DATA.clear()
                        LOG_BONUS = 0
                except AttributeError : pass

if __name__ == "__main__":
    TruAR()
