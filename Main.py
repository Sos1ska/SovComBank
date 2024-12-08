import json, os, requests, bs4, getpass, subprocess, time

log = []
LOG_GUI = []
LOG_OPERATOR = 0
LOG_DATA= []
LOG_DASH = []

class data:
    name_98 = "98_OperatorAggregates.xml"
    dash = "dash.json"

# Класс с основой системой подрузки данных + подсчёт
class TruAR(data):
    username = getpass.getuser()
    #username = "volkovva"
    def __init__(self):
        #self.calc("volkovva")
        #self.calc(self.username)
        self.get_datas("volkovva", "dash")
    def _open_report(self, report):
        def find(file, path) -> bool:
            try:
                _files = os.listdir(path)
                if file in _files : return True
                else : return False
            except FileNotFoundError:
                log.append('File Not Found')
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
    def get_datas(self, get_group, report):
        global LOG_OPERATOR
        global LOG_DATA
        get_data = self._open_report(report)
        if report == "98" : data = bs4.BeautifulSoup(get_data, 'xml')
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
                LOG_DATA.append(_operator)
                LOG_OPERATOR = LOG_OPERATOR + 1
            except AttributeError : print('Сотрудник: %s не найден в отчёте!\n-------' % (operator["name"]))
        print('Кол-во сотрудников в отчёте: %s' % (LOG_OPERATOR))


if __name__ == "__main__":
    TruAR()