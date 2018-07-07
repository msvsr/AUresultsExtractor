import warnings
import requests
import contextlib
from bs4 import BeautifulSoup
from openpyxl import Workbook

#Handling SSL Certification.
try:
    from functools import partialmethod
except ImportError:
    from functools import partial

    class partialmethod(partial):
        def __get__(self, instance, owner):
            if instance is None:
                return self

            return partial(self.func, instance, *(self.args or ()), **(self.keywords or {}))

@contextlib.contextmanager
def no_ssl_verification():
    old_request = requests.Session.request
    requests.Session.request = partialmethod(old_request, verify=False)

    warnings.filterwarnings('ignore', 'Unverified HTTPS request')
    yield
    warnings.resetwarnings()

    requests.Session.request = old_request


#Generates data.
def generate_data(number,flag):
    with no_ssl_verification():
        res = requests.post('https://aucoe.info/RDA/resultsnew/result_grade.php', data={'serialno': 1047,
                                                                                        'course': 'B.E./B.TECH/B.ARCH/INTEGRATED COURSE THIRD YEAR SECOND SEMESTER',
                                                                                        'degree': 'B.E./B.TECH/B.Arch/Integrated Course',
                                                                                        'table': 'gradestructure3',
                                                                                        'appearing_year': 'APRIL 2018',
                                                                                        'Date_time': '2018-07-05 16:50:19',
                                                                                        'regno': number,
                                                                                        'revdate': '2018-07-19',
                                                                                        'revfee': 750
                                                                                        })

        soup = BeautifulSoup(res.text,"lxml")
        data=[]
        for table in soup.find_all("table")[3:4]:
            trs = table.find_all("tr")
            data.append(trs[0].text.split(':')[flag].strip())
            data.append(trs[1].text.split(':')[flag].strip())

        for table in soup.find_all("table")[4:5]:
            trs = table.find_all("tr")[1:]
            for tr in trs:
                tds = tr.find_all('td')
                data.append(tds[flag].text)
        return data


#Gettings marks and storing in a file.
def generate_marks(starting_no,ending_no,file_name,wb):

    ws1 = wb.create_sheet(file_name)
    ws1.append(generate_data(starting_no,0))
    for i in range(starting_no,ending_no+1):
        ws1.append(generate_data(i,1))

    wb.save("3-2Results.xlsx")


if __name__=='__main__':
    wb = Workbook()

    generate_marks(315175710001,315175710288,'CSE',wb)
    generate_marks(315175711001, 315175711184, 'IT', wb)
    generate_marks(315175714001, 315175714283, 'EEE', wb)

    generate_marks(315175712001, 315175712288, 'ECE', wb)
    generate_marks(315175720001, 315175720358, 'MECH', wb)
    generate_marks(315175708001, 315175708218, 'CIVIL', wb)

