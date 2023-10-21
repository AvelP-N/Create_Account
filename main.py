import sqlite3 as sq
import subprocess
import random
from time import sleep
from string import ascii_letters, punctuation, digits
from win32api import SetConsoleTitle
from colorama import init, Fore, Back
from xlrd import open_workbook
from re import findall
from os import getlogin
from transliterate import translit
from subprocess import run
from tabulate import tabulate


"""
CreateAccount v1.4
При создании случайного пароля для ЛИС систем в Портале не принимает символы ['<>], теперь заменяются цифрами

CreateAccount v1.3
Добавлена пауза в 3 секунды, после добавления доступов. Группа YandexUser не успевала прописаться в профиле пользователя

CreateAccount v1.2
В метод show_other_permissions добавлен вывод информации о дочернем запросе Битрикс 24

CreateAccount v1.1
В классе CreateAccount добавлен try/except где происходит подключение к Базе Данных

CreateAccount v1.0
Создание нового пользователя, добавление групп доступов, вывод заполненного шаблона с данными предоставленных доступов.
"""


class Console:
    """Вывод информации о пользователе и доступов в консоль"""

    @staticmethod
    def decoration_console():
        """Заголовок консоли и описание"""

        SetConsoleTitle('Create Account v1.4')
        print(f'{Fore.BLACK}{Back.YELLOW} *** Create new account *** ', '\n')

    @staticmethod
    def format_user_data(data_user):
        """Отформатировать данные о пользователе в табличном виде"""

        list_users = []
        headers_data = ('ФИО Сотрудника', 'Регион*', 'Должность', 'Компания', 'Подразделение', 'ФИО Руководителя',
                        'Рабочий телефон', 'Мобильный', 'e-mail')

        for object_user in data_user:
            list_users.append([object_user.user_name, object_user.region, object_user.job_title, object_user.company,
                               object_user.department, object_user.manager, object_user.office_phone,
                               object_user.mobile_phone, object_user.mail])

        return tabulate(list_users, headers=headers_data, tablefmt='pretty')

    @staticmethod
    def format_access_data(data_programs, data_folder_access, data_lis_access, data_access_1c, data_crm_access):
        """Отформатировать данные о доступах в табличном виде"""

        headers_access = ('Нестандартные программы', 'Доступы к папкам', 'Доступы в ЛИС', 'Доступы в 1С', 'CRM')
        access_list = [data_programs, data_folder_access, data_lis_access, data_access_1c, data_crm_access]
        print()

        max_len = max(len(access) for access in access_list)
        access_data = []

        for index in range(max_len):
            temp_list_access = []
            for access in access_list:
                if len(access) > index:
                    temp_list_access.append(access[index])
                else:
                    temp_list_access.append('-')
            access_data.append(temp_list_access)

        return tabulate(access_data, headers=headers_access, tablefmt='pretty', stralign='left')

    @staticmethod
    def get_user_account(full_name):
        """Получить информацию о пользователе с помощью команды в PowerShell"""
        try:
            ps_show_account = f'''net user {CreateAccount.fullname_and_login[full_name]} /domain'''
            show_account = run(['powershell.exe', '-Command', ps_show_account], capture_output=True, text=True,
                               encoding='866')
            if show_account.returncode == 0:
                return f'{Fore.GREEN}{show_account.stdout}'
            elif show_account.returncode == 1:
                return f'{Fore.RED}{show_account.stderr}'

        except subprocess.CalledProcessError as error:
            return f'{Fore.RED}Error {error}'

    @staticmethod
    def get_user_account_form(data):
        """Получить заполненную форму учётной записи пользователя"""

        form = f'Учетная запись в домене:\n{Fore.BLUE}{data.user_name}\n'
        form += f'{Fore.YELLOW}Логин: {Fore.BLUE}{CreateAccount.fullname_and_login[data.user_name]}\n'

        if str(data.mail).lower() == 'да':
            form += f'{Fore.YELLOW}Личный почтовый ящик: ' \
                    f'{Fore.BLUE}{CreateAccount.fullname_and_login[data.user_name]}@kdl.ru\n'

        form += f'{Fore.YELLOW}Пароль: {Fore.BLUE}Nemo@126\n'
        return form

    @staticmethod
    def show_user_access_lis_form(full_name, access_lis):
        """Вывести в консоль заполненную форму с доступами в Лапорт, Портал, УВ, Сервис-онлайн, TCLE"""

        lis = {'LIS - Сервисы Онлайн (Выезд на дом)': 'Сервис-онлайн', 'LIS - Портал (просмотр заявок)': 'Портал',
               'LIS - TCLE': 'TCLE', 'LIS - Laport': 'Лапорт', 'LIS - УВ (заведение заявок)': 'УВ'}
        name_lis_listing = [lis[inf_system] for inf_system in access_lis if inf_system in lis]

        print(f'Доступ в ', end='')
        print(*set(name_lis_listing), sep='\\', end=':\n')
        print(f'{Fore.BLUE}{full_name}')
        print(f'{Fore.YELLOW}Логин: {Fore.BLUE}{CreateAccount.fullname_and_login[full_name]}')
        password = "".join(random.choices(ascii_letters + digits + punctuation, k=8))
        password_edit = ''.join([random.choice(digits) if char in "'<>" else char for char in password])
        print(f'{Fore.YELLOW}Пароль: {Fore.BLUE}{password_edit}')

    @staticmethod
    def show_other_permissions():
        """Вывести в консоль заполненную форму для обращения с остальными доступами"""

        if data_file.folder_access:
            print(f'{Fore.YELLOW}Доступ к папкам предоставлен.', '\n')

        if data_file.access_1C:
            print(f'{Fore.YELLOW}Доступ в 1С будет предоставлен в дочернем запросе:', '\n')

        if data_file.programs:
            if 'Битрикс 24' in data_file.programs:
                print(f'{Fore.YELLOW}Доступ в "Битрикс 24" будет предоставлен в дочернем запросе:', '\n')

        if data_file.crm_access:
            print(f'{Fore.YELLOW}Доступ в CRM будет предоставлен в дочернем запросе:', '\n')

        if data_file.hardware:
            if data_file.users[0].region == 'Москва':
                print(f'{Fore.YELLOW}Заявка на подготовку оборудования направлена на вторую линию поддержки в качестве '
                      f'дочернего запроса:', '\n')
            else:
                print(f'{Fore.YELLOW}Для предоставления оборудования (телефон, гарнитура, принтер, сканер) обратитесь, '
                      f'пожалуйста, к руководству или IT специалисту (или аутсорс- компании) в Вашем филиале.')

        if 'QlikView' in data_file.programs:
            print(f'{Fore.YELLOW}Для предоставления доступа к системе QlikView необходимо обратиться к Малышкину '
                  f'Сергею. После предоставления доступа к QlikView сотрудник может обратиться к нам самоcтоятельно '
                  f'для настройки программы.')


class User:
    """Создаёт объект пользователя"""

    def __init__(self, user_name, city, job_title, company, department, manager, office_phone, mobile_phone, mail):
        self.user_name = str(user_name).strip()
        self.region = str(city).strip()
        self.job_title = str(job_title).strip()
        self.company = str(company).strip()
        self.department = str(department).strip()
        self.manager = str(manager).strip().title()
        self.office_phone = self.format_phone(office_phone)
        self.mobile_phone = self.format_phone(mobile_phone)
        self.mail = str(mail).strip()

    @staticmethod
    def format_phone(phone_number):
        """Проверка телефонного номера и форматирование"""

        if not phone_number:
            return ''

        if isinstance(phone_number, (int, float)):
            phone_number = str(int(phone_number))

        digit_list = findall(r'\d', phone_number)
        string_phone_number = ''.join(digit_list)

        if len(string_phone_number) == 7:  # Формат телефона (707)1234
            return f'({string_phone_number[:3]}){string_phone_number[3:]}'

        elif len(string_phone_number) == 10:  # Формат телефона +7 (123) 45-67-890
            return f'+7 ({string_phone_number[:3]}) {string_phone_number[3:6]}-' \
                   f'{string_phone_number[6:8]}-{string_phone_number[8:]}'

        return ''


class CreateAccount:
    """Создать аккаунт пользователю и добавить группы доступа"""

    login = None
    password = 'Nemo@126'
    fullname_and_login = {}

    try:
        with sq.connect(fr'C:\Users\{getlogin()}\PycharmProjects\CreateAccount\AccountDB.db') as connectDB:
            cursor = connectDB.cursor()
    except sq.OperationalError as error:
        print(f'{Fore.RED}{str(error).capitalize()}!')
        quit(input(f'{Fore.RED}Press Enter to exit the program!'))

    @staticmethod
    def check_domain_user(username):
        """Проверить логин в Active Directory """

        command_ad = f'powershell.exe Get-ADUser –Identity "{username}"'
        try:
            result = run(command_ad, capture_output=True, text=True, shell=True, encoding='866')
            if result.returncode == 0:
                print(f'{Fore.RED}{result.stdout.strip()}')
                print()
                print(f'{Fore.BLUE}{username} {Fore.RED}- this username already exists!')
                return 0
            elif result.returncode == 1:
                if 'Не удалось найти сервер' in result.stderr:
                    print(f'{Fore.RED}{result.stderr}')
                    quit(input(f'{Fore.RED}Press Enter to exit the program!'))
                elif 'Имя "Get-ADUser" не распознано' in result.stderr:
                    print(f'{Fore.RED}{result.stderr}')
                    quit(input(f'{Fore.RED}Press Enter to exit the program!'))
                else:
                    print(f'{Fore.BLUE}{username} {Fore.GREEN}- not found in Active Directory.', '\n')

        except subprocess.CalledProcessError as error:
            print(f'{Fore.RED}Error {error}')

    @classmethod
    def edit_login_yes_or_no(cls):
        """Пользовательский ввод. Редактировать логин или нет"""

        while True:
            change_login = input(f'{Fore.YELLOW}Do you want to change the login {Fore.BLUE}{cls.login}{Fore.YELLOW}? '
                                 f'Enter {Fore.GREEN}yes{Fore.YELLOW} or {Fore.GREEN}no{Fore.YELLOW}: ').strip()
            if change_login.lower() == 'yes':
                return 'yes'
            elif change_login.lower() == 'no':
                return 'no'
            elif change_login:
                print(f'{Fore.RED}Invalid input. Please enter "yes" or "no".\n')

    @classmethod
    def create_login(cls, full_name):
        """Создать логин для нового пользователя"""

        full_name_list = str(full_name).split()
        cls.login = str(translit(f'{full_name_list[1][0]}{full_name_list[0]}', reversed=True)).replace("'", "")

        while True:
            if cls.check_domain_user(cls.login) == 0:
                print()
                cls.login = input(f'{Fore.GREEN}Input new login: ')
                print(f'{Fore.GREEN}New login - {Fore.BLUE}{cls.login}', '\n')
            else:
                edit_login = cls.edit_login_yes_or_no()
                if edit_login == 'no':
                    print()
                    break

                print()
                cls.login = input(f'{Fore.GREEN}Input new login: ')
                print(f'{Fore.GREEN}New login - {Fore.BLUE}{cls.login}', '\n')

    @staticmethod
    def create_account_yes_or_no(login):
        """Пользовательский ввод. Создать учётную запись пользователю или нет"""

        while True:
            response = input(f'{Fore.YELLOW}Create a new account {Fore.BLUE}{login} {Fore.YELLOW}- '
                             f'{Fore.GREEN}yes {Fore.YELLOW}or {Fore.GREEN}no: ').strip().lower()
            if response == 'yes':
                return True
            elif response == 'no':
                return False
            elif response:
                print(f'{Fore.RED}Please enter "yes" or "no"!\n')

    def get_variables_powershell(self, user_name, path_region):
        """Получить переменные для создания пользователя в PowerShell"""

        given_name = str(translit(str(user_name).split()[1], reversed=True)).replace("'", "")
        surname = str(translit(str(user_name).split()[0], reversed=True)).replace("'", "")
        mail = f'{self.login}@kdl.ru'

        data_db = self.cursor.execute(f'SELECT full_name_region, path_powershell FROM path_regions '
                                      f'WHERE path_name LIKE "%{path_region}%"').fetchone()
        if data_db:
            region, path = data_db
            return given_name, surname, mail, region, path

        print(f'{Fore.RED}The database returned nothing!')
        print(f'Check the region - {Fore.BLUE}{path_region}')
        quit(input(f'{Fore.RED}Press Enter to exit the program...'))

    def create_user_account(self, data_user):
        """Создать учётную запись новому пользователю"""

        print(f'{Fore.BLUE}{data_user.user_name} {Fore.YELLOW}- checking a user in Active Directory.')

        # Создать логин для нового пользователя, который пришёл на вход
        self.create_login(data_user.user_name)
        self.fullname_and_login[data_user.user_name] = self.login

        given_name, surname, mail, region, path = self.get_variables_powershell(data_user.user_name, data_user.region)

        ps_create_ad_user = f'''New-ADUser -Name "{data_user.user_name}" -Path {path} `
                                -GivenName "{given_name}" -Surname "{surname}" -Title "{data_user.job_title}" `
                                -Department "{data_user.department}" -UserPrincipalName "{mail}" `
                                -DisplayName "{data_user.user_name}" -SamAccountName {self.login} -City "{region}" `
                                -AccountPassword (ConvertTo-SecureString "{self.password}" -AsPlainText -Force) `
                                -Enabled $true -Company '{data_user.company}' -OfficePhone "{data_user.office_phone}" `
                                -MobilePhone "{data_user.mobile_phone}" -State "{region}" `
                                -Manager (Get-ADUser -Filter  {{DisplayName -eq "{data_user.manager}"}} ).SamAccountName'''

        ps_show_user_account = f'Get-ADUser -Identity "{self.login}"'

        # Создание нового пользователя
        try:
            if self.create_account_yes_or_no(self.login):
                ad_user = run(['powershell.exe', '-Command', ps_create_ad_user], capture_output=True, text=True,
                              encoding='866')
                if ad_user.returncode == 0:
                    print(ad_user.stdout)
                    result = run(['powershell.exe', '-Command', ps_show_user_account], capture_output=True,
                                 text=True, encoding='866')
                    print(result.stdout)

                    return 'New account created'

                else:
                    print()
                    print(f'{Fore.RED}{ad_user.stderr}')
                    quit(input(f'{Fore.RED}Press Enter to exit the program!'))

        except subprocess.CalledProcessError as error:
            print(f'{Fore.RED}Error {error}')

    def add_mail_portal(self):
        """Создать почтовый ящик и дать доступ в Портал самообслуживания"""

        ps_add_mail = f'''Add-ADGroupMember -Identity 'YandexUsers' -Members {self.login} `
                            -server msk-dc-01.ad.kdl-test.ru; Set-ADUser {self.login} -email {self.login}@kdl.ru'''
        ps_add_portal = f'''Add-ADGroupMember -Identity 'ITSMPORTALUsers' -Members {self.login}'''
        access_dict = {'E-mail': (ps_add_mail, 'YandexUsers'), 'PORTAL Users': (ps_add_portal, 'ITSMPORTALUsers')}

        try:
            for key, values in access_dict.items():
                command = run(['powershell.exe', '-Command', values[0]], capture_output=True, text=True,
                              encoding='866')

                if command.returncode == 0:
                    print(f'{Fore.YELLOW}{key} {Fore.GREEN}add group - {Fore.BLUE}{values[1]}')
                elif command.returncode == 1:
                    print(f'{Fore.RED}{command.stderr}')

        except subprocess.CalledProcessError as error:
            print(f'{Fore.RED}Error {error}')

    def add_access_programs(self, programs):
        """Предоставление пользователю доступов к программам"""

        for program in programs:
            data_db = self.cursor.execute(f'SELECT ad_group FROM programs '
                                          f'WHERE name_program LIKE "%{program}%"').fetchone()
            if data_db:
                self.add_access(data_db[0], program)

    def add_access_folders(self, folders):
        """Добавление доступов к сетевым папкам"""

        for folder in folders:
            data_db = self.cursor.execute(f'SELECT ad_group FROM folders '
                                          f'WHERE name_folder == "{folder}"').fetchone()
            if data_db:
                self.add_access(data_db[0], folder)

    def add_access_1c(self, access_1c):
        """Предоставление доступов к 1С системам"""

        # Добавление группы RDP OS, т.к. требуется всегда
        self.add_access('RDP OC', 'RDP OC')

        # Основной цикл для добавления нужных групп
        for name_1c in access_1c:
            data_db = self.cursor.execute(f'SELECT ad_group FROM access_1C '
                                          f'WHERE name_1c == "{name_1c}"').fetchone()
            if data_db:
                self.add_access(data_db[0], name_1c)

    def add_access_crm(self, access_crm):
        """Предоставление доступа к CRM системе"""

        for region_crm in access_crm:
            data_db = self.cursor.execute(f'SELECT ad_group FROM crm '
                                          f'WHERE region == "{region_crm}"').fetchone()
            if data_db:
                self.add_access(data_db[0], region_crm)

    def add_access(self, item_db, item_label):
        try:
            ps_command = f'''Add-ADGroupMember -Identity '{item_db}' -Members {self.login}'''
            add_group = run(['powershell.exe', '-Command', ps_command], capture_output=True, text=True,
                            encoding='866')
            if add_group.returncode == 0:
                print(f'{Fore.YELLOW}{item_label} {Fore.GREEN}add group - {Fore.BLUE}{item_db}')
            elif add_group.returncode == 1:
                print(f'{Fore.RED}{add_group.stderr}')

        except subprocess.CalledProcessError as error:
            print(f'{Fore.RED}Error {error}')


class ParsFile(Console, CreateAccount):
    """Сбор данных из файла"""

    def __init__(self, path):
        self.path = path
        self.sheet = None
        self.users = []
        self.programs = []
        self.folder_access = []
        self.lis_access = []
        self.access_1C = []
        self.crm_access = []
        self.hardware = []
        self.start_row_user = None
        self.end_row_user = None
        self.start_row_access = None
        self.end_row_access = None

    def data_search_area(self):
        """Проверка области поиска для сбора данных пользователей и доступов"""

        for number_row in range(self.sheet.nrows):
            if 'ФИО' in self.sheet.cell_value(number_row, 1):
                self.start_row_user = number_row + 1
            elif 'Нестандартные' in self.sheet.cell_value(number_row, 1):
                self.end_row_user = number_row
                self.start_row_access = number_row + 1
            elif 'Дополнительные ваши' in self.sheet.cell_value(number_row, 1):
                self.end_row_access = number_row

    def parse_users(self):
        """Проверка пользователей в файле и сбор данных"""

        try:
            with open_workbook(self.path) as book:
                self.sheet = book.sheet_by_index(0)

                # Поиск диапазона строк в файле для сбора данных пользователей и доступов
                self.data_search_area()
                if not self.end_row_access:
                    quit(input(f'{Fore.RED}Wrong form to fill out!!!'))

                for number_row in range(self.start_row_user, self.end_row_user):
                    # print(self.sheet.row_values(number_row))
                    data = self.sheet.row_values(number_row, 1, 11)
                    cell_data = findall(r'[А-Я][а-яё]+', data[0])

                    if len(cell_data) == 3:
                        user_details = User(data[0], data[2], data[3], data[4], data[5], data[6], data[7], data[8],
                                            data[9])
                        self.users.append(user_details)

        except FileNotFoundError as file_not:
            print(f'{Fore.RED}{file_not}')
            quit(input(f'{Fore.RED}Press Enter for out of programs!'))

    def pars_access(self):
        """Проверка доступов для пользователя и сбор данных"""

        for number_row in range(self.start_row_access, self.end_row_access):
            find_program = str(self.sheet.cell_value(number_row, 1)).strip()
            find_folder = str(self.sheet.cell_value(number_row, 3)).strip()
            find_lis = str(self.sheet.cell_value(number_row, 5)).strip()
            find_1c = str(self.sheet.cell_value(number_row, 6)).strip()
            find_crm = str(self.sheet.cell_value(number_row, 8)).strip()
            find_hardware = str(self.sheet.cell_value(number_row, 9)).strip()

            if find_program:
                self.programs.append(find_program)
            if find_folder:
                self.folder_access.append(find_folder)
            if find_lis:
                self.lis_access.append(find_lis)
            if find_1c:
                self.access_1C.append(find_1c)
            if find_crm:
                self.crm_access.append(find_crm)
            if find_hardware:
                self.hardware.append(find_hardware)


if __name__ == '__main__':
    init(autoreset=True)
    Console.decoration_console()
    data_file = ParsFile(fr'C:\Users\{getlogin()}\Downloads\account.xls')
    data_file.parse_users()
    data_file.pars_access()

    print(data_file.format_user_data(data_file.users))
    print(data_file.format_access_data(data_file.programs, data_file.folder_access, data_file.lis_access,
                                       data_file.access_1C, data_file.crm_access))
    # print(data_file.__dict__)
    # print()

    # Создание учётной записи пользователя и добавление групп доступа
    if data_file.users:
        for user in data_file.users:
            account = data_file.create_user_account(user)
            print()

            if account == 'New account created':
                print(f'{Fore.BLUE}{user.user_name}')
                print(data_file.login, '\n')

                if str(user.mail).lower() == 'да':
                    print(f'{Fore.CYAN}Access to mail and self-service portal:')
                    data_file.add_mail_portal()
                    print()

                if data_file.programs:
                    print(f'{Fore.CYAN}Access to the program:')
                    data_file.add_access_programs(data_file.programs)
                    print()

                if data_file.folder_access:
                    print(f'{Fore.CYAN}Access to network folders:')
                    data_file.add_access_folders(data_file.folder_access)
                    print()

                if data_file.access_1C:
                    print(f'{Fore.CYAN}Access to 1C:')
                    data_file.add_access_1c(data_file.access_1C)
                    print()

                if data_file.crm_access:
                    print(f'{Fore.CYAN}Access to CRM:')
                    data_file.add_access_crm(data_file.crm_access)
                    print()

                sleep(3)
                print(data_file.get_user_account(user.user_name))

        # Вывести в консоль заполненную форму для обращения
        for user_object in data_file.users:
            print(data_file.get_user_account_form(user_object))
            print()

            if data_file.lis_access:
                data_file.show_user_access_lis_form(user_object.user_name, data_file.lis_access)
                print()

        data_file.show_other_permissions()
        print()

    else:
        print(f'{Fore.RED}No users available!\n')

    while True:
        end = input(f'{Fore.GREEN}Write {Fore.BLUE}end {Fore.GREEN}for exit of program: ')
        if end == 'end':
            break
