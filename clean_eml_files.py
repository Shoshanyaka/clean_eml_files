import os
import sys
import pandas as pd
    
    
def open_and_clean(file_type):

    user_profile_path = os.environ['USERPROFILE']
    docs_code_path = '\\!Worker\\'
    scr_path = os.path.join(user_profile_path, docs_code_path)
    if not os.path.exists(scr_path):
        os.mkdir(scr_path)
    print('Сейчас откроется папка, куда необходимо положить файл для обработки...\n'
          'При копировании файла, скопируй в буфер обмена имя файла\n'
          'После того как переместишь файл в необходимую папку, вернись в эту консоль')
    input('Для продолжения нажми любую клавишу...')
    os.system(f'explorer {scr_path}')
    print('Для того, чтобы не париться с переписыванием имени файла и не удивляться тому, что тут не работает Ctrl+V\n'
          'Нужно нажимать Shift+Insert(на некоторых клавиатурах Ins)')
    file_name = input('Введите имя файла: ')
    if file_type == '1':
        data = pd.read_csv(f"{scr_path}\\{file_name}.csv", sep=';', low_memory=False) 
    else:
        data = pd.read_excel(f"{scr_path}\\{file_name}.xlsx", dtype=None)

    if ((not 'Рабочий e-mail') or (not 'Контакт: Рабочий e-mail') or (not 'e-mail') or (not 'email')) in data.columns:
        print('File incorrect')
        input('Для выхода нажмите любую клавишу...')
        sys.exit()

    if 'Рабочий e-mail' in data.columns:
        private_e_mail = data['Рабочий e-mail'].dropna()
        if 'Частный e-mail' not in data.columns:
            work_e_mail = data['Рабочий e-mail'].dropna()
        else:
            work_e_mail = data['Частный e-mail'].dropna()
        if 'E-mail для рассылок' not in data.columns:
            e_mail_for_mailing = data['Рабочий e-mail'].dropna()
        else:
            e_mail_for_mailing = data['E-mail для рассылок'].dropna()
        if 'Другой e-mail' not in data.columns:
            other_e_mail = data['Рабочий e-mail'].dropna()
        else:
            other_e_mail = data['Другой e-mail'].dropna()  

    if 'Контакт: Рабочий e-mail' in data.columns:
        private_e_mail = data['Контакт: Рабочий e-mail'].dropna()
        if 'Контакт: Частный e-mail' not in data.columns:
            work_e_mail = data['Контакт: Рабочий e-mail'].dropna()
        else:
            work_e_mail = data['Контакт: Частный e-mail'].dropna()
        if 'Контакт: E-mail для рассылок' not in data.columns:
            e_mail_for_mailing = data['Контакт: Рабочий e-mail'].dropna()
        else:
            e_mail_for_mailing = data['Контакт: E-mail для рассылок'].dropna()
        if 'Контакт: Другой e-mail' not in data.columns:
            other_e_mail = data['Контакт: Рабочий e-mail'].dropna()
        else:
            other_e_mail = data['Контакт: Другой e-mail'].dropna()

    if 'e-mail' in data.columns:
        private_e_mail = data['e-mail'].dropna()
        if 'Контакт: Частный e-mail' not in data.columns:
            work_e_mail = data['e-mail'].dropna()
        else:
            work_e_mail = data['e-mail'].dropna()
        if 'Контакт: E-mail для рассылок' not in data.columns:
            e_mail_for_mailing = data['e-mail'].dropna()
        else:
            e_mail_for_mailing = data['Контакт: E-mail для рассылок'].dropna()
        if 'Контакт: Другой e-mail' not in data.columns:
            other_e_mail = data['e-mail'].dropna()
        else:
            other_e_mail = data['e-mail'].dropna()   

    if 'email' in data.columns:
        private_e_mail = data['email'].dropna()
        if 'Контакт: Частный e-mail' not in data.columns:
            work_e_mail = data['email'].dropna()
        else:
            work_e_mail = data['email'].dropna()
        if 'Контакт: E-mail для рассылок' not in data.columns:
            e_mail_for_mailing = data['email'].dropna()
        else:
            e_mail_for_mailing = data['Контакт: E-mail для рассылок'].dropna()
        if 'Контакт: Другой e-mail' not in data.columns:
            other_e_mail = data['email'].dropna()
        else:
            other_e_mail = data['email'].dropna() 

    all_email = pd.concat( 
                    [private_e_mail, work_e_mail, 
                    e_mail_for_mailing, 
                    other_e_mail],
                    ignore_index=True, sort=False
                    ).drop_duplicates() 
    print(all_email)
    result = []

    for email in (all_email.tolist()): 
        if 'wazzup' in email:  
            continue
        elif ',' in email: 
            a = email.split(",")
            for i in a:
                result.append(i)
        else:           
            result.append(email)            

    final_lisе = pd.DataFrame(result).drop_duplicates()  
    writer = pd.ExcelWriter(f'{scr_path}\\eml_{file_name}.xlsx') 
    final_lisе.to_excel(writer, index=False) 
    writer.close()
    

if __name__ == "__main__":
    print('Добро пожаловать в скрипт автоматической чистки .CSV и .XLSX файлов\n'
    'Содержащие в себе e-mail адреса\n'
    'На вход принимаются файлы CSV формата с разделителем ";"\n'
    'Файл должен находиться в той же директории, что и исполняемый скрипт\n'
    'В данной версии принимаются файлы экспорта в первой строке которых\n'
    'Содержатся пометки вида:\n'
    'Рабочий e-mail\n'
    'Контакт: Рабочий e-mail\n'
    'e-mail\n'
    'email\n'
    )
    print('Выберите тип файла: *нажимая 1 или 2, Марина, 1 или 2*\n'
          '1. csv\n'
          '2. xlsx\n')
    file_type = input()
    if ((file_type == '1') or (file_type == '2')):
        open_and_clean(file_type)
    else:
        print('Введён неверный ключ')
        input('Для выхода нажмите любую клавишу...')