#Отпуска, больничные
#Вкладка СотрудникиНовая БД
import tkinter as tk
import sqlite3 as sq
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
from sqlite3 import Error


def create_connection(db_file):
    conn = None
    try:
        conn = sq.connect(db_file)
        return conn
    except Error as e:
        print(e)

    return conn


def create_table(conn, create_table_sql):
    try:
        c = conn.cursor()
        c.execute(create_table_sql)
    except Error as e:
        print(e)


def main():
    database = r"vacation_shedule.db"
    # описание столбцов Сотрудников - id номер, слово и значение
    sql_create_not_working_table = """ CREATE TABLE IF NOT EXISTS not_working_days (
                                        id_employees integer NOT NULL,
                                        date_with text NOT NULL,
                                        date_to text NOT NULL
                                    ); """
    # подключение к базе
    conn = create_connection(database)

    # создание таблицы dictionary
    if conn is not None:
        create_table(conn, sql_create_employees_table)
    else:
        print("Ошибка: не удалось подключиться к базе.")


if __name__ == '__main__':
    main()


#------------------------
class NotWorking:
    db_name = 'vacation_shedule.db'

    def __init__(self, window):

        self.wind = window

        # создание элементов для ввода слов и значений
        frame = LabelFrame(self.wind, text='Введите данные сотрудника')
        frame.grid(row=0, column=0, columnspan=3, pady=20)
        Label(frame, text='ФИО: ').grid(row=1, column=0)

        con = sq.connect('vacation_shedule.db')
        cursor_em = con.cursor()
        # получаем все данные из таблицы employees1
        cursor_em.execute('SELECT name FROM employees1 ORDER BY name DESC')
        lis = cursor_em.fetchall()
        self.combobox = ttk.Combobox(frame, values=lis)
        self.combobox.grid(row=1, column=1)

        Label(frame, text='Дата начала отпуска/больничного: ').grid(row=2, column=0)
        self.date_with = Entry(frame)
        self.date_with.grid(row=2, column=1)
        Label(frame, text='Дата окончания отпуска/больничного: ').grid(row=3, column=0)
        self.date_to = Entry(frame)
        self.date_to.grid(row=3, column=1)

        ttk.Button(frame, text='Сохранить', command=self.add_not_working).grid(row=4, columnspan=3, sticky=W + E)
        self.message = Label(text='', fg='green')
        self.message.grid(row=4, column=0, columnspan=3, sticky=W + E)
        # таблица слов и значений
        columns = ("#1", "#2")
        self.tree = ttk.Treeview(frame, height=10, columns=columns)
        self.tree.grid(row=5, column=0, columnspan=3)
        self.tree.heading('#0', text='Сотрудник', anchor=CENTER)
        self.tree.heading('#1', text='Дата старта отпуска/больничного', anchor=CENTER)
        self.tree.heading('#2', text='Дата окончания отпуска/больничного', anchor=CENTER)

        # кнопки редактирования записей
        ttk.Button(frame, text='Удалить', command=self.delete_not_working).grid(row=6, column=0, sticky=W + E)
        ttk.Button(frame, text='Изменить', command=self.edit_word).grid(row=6, column=1, sticky=W + E)

        # заполнение таблицы
        self.get_employees()

    # подключение и запрос к базе
    def run_query(self, query, parameters=()):
        with sq.connect(self.db_name) as conn:
            cursor = conn.cursor()
            result = cursor.execute(query, parameters)
            conn.commit()
        return result

    # заполнение таблицы словами и их значениями
    def get_employees(self):
        records = self.tree.get_children()
        for element in records:
            self.tree.delete(element)
        query = 'select em.name, wd.date_with, wd.date_to from not_working_days wd join employees1 em on wd.id_employees = em.id'
        db_rows = self.run_query(query)
        for row in db_rows:
            self.tree.insert('', 0, text=row[0], values=row[1:])

    # валидация ввода
    def validation(self):
        return len(self.combobox.get()) != 0 and len(self.date_with.get()) != 0 and len(self.date_to.get()) != 0

    # добавление отпуска или больничного сотрудника
    def add_not_working(self):
        if self.validation():
            query = 'INSERT INTO not_working_days SELECT id AS id_employees, ?, ? from employees1 em where em.name = ?'
            parameters = (self.date_with.get(), self.date_to.get(), self.combobox.get())
            self.run_query(query, parameters)
            self.message['text'] = 'Сотрудник {} добавлен в штат'.format(self.combobox.get())
            self.combobox.delete(0, END)
            self.date_with.delete(0, END)
            self.date_to.delete(0, END)
        else:
            self.message['text'] = 'Введите ФИО и дату устройства на работу'
        self.get_employees()

    # удаление слова
    def delete_not_working(self):
        self.message['text'] = ''
        try:
            self.tree.item(self.tree.selection())['text'][0]
        except IndexError as e:
            self.message['text'] = 'Выберите сотрудника, которого нужно удалить'
            return
        self.message['text'] = ''
        sp = self.tree.item(self.tree.selection())
        print(sp)
        name = self.tree.item(self.tree.selection())['text']
        name1 = self.tree.item(self.tree.selection())['values']
        name2 = name1[0]
        print(name1)
        print(name2)
        query = 'DELETE FROM not_working_days where id_employees = (select em.id from employees1 em where name=?) and date_with = ?'
        self.run_query(query, (name,name2))
        self.message['text'] = 'Сотрудник {} успешно удален'.format(name)
        self.get_employees()

    # рeдактирование слова и/или значения
    def edit_word(self):
        self.message['text'] = ''
        try:
            self.tree.item(self.tree.selection())['values'][0]
        except IndexError as e:
            self.message['text'] = 'Выберите Сотрудника для изменения'
            return
        name = self.tree.item(self.tree.selection())['text']
        old_device_date = self.tree.item(self.tree.selection())['values'][0]
        self.edit_wind = Toplevel()
        self.edit_wind.title = 'Изменить сотрудника'

        Label(self.edit_wind, text='Прежнее слово:').grid(row=0, column=1)
        Entry(self.edit_wind, textvariable=StringVar(self.edit_wind, value=word), state='readonly').grid(row=0,
                                                                                                         column=2)

        Label(self.edit_wind, text='Новое слово:').grid(row=1, column=1)
        # предзаполнение поля
        new_word = Entry(self.edit_wind, textvariable=StringVar(self.edit_wind, value=word))
        new_word.grid(row=1, column=2)

        Label(self.edit_wind, text='Прежнее значение:').grid(row=2, column=1)
        Entry(self.edit_wind, textvariable=StringVar(self.edit_wind, value=old_meaning), state='readonly').grid(row=2,
                                                                                                                column=2)

        Label(self.edit_wind, text='Новое значение:').grid(row=3, column=1)
        # предзаполнение поля
        new_meaning = Entry(self.edit_wind, textvariable=StringVar(self.edit_wind, value=old_meaning))
        new_meaning.grid(row=3, column=2)

        Button(self.edit_wind, text='Изменить',
               command=lambda: self.edit_records(new_word.get(), word, new_meaning.get(), old_meaning)).grid(row=4,
                                                                                                             column=2,
                                                                                                             sticky=W)
        self.edit_wind.mainloop()

    # внесение изменений в базу
    def edit_records(self, new_word, word, new_meaning, old_meaning):
        query = 'UPDATE dictionary SET word = ?, meaning = ? WHERE word = ? AND meaning = ?'
        parameters = (new_word, new_meaning, word, old_meaning)
        self.run_query(query, parameters)
        self.edit_wind.destroy()
        self.message['text'] = 'слово {} успешно изменено'.format(word)
        self.get_words()


if __name__ == '__main__':
    window = Tk()
    application = NotWorking(window)
    window.mainloop()