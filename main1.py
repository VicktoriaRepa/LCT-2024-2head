from tkinter import Tk, ttk
import tkinter as tk

#Вкладка СотрудникиНовая БД
from tab_d import Employees as TabD
#Вкладка соответствия сотрудников и типов исследований
from tab_e import Employees_study as TabE
#Вкладка соответствия сотрудников и типов исследований доп.модальностей
from tab_e1 import Employees_study_dop as TabED
#Вкладка с больничными и отпускными
from tab_f import NotWorking as TabF
#Вкладка с выгрузкой графика
from ex1 import Excel as Ex


def clear():
    # Удаление предыдущего окна, чтобы окна не наслаивались одно на другое
    if frm.winfo_children():
        frm.winfo_children()[0].destroy()
def Employees_vk():
    u"""Построение кнопки первого меню"""
    clear()
    if __name__ == '__main__':
        application = TabD(frm)
def Employees_study_vk():
    u"""Построение кнопки второго меню"""
    clear()
    if __name__ == '__main__':
        application = TabE(frm)
def Employees_study_dop_vk():
    u"""Построение кнопки второго меню"""
    clear()
    if __name__ == '__main__':
        application = TabED(frm)
def NotWorking_vk():
    u"""Построение кнопки второго меню"""
    clear()
    if __name__ == '__main__':
        application = TabF(frm)
def Excel_vk():
    u"""Построение кнопки второго меню"""
    clear()
    if __name__ == '__main__':
        application = Ex(frm)


root = tk.Tk()
root.geometry("700x700+100+100")
MB = tk.Menu(root)
MN = tk.Menu(MB)
MN.add_command(label=u"Сотрудники", command=Employees_vk)
MN.add_command(label=u"Типы исследований", command=Employees_study_vk)
MN.add_command(label=u"Типы исследований. Доп. модальности", command=Employees_study_dop_vk)
MN.add_command(label=u"Отпуска и больничные", command=NotWorking_vk)
MN.add_command(label=u"Расчет графика", command=Excel_vk)
MB.add_cascade(label=u"Меню", menu=MN)
root.config(menu=MB)
frm = tk.Frame(root, width=700, height=650, bg="blue")
frm.grid()
root.focus_force()
root.mainloop()