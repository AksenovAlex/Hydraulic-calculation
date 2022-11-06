import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog as fd
import  pandas as pd
from string import *

parameters = {
    1: 'Расход жидкости, куб.м/ч',
    2: 'Внутренний диаметр трубы, мм',
    3: 'Скорость потока, м/с'
            }

material_speed = {
    'Сталь, чугун': 2.5,
    'Медь': 0.9,
    'Алюминиевая латунь': 2.0,
    'МН сплав CuNi5Fe': 2.0,
    'МН сплав CuNi10Fe': 2.5,
    'МН сплав CuNi30Fe': 3.5,
    'Титановые сплавы': 10.0
                }
material = []
for key in material_speed:
    material.append(key)


def select_check():
    ch = check.get()
    if ch == 1:
        flow['state'] = tk.DISABLED
        diameter['state'] = tk.NORMAL
        rate['state'] = tk.NORMAL
    if ch == 2:
        diameter['state'] = tk.DISABLED
        flow['state'] = tk.NORMAL
        rate['state'] = tk.NORMAL
    if ch == 3:
        rate['state'] = tk.DISABLED
        flow['state'] = tk.NORMAL
        diameter['state'] = tk.NORMAL

def state_parameters(parameter, value):
    parameter.config(state=tk.NORMAL)
    parameter.insert(0, eval(value))
    parameter.config(disabledforeground='black')
    parameter.config(state=tk.DISABLED)

def clear_all_parameters(parameter):
    flow.get()
    diameter.get()
    rate.get()
    ch = check.get()
    if ch == 1:
        flow.config(state=tk.NORMAL)
        flow.delete(0, tk.END)
        flow.config(state=tk.DISABLED)
        diameter.delete(0, tk.END)
        rate.delete(0, tk.END)
    if ch == 2:
        diameter.config(state=tk.NORMAL)
        diameter.delete(0, tk.END)
        diameter.config(state=tk.DISABLED)
        flow.delete(0, tk.END)
        rate.delete(0, tk.END)
    if ch == 3:
        rate.config(state=tk.NORMAL)
        rate.delete(0, tk.END)
        rate.config(state=tk.DISABLED)
        flow.delete(0, tk.END)
        diameter.delete(0, tk.END)

def except_error(div):
    try:
        if float(div) == 0:
            raise  ZeroDivisionError
    except ZeroDivisionError:
        messagebox.showerror('Ошибка!', 'Параметры не должны быть равны 0')

def calculate():
    fl = flow.get()
    d = diameter.get()
    r = rate.get()
    try:
        if fl.isalpha() or d.isalpha() or r.isalpha():
            raise ValueError
    except ValueError:
        messagebox.showerror('Ошибка!', 'Некорректные значения!')

    ch = check.get()
    material_tube = speed_check.get()
    if ch == 1:
        fl = f'{r}*({d}**2)/354'
        except_error(d)
        state_parameters(flow, fl)
    elif ch == 2:
        d = f'(354*{fl}/{r})**0.5'
        except_error(r)
        state_parameters(diameter, d)
    elif ch == 3:
        r = f'354*{fl}/({d}**2)'
        except_error(d)
        state_parameters(rate, r)
        if int(eval(r)) > material_speed[material_tube]:
            messagebox.showinfo('Внимание!', f'Скорость более {material_speed[material_tube]} м/с')


def clear_values():
    clear_all_parameters(flow)
    clear_all_parameters(diameter)
    clear_all_parameters(rate)
    speed_check.current(0)

def add_menu():
    menubar = tk.Menu(win)
    win.config(menu=menubar)
    settings_menu = tk.Menu(menubar, tearoff=0)
    #settings_menu.add_command(label='Настройки')
    settings_menu.add_command(label='Экспорт', command=export)
    settings_menu.add_command(label='Выход', command=close_window)
    menubar.add_cascade(label='Файл', menu=settings_menu)

def close_window():
    answer = messagebox.askyesno(title='Выход', message='Закрыть программу')
    if answer:
        win.destroy()

def export():
    fl = flow.get()
    d = diameter.get()
    r = rate.get()
    sch = speed_check.get()

    df = pd.DataFrame({
                'Параметр':['Расход жидкости', 'Диаметр трубы', 'Скорость жидкости', 'Допустимая скорость'],
                'Значение':[fl, d, r, material_speed[sch]],
                'Размерность':['куб.м/ч', 'мм', 'м/с', 'м/с']
    })

    #df = {
    #    'Параметр': ['Расход жидкости', 'Диаметр трубы', 'Скорость жидкости', 'Допустимая скорость'],
    #    'Значение': [fl, d, r, sch],
    #    'Размерность': ['куб.м/ч', 'мм', 'м/с', 'м/с']
    #}

    #file_name = fd.asksaveasfilename(title='Сохранить файл', defaultextension='.xlsx')
    #if file_name:
    #    file_name.write(df)
    #    file_name.close()
    df.to_excel(r'Result.xlsx')


win = tk.Tk()
win.title("Гидравлический расчет")
win.geometry('400x200+100+100')
win['bg'] = 'aquamarine'

check = tk.IntVar()

tk.Label(win, text='Материал трубопровода', font=('Arial', 10), anchor='w', bg='aquamarine', padx=23).grid(row=3, column=0,
                                                                                                     stick='we')

flow = tk.Entry(win, justify=tk.RIGHT, font=('Arial', 10), bd=2, width=5)
flow.grid(row=0, column=1, stick='we')

diameter = tk.Entry(win, justify=tk.RIGHT, font=('Arial', 10), bd=2, width=5)
diameter.grid(row=1, column=1, stick='we')

rate = tk.Entry(win, justify=tk.RIGHT, font=('Arial', 10), bd=2, width=5)
rate.grid(row=2, column=1, stick='we')

tk.Radiobutton(win, text='Расход жидкости, куб.м/ч', font=('Arial', 10), bg='aquamarine', anchor='w', variable=check,
               value=1, command=select_check).grid(row=0, column=0, stick='w')
tk.Radiobutton(win, text='Внутренний диаметр трубы, мм', font=('Arial', 10), bg='aquamarine', anchor='w',
               variable=check, value=2, command=select_check).grid(row=1, column=0, stick='w')
tk.Radiobutton(win, text='Скорость потока, м/с', font=('Arial', 10), bg='aquamarine', anchor='w', variable=check,
               value=3, command=select_check).grid(row=2, column=0, stick='w')

speed_check = ttk.Combobox(win, values=material, justify=tk.LEFT, font=('Arial', 10))
speed_check.grid(row=3, column=1, stick='we', padx=5, pady=5)
speed_check.current(0)

calculate_button = tk.Button(win, text='Рассчитать', font=('Arial', 10), width=1, height=1, command=calculate).grid(row=4, column=1,
                                                                                                 stick='we')
clear_button = tk.Button(win, text='Сброс', font=('Arial', 10), width=1, height=1, command=clear_values).grid(row=4, column=0,
                                                                                                 stick='we', padx=5)

win.grid_columnconfigure(0, minsize=150)
win.grid_columnconfigure(1, minsize=175)

add_menu()

win.mainloop()
