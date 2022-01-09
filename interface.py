from tkinter import *
from tkinter import messagebox as mb
from tkinter import filedialog
import  tkinter.filedialog
from proverka import *
class Job(Frame):
    def __init__(self):
        super().__init__()
        self.cr()
    def cr(self):
        self.fra1 = Frame(self.master, width=250, height=700, background='gainsboro')
        self.fra1.place(x=0, y=0)
        self.fra2 = Frame(self.master, width=800, height=700)
        self.fra2.place(x=250, y=0)
        but1_fra1 = Button(self.fra1, text='Прибавление дней к дате', border=3, font='Arial 13', command=self.operations1)
        but1_fra1.place(x=20, y=100)
        but2_fra1 = Button(self.fra1, text='Проверка количества дней\nотпуска на заданную дату', border=3, font='Arial 13', command=self.operations2)
        but2_fra1.place(x=15, y=200)
    def operations1(self):
        for widget in self.fra2.winfo_children():
            widget.destroy()
        self.labfra1 = LabelFrame(self.fra2, width=600, text="Дата", font="Arial 12", height=110)
        self.labfra1.place(x=50, y=50)
        self.lab1_labfra1 = Label(self.labfra1, text="Введите дату в формате ДД.ММ.ГГГГ", font="Arial 11")
        self.lab1_labfra1.place(x=10, y=10)
        self.ent1_labfra1 = Entry(self.labfra1, font="Arial 11", width=3, background="gainsboro")
        self.ent1_labfra1.place(x=30, y=50)
        self.lab2_labfra1 = Label(self.labfra1, font="Arial 12", text=".")
        self.lab2_labfra1.place(x=60, y=50)
        self.ent2_labfra1 = Entry(self.labfra1, font="Arial 11", width=3, background="gainsboro")
        self.ent2_labfra1.place(x=75, y=50)
        self.lab3_labfra1 = Label(self.labfra1, font="Arial 12", text=".")
        self.lab3_labfra1.place(x=105, y=50)
        self.ent3_labfra1 = Entry(self.labfra1, font="Arial 11", width=5, background="gainsboro")
        self.ent3_labfra1.place(x=120, y=50)

        self.labfra2 = LabelFrame(self.fra2, width=600, text="Дни", font="Arial 12", height=110)
        self.labfra2.place(x=50, y=200)
        self.lab1_labfra2 = Label(self.labfra2, text="Введите количество дней", font="Arial 11")
        self.lab1_labfra2.place(x=10, y=10)
        self.ent1_labfra2 = Entry(self.labfra2, font="Arial 11", width=5, background="gainsboro")
        self.ent1_labfra2.place(x=30, y=50)
        self.button1_op1 = Button(self.fra2, text="Подсчитать", font="Arial 12", border=3, command=self.result_operation1)
        self.button1_op1.place(x=550, y=350)
    def result_operation1(self):
        labfra3 = LabelFrame(self.fra2, width=600, text="Результат", font="Arial 12", height=110)
        labfra3.place(x=50, y=450)
        try:
            self.data1 = int(self.ent1_labfra1.get())
            self.data2 = int(self.ent2_labfra1.get())
            self.data3 = int(self.ent3_labfra1.get())
            self.chislo = int(self.ent1_labfra2.get())
            dat = datetime.date(self.data3, self.data2, self.data1)
            result_dat = dat + datetime.timedelta(days=self.chislo - 1)
            result_dat = result_dat.strftime("%d.%m.%Y")
            lab_labfra3 = Label(labfra3, text=f"{self.data1}.{self.data2}.{self.data3} + {self.chislo} дня(-ей) = {result_dat}",
                                font="Arial 13")
            lab_labfra3.place(x=20, y=20)
        except ValueError:
            mb.showerror("Ошибка",
                         "Возможно вы упустили какую-то ячейку при заполнении\nЛибо неправильно ввели формат даты")
    def operations2(self):
        for widget in self.fra2.winfo_children():
            widget.destroy()
        self.labfra1 = LabelFrame(self.fra2, width=600, text="Выберите файл", font="Arial 12", height=80)
        self.labfra1.place(x=50, y=20)
        self.lab1_1_labfra1=Label(self.labfra1, background="gainsboro", width=50, font="Arial 13")
        self.lab1_1_labfra1.place(x=10, y=12)
        self.button1_labfra1 = Button(self.labfra1, text="Выбрать", font="Arial 11", command=self.Open)
        self.button1_labfra1.place(x=500, y=10)

    def Open(self):
        ftypes = [('Excel файлы', '*.xlsx'), ('Все файлы', '*')]
        dlg = filedialog.Open(filetypes=ftypes)
        self.fl = dlg.show()
        self.lab1_1_labfra1["text"]=self.fl
        if self.fl != "":
            self.labfra2_2=LabelFrame(self.fra2, width=600, text="Впишите номер листа", font="Arial 12", height=60)
            self.labfra2_2.place(x=50, y=120)
            self.ent1_labfra2_2=Entry(self.labfra2_2, font="Arial 12", width=20, background="gainsboro")
            self.ent1_labfra2_2.place(x=10, y=10)
            self.button1_labfra2_2=Button(self.labfra2_2, text="Далее", font="Arial 11", border=3, command=self.vvod_operations2)
            self.button1_labfra2_2.place(x=500, y=0)

        else:
            mb.showerror("Ошибка",
                         "Вы не выбрали файл")
    def vvod_operations2(self):
        try:
            self.num=int(self.ent1_labfra2_2.get())-1
            self.labfra3_2 = LabelFrame(self.fra2, width=350, text="Выберите столбцы для ввода значений", font="Arial 12", height=270)
            self.labfra3_2.place(x=10, y=220)
            self.lab1_labfra3_2=Label(self.labfra3_2, text="Начальный ряд:", font="Arial 11")
            self.lab1_labfra3_2.place(x=10, y=10)
            self.ent1_labfra3_2=Entry(self.labfra3_2, background="gainsboro", width=6)
            self.ent1_labfra3_2.place(x=300, y=12)
            self.lab2_labfra3_2 = Label(self.labfra3_2, text="Конечный ряд:", font="Arial 11")
            self.lab2_labfra3_2.place(x=10, y=35)
            self.ent2_labfra3_2 = Entry(self.labfra3_2, background="gainsboro", width=6)
            self.ent2_labfra3_2.place(x=300, y=37)
            self.lab3_labfra3_2 = Label(self.labfra3_2, text="Столбец,где табельный номер:", font="Arial 11")
            self.lab3_labfra3_2.place(x=10, y=60)
            self.ent3_labfra3_2 = Entry(self.labfra3_2, background="gainsboro", width=6)
            self.ent3_labfra3_2.place(x=300, y=62)
            self.lab4_labfra3_2 = Label(self.labfra3_2, text="Столбец,где ФИО сотрудников:", font="Arial 11")
            self.lab4_labfra3_2.place(x=10, y=85)
            self.ent4_labfra3_2 = Entry(self.labfra3_2, background="gainsboro", width=6)
            self.ent4_labfra3_2.place(x=300, y=87)
            self.lab5_labfra3_2 = Label(self.labfra3_2, text="Столбец,где остаток дней отпуска:", font="Arial 11")
            self.lab5_labfra3_2.place(x=10, y=110)
            self.ent5_labfra3_2 = Entry(self.labfra3_2, background="gainsboro", width=6)
            self.ent5_labfra3_2.place(x=300, y=112)
            self.lab6_labfra3_2 = Label(self.labfra3_2, text="Столбец,где дата начала отпуска:", font="Arial 11")
            self.lab6_labfra3_2.place(x=10, y=135)
            self.ent6_labfra3_2 = Entry(self.labfra3_2, background="gainsboro", width=6)
            self.ent6_labfra3_2.place(x=300, y=137)
            self.lab7_labfra3_2 = Label(self.labfra3_2, text="Столбец,где количество дней на отпуск:", font="Arial 11")
            self.lab7_labfra3_2.place(x=10, y=160)
            self.ent7_labfra3_2 = Entry(self.labfra3_2, background="gainsboro", width=6)
            self.ent7_labfra3_2.place(x=300, y=162)
            self.lab8_labfra3_2 = Label(self.labfra3_2, text="Коэффициент:", font="Arial 11")
            self.lab8_labfra3_2.place(x=10, y=185)
            self.ent8_labfra3_2 = Entry(self.labfra3_2, background="gainsboro", width=6)
            self.ent8_labfra3_2.place(x=300, y=187)

            self.labfra4_2 = LabelFrame(self.fra2, width=350, text="Выберите столбцы для вывода значений",font="Arial 12", height=270)
            self.labfra4_2.place(x=410, y=220)
            self.lab1_labfra4_2 = Label(self.labfra4_2, text="Столбец, где полученный остаток:", font="Arial 11")
            self.lab1_labfra4_2.place(x=10, y=10)
            self.ent1_labfra4_2 = Entry(self.labfra4_2, background="gainsboro", width=6)
            self.ent1_labfra4_2.place(x=300, y=12)
            self.lab2_labfra4_2 = Label(self.labfra4_2, text="Столбец, где конечная дата:", font="Arial 11")
            self.lab2_labfra4_2.place(x=10, y=35)
            self.ent2_labfra4_2 = Entry(self.labfra4_2, background="gainsboro", width=6)
            self.ent2_labfra4_2.place(x=300, y=37)
            self.lab3_labfra4_2 = Label(self.labfra4_2, text="Столбец, где вывод:", font="Arial 11")
            self.lab3_labfra4_2.place(x=10, y=60)
            self.ent3_labfra4_2 = Entry(self.labfra4_2, background="gainsboro", width=6)
            self.ent3_labfra4_2.place(x=300, y=62)
            self.button_res=Button(self.fra2, text="Начать проверку", font="Arial 12",border=3, command=self.result_operations2)
            self.button_res.place(x=310, y=600)
            self.labfra5_2 = LabelFrame(self.fra2, width=600, text="Выберите название файла и его путь для сохранения",
                                        font="Arial 12", height=80)
            self.labfra5_2.place(x=10, y=500)
            self.lab_labfra5_2 = Label(self.labfra5_2, background="gainsboro", width=50, font="Arial 13")
            self.lab_labfra5_2.place(x=10, y=12)
            self.but_labfra5_2 = Button(self.labfra5_2, text="Обзор", font="Arial 11", command=self.save)
            self.but_labfra5_2.place(x=500, y=10)


        except ValueError:
            mb.showerror("Ошибка",
                         "Вы не ввели номер страницы с которой\nвы будете работать")
    def save(self):
        self.s = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx")
        self.lab_labfra5_2['text']=self.s
    def result_operations2(self):
        try:

            self.id=(self.ent3_labfra3_2.get()).upper()
            self.fio=(self.ent4_labfra3_2.get()).upper()
            self.start_ostatok=(self.ent5_labfra3_2.get()).upper()
            self.start_date=(self.ent6_labfra3_2.get()).upper()
            self.kol_dney=(self.ent7_labfra3_2.get()).upper()
            self.end_ostatok=(self.ent1_labfra4_2.get()).upper()
            self.end_date=(self.ent2_labfra4_2.get()).upper()
            self.conclusion=(self.ent3_labfra4_2.get()).upper()
            self.start_row = int(self.ent1_labfra3_2.get())
            self.end_row = int(self.ent2_labfra3_2.get())
            self.koef = float(self.ent8_labfra3_2.get())

            try:
                process(self.fl,self.num,self.id, self.fio,self.start_ostatok,self.start_date,self.kol_dney,self.end_ostatok,self.end_date,self.conclusion,self.start_row,self.end_row,self.koef, self.s)
                mb.showinfo("Результат",
                             "Проверка прошла успешна\nФайл сохранён в указанном пути")
            except AttributeError:
                mb.showerror("Ошибка",
                             "Возможно вы ввели неправильный столбец\nПерепроверьте значения столбцов")

        except ValueError:
            mb.showerror("Ошибка",
                         "Возможно вы упустили какую-то ячейку при заполнении\nЛибо ввели неправильный формат значения")



def main():
    root = Tk()
    Job()
    root.geometry("1050x700")
    root.title("My project")
    root.mainloop()


if __name__ == '__main__':
    main()




