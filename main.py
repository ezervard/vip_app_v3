from kivy.uix.boxlayout import BoxLayout
from kivy.core.window import Window
from kivymd.app import MDApp
from tkinter import Tk, filedialog, messagebox
from open_excel import Openfile
from find_same import Find
from create_raport import CreateRaport

Window.size = (300, 250)


def catch_db_except(func):#Декоратор отлова ошибки
    def inner(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except AttributeError:
            messagebox.showinfo('error', 'Нужно открыть Базу Данных')
    return inner
class Container(BoxLayout):
    def __init__(self, **kwargs):
        super(Container, self).__init__(**kwargs)
        self.orientation = 'vertical'
        self.display_text = ''
        self.raport = CreateRaport()
        self.row_index = 9

    def update_label(self): #Получение данных из текстового поля
          self.display_text = self.ids.text.text


    @catch_db_except
    def check(self): #Проверка наличия номеров в базе данных и отработка сценариев если нет
        search = Find()
        self.update_label()
        self.res = search.same_find(self.display_text, self.file_path)
        if self.res[7] == 'CC':
            self.res[7] = input('Введите Фамилию: ')
        elif self.res[7] == 'MB':
            self.res[7] = input('Точно вьезжает?\nЕсли да то введи фамилию!: ')


    @catch_db_except
    def entrance(self): #Функция отработки вьезда
        self.update_label()
        self.check()
        entr = True
        self.raport.get_raport(self.res, entr, self.row_index)


    @catch_db_except
    def exit(self): #Функция отработки выезда
        self.update_label()
        self.check()
        entr = False
        self.raport.get_raport(self.res, entr, self.row_index)
        self.row_index += 1
        print(self.res)


    def save(self):
        self.raport.save_workbook()

    def select_file(self):#Функция открытия файла Эксель
        exFile = Openfile()
        root = Tk()
        root.withdraw()
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*xlsx")])
        exFile.open(self.file_path)

class Combiner(MDApp):
    def __init__(self, **kwargs):
        self.title = 'Combiner v3.0'
        super().__init__(**kwargs)
    def build(self):
        return Container()

if __name__ == "__main__":
    Combiner().run()
