from open_excel import Openfile
import datetime
class Find:
    def same_find(self, finder, file_path):
        exFile = Openfile()

        exFile.open(file_path)
        sheet = exFile.get_worksheet()

        card_db = []
        correct_card = []
        plate_db = []
        plate_replace = []
        names_db = []
        places_db = []
        control_db = []

        for row in sheet.iter_rows():
            if row is not None:
                plates_data = row[3].value
                names_data = row[8].value
                places_data = row[5].value
                control_data = row[7].value
                card_data = row[0].value
                card_db.append(card_data)
                control_db.append(control_data)
                places_db.append(places_data)
                plate_db.append(plates_data)
                names_db.append(names_data)

        for i in card_db:
            if i == 'DEPARTMENT LEADER':
                correct_card.append('TOP MANAGEMENT')
            elif i == 'DIRECTOR':
                correct_card.append('TOP MANAGEMENT')
            else:
                correct_card.append(i)

        current_datetime = datetime.datetime.now()
        current_time = current_datetime.strftime("%H:%M")

        for row in sheet.iter_rows(): #Перебор всех значений в Базе Данных и сравнение с искомым значением
            plate = row[6].value

            if finder.upper() in plate_db:
                index = plate_db.index(finder.upper())
                name = names_db[index]
                if name == None:
                    name = 'CC'
                card = card_db[index]
                place = places_db[index]
                control = control_db[index]

                #Создание пакета данных для найденных машин
                pack = [card, 'OSOBOWY', 'LG', finder.upper(),place,current_time, control, name.upper() ]
                return pack
            elif finder.upper() not in plate_db:
                pack = [' ', 'OSOBOWY', 'LG', finder.upper(),' ',current_time, 'TAK', 'MB' ]
                return pack


