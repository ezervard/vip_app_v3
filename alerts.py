from kivy.uix.popup import Popup
from kivy.uix.label import Label

class Alerts():
    def show_popup(self, text, title):
        content = Label(text=text, font_size=20)
        popup = Popup(title=title, content=content, size_hint=(None, None), size=(250, 200))
        popup.open()