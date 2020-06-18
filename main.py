  
# -*- coding: utf-8 -*-
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.gridlayout import GridLayout
from kivy.core.window import Window
from kivy.clock import Clock
from kivy.lang.builder import Builder

import easygui,  shutil
from pathlib import Path

from get_stocks import open_chrome, read_yahoo_page, get_stocks, save_results
import time


i=0
class StockWindow(BoxLayout):
    def call_main(self):
        Clock.schedule_interval(self.main_func,1)

    def main_func(self,dt):
        global i, html, driver, stock, frequency
        if i == 0:
            #Clean scrollview previous results, if any and chek if info were provided
            stock_container = self.ids.stok_info
            stock_container.clear_widgets(stock_container.children)
            stock = self.ids.cod_acao.text
            frequency = self.ids.frequencia.text
            if stock == '' or frequency == 'Frequência':
                self.ids.head_text.text = "Erro - Entre com uma Ação e uma Frequência"
                return False
            elif len(stock) <= 4:
                self.ids.head_text.text = "Erro - Ação inválida"
                return False
            else:
                i = 1
                self.ids.head_text.text = "Abrindo Chrome..."
                return True #this will return back to clock, and every time it return screen is updated
        
        elif i == 1:
            try:
                driver = open_chrome(stock, frequency)
                i=2
                self.ids.head_text.text = "Conectando com Yahoo Finance..."
                return True
            except:
                self.ids.head_text.text = "Erro ao abrir o Chrome"
                i=0
                return False


        elif i == 2:
            # Aqui vai a função que puxa as ações
            html = read_yahoo_page(driver)
            if html == "Sem Internet":
                self.ids.head_text.text = "Erro - Sem Internet"
                i=0
                return False
            else:
                i=3
                self.ids.head_text.text = "Lendo dados do Yahoo Finance..."
                return True
        
        elif i == 3:
            history_data = get_stocks(html, driver)
            if history_data == 'Nenhum':
                self.ids.head_text.text = "Erro - Ação não encontrada"
                i=0
                return False
                
            for data in history_data:
                if len(data) < 7:
                    pass
                else:
                    stock_container = self.ids.stok_info
                    #add a BoxLayout inside of stock_container
                    box_layout = BoxLayout(size_hint_y=None, height=30)
                    stock_container.add_widget(box_layout)
                    #Stock info
                    date_ = data['Data']
                    open_ = data['Abertura']
                    max_ = data['Maxima']
                    min_ = data['Minima']
                    close = data['Fechamento']
                    adjst_close = data['Fechamento Ajustado']
                    volume = data['Volume']
                    #add labels inside of box_layout with the Stock info as text
                    box_layout.add_widget(Label(text=date_))
                    box_layout.add_widget(Label(text=open_))
                    box_layout.add_widget(Label(text=max_))
                    box_layout.add_widget(Label(text=min_))
                    box_layout.add_widget(Label(text=close))
                    box_layout.add_widget(Label(text=adjst_close))
                    box_layout.add_widget(Label(text=volume))

            save_results(history_data, stock)         
            self.ids.head_text.text = "Completo"
            i=0
            return False

    def export_data(self):
        path_origin = Path(__file__).parent.absolute()
        dest_path = easygui.filesavebox(msg = "Arquivo xls", title= "Task for Bot -", default="Histórico")
        path = str(path_origin) + "\\results.xls"
        shutil.copy(path, str(dest_path)+ '.xls')
                   



class Task_for_Bot(App):
    def build(self):
        Window.size = (1000, 600)
        Builder.load_string(open("main.kv", encoding="utf-8").read(), rulesonly=True)
        return StockWindow()
        

Task_for_Bot().run()