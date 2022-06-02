from ast import Del
from distutils.util import execute
import pyautogui as aut
import time
from PySimpleGUI import PySimpleGUI as sg


#layout
#theme_name_list = sg.theme_list()
#print(theme_name_list)
# def layout_login():
#     sg.theme('DarkGreen7')
#     layout = [
#         [sg.Text('usuario', size=(23,1))],
#         [sg.Input(size=(23,1))]
#         [sg.Button('Sim',size=(10,1)),sg.Button('Não',size=(10,1))]
#     ]
#     return sg.Window('Login',layout=layout,finalize=True)
#def layout_autobot():
sg.theme('DarkBrown2')
font = ("Helvetica")
layout = [    [sg.Text('Deseja começar a Automação ?',font= font, size=(27,1))],
    
    [sg.Button('Sim',key='Sim',font= font,size=(10,1)),sg.Button('Não',key='Não',font= font,size=(10,1))]
    ]
    #return sg.Window('autobot',layout=layout,finalize=True)
    #Janela

janela = sg.Window('Preparo de Reunião', layout,font= font, size=(280,70))

    #Ler Eventos

while True:
    eventos, valores = janela.read()
    if eventos == 'Sim':
        #                                           AUTOMAÇÃO PARA REUNIÃO DE SEXTA
        aut.moveTo(387,1061)#powerpoint
        aut.click()
        time.sleep(2)
        aut.moveTo(141,380)#tela2
        aut.click()
        time.sleep(1)
        aut.moveTo(972,524)
        aut.click()
        time.sleep(1)
        aut.press('Del')
        time.sleep(1)
                                                            #abrindo o outlook
        aut.moveTo(277,1063)#abrir outlook
        aut.click()
        time.sleep(4)
        aut.moveTo(1979,788)#abrir o Backup
        aut.click()
        time.sleep(1)
        aut.moveTo(2222,428)
        aut.click()
        time.sleep(1)
        aut.write("linfs01")
        time.sleep(4)
        aut.moveTo(3117,1015)                                #reduzir zoon
        aut.click()
        time.sleep(1)
        aut.click()
        time.sleep(1)
        aut.click()
        time.sleep(1)
        aut.hotkey('shift','win','s')
        time.sleep(3)
        aut.dragTo(2471,579, button='left')
        aut.dragTo(3242,746, 1, button='left')
        time.sleep(1)
                                                           #colar o backup na Reunião
        aut.moveTo(914,478)
        aut.doubleClick()
        time.sleep(1)
        aut.press('Del')
        time.sleep(1)
        aut.hotkey('ctrl','v')
                                                           #abrindo a tela do antivirus(reunião)
        aut.moveTo(138,685)
        aut.doubleClick()
        time.sleep(1)
        aut.moveTo(1021,522)
        aut.doubleClick()
        aut.press('Del')
        time.sleep(1)
        aut.moveTo(1979,766)                                 #abrir tela do antvirus
        aut.click()
        time.sleep(1)
        aut.moveTo(2221,431)
        aut.click()
        time.sleep(1)
        aut.write("LIN")
        aut.moveTo(2232,524)
        aut.click()
        time.sleep(5)
        aut.press('down', presses=10)
        time.sleep(1)
        aut.moveTo(3063,755)
        aut.doubleClick()
        time.sleep(2)
        aut.moveTo(3115,1015)#zoon
        aut.click()
        time.sleep(1)
        aut.click()
        time.sleep(1)
        aut.click()
        time.sleep(1)
        aut.click()
        time.sleep(1)
        aut.press('Enter')
        aut.hotkey('shift','win','s')                        #tira print 
        time.sleep(3)
        aut.dragTo(2448,554, button='left')
        aut.dragTo(3251,940, 1, button='left')
        time.sleep(1)
        aut.moveTo(970,530)
        aut.click()
        aut.hotkey('ctrl','v')
                                                        #buscando informações dos chamados
        aut.moveTo(204,1064)
        aut.click()
        time.sleep(1)
        aut.moveTo(128,15) #aba dos chamados
        aut.click()
        time.sleep(1)
        aut.moveTo(92,362) #Pesquisar
        aut.click()
        time.sleep(1)
        aut.moveTo(278,682) #Campo de Analista
        aut.click()
        time.sleep(1)
        aut.moveTo(413,717) #Campo de pesquisar
        aut.click()
        time.sleep(1)
        aut.write('Ricardo Rodrigues Lima')
        time.sleep(1)
        aut.press('Tab')
        time.sleep(1)
        aut.moveTo(403,1025)
        aut.click()
        time.sleep(1)
        aut.dragTo(1376,173, button='left')             #Ajustar a tela de chamados
        aut.dragTo(1248,176, 1, button='left')
        aut.hotkey('shift','win','s')
        time.sleep(2)
        aut.dragTo(685,194, button='left')             #tirar print da tela de chamados
        aut.dragTo(1660,281, 1, button='left')
        time.sleep(1)
        aut.moveTo(389,1064)
        aut.click()
        time.sleep(1)
        aut.moveTo(136,847)
        aut.click()

        aut.moveTo(886,356)   
        aut.click()                     # apagar Wagner
        aut.press('Del')
        time.sleep(1)

        aut.moveTo(912,484)             #apagar Ricardo
        aut.click()
        aut.press('Del')
        time.sleep(1)

        aut.moveTo(850,663)             #apagar Decio
        aut.click()
        aut.press('Del')
        time.sleep(1)

        aut.moveTo(878,829)             #apagar Antonio
        aut.click()
        aut.press('Del')
        time.sleep(1)
        
        aut.hotkey('ctrl','v')          #Colar chamados Ricardo
        aut.dragTo(1043,566, button='left')             #Ajustar a tela de chamados na Apresentação
        aut.dragTo(949,488, 1, button='left')

        aut.moveTo(204,1064)
        aut.click()
        aut.moveTo(413,717) #Campo de pesquisar
        aut.click()
        time.sleep(1)
        aut.click()
        aut.doubleClick()
        time.sleep(1)
        aut.press('Del')
        aut.write('Wagner Ferreira')
        time.sleep(1)
        aut.press('Tab')
        time.sleep(1)
        aut.moveTo(403,1025)
        aut.click()
        time.sleep(1)
        aut.hotkey('shift','win','s')
        time.sleep(2)
        aut.dragTo(685,194, button='left')             #tirar print da tela de chamados
        aut.dragTo(1660,281, 1, button='left')
        time.sleep(1)
        aut.moveTo(389,1064)
        aut.click()
        aut.moveTo(1695,559)
        aut.click()
        aut.hotkey('ctrl','v')                           #Colar chamados WAGNER
        aut.click()                         
        aut.dragTo(1064,567, button='left')             #Ajustar a tela de chamados na Apresentação
        aut.dragTo(977,344, 1, button='left')
                                                        #Pesquisar chamados Decio
        aut.moveTo(204,1064)
        aut.click()
        aut.moveTo(413,717) #Campo de pesquisar
        aut.click()
        time.sleep(1)
        aut.click()
        aut.doubleClick()
        time.sleep(1)
        aut.press('Del')
        aut.write('Decio Greco')
        time.sleep(1)
        aut.press('Tab')
        time.sleep(1)
        aut.moveTo(403,1025)
        aut.click()
        time.sleep(1)
        aut.hotkey('shift','win','s')
        time.sleep(2)
        aut.dragTo(685,194, button='left')             #tirar print da tela de chamados
        aut.dragTo(1660,281, 1, button='left')
        time.sleep(1)
        aut.moveTo(389,1064)
        aut.click()
        aut.moveTo(1695,559)
        aut.click()
        aut.hotkey('ctrl','v')                           #Colar chamados Decio
        aut.click()                         
        aut.dragTo(1049,566, button='left')             #Ajustar a tela de chamados na Apresentação
        aut.dragTo(963,651, 1, button='left')
                                                                #Pesquisar chamados Antonio
        aut.moveTo(204,1064)
        aut.click()
        aut.moveTo(413,717) #Campo de pesquisar
        aut.click()
        time.sleep(1)
        aut.click()
        aut.doubleClick()
        time.sleep(1)
        aut.press('Del')
        aut.write('Antonio Marcos')
        time.sleep(1)
        aut.press('Tab')
        time.sleep(1)
        aut.moveTo(403,1025)
        aut.click()
        time.sleep(1)
        aut.hotkey('shift','win','s')
        time.sleep(2)
        aut.dragTo(685,194, button='left')             #tirar print da tela de chamados
        aut.dragTo(1660,281, 1, button='left')
        time.sleep(1)
        aut.moveTo(389,1064)
        aut.click()
        aut.moveTo(1695,559)
        aut.click()
        aut.hotkey('ctrl','v')                           #Colar chamados Antonio
        aut.click()                         
        aut.dragTo(1063,563, button='left')             #Ajustar a tela de chamados na Apresentação
        aut.dragTo(980,819, 1, button='left')
                                                        #Tela de chamados em Espera
        aut.moveTo(120,983)
        aut.click()
        aut.moveTo(1044,556)
        aut.doubleClick()
        aut.press('Del')
        
        aut.moveTo(205,1057)                            #Abrir google
        aut.click()
        aut.moveTo(466,136)
        aut.click()
        time.sleep(1)
        aut.hotkey('shift','win','s')
        time.sleep(2)
        aut.dragTo(251,520, button='left')             #Ajustar a tela de chamados na Apresentação
        aut.dragTo(1146,738, 1, button='left')
        time.sleep(2)
        aut.moveTo(391,1059)
        aut.click()
        aut.moveTo(1044,556)
        aut.doubleClick()
        aut.hotkey('ctrl','v')
        time.sleep(2)
        aut.moveTo(142,807)
        aut.click()
        aut.press('down', presses=4)
        aut.moveTo(1037,568)
        aut.click()
        aut.press('Del')
        time.sleep(1)
        aut.moveTo(205,1057)                            #Abrir google
        aut.click()
        time.sleep(1)
        aut.moveTo(746,18)                              #Abrir Aba
        aut.click()
        time.sleep(1)
        aut.moveTo(155,85)
        aut.click()
        time.sleep(2)
        aut.moveTo(1266,324)
        aut.click()
        aut.write('admin')
        time.sleep(1)
        aut.moveTo(1251,365)
        aut.click()
        aut.write('S3gur4nc4')
        time.sleep(1)
        aut.moveTo(1272,435)
        aut.click()
        time.sleep(3)
        aut.moveTo(164,291)
        time.sleep(1)
        aut.dragTo(164,291)             #Ajustar a tela de chamados na Apresentação
        aut.dragTo(122,265, 1,)
        aut.click()
        aut.hotkey('shift','win','s')
        aut.dragTo(200,146, button='left')             #Ajustar a tela de chamados na Apresentação
        aut.dragTo(946,559, 1, button='left')
        time.sleep(1)
        aut.moveTo(386,1060)
        aut.click()
        aut.hotkey('ctrl','v')



    if eventos == sg.WIN_CLOSED or eventos == 'Não':
        break
