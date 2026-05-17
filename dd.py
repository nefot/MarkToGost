import pyautogui
import keyboard
import time

ROWS = 18
COLS = 7

pyautogui.PAUSE = 0.05

print("Наведи мышку на ЛЕВЫЙ ВЕРХНИЙ угол первой ячейки")
time.sleep(5)
x1, y1 = pyautogui.position()
print(f"Первая точка: {x1}, {y1}")

print("Наведи мышку на ПРАВЫЙ НИЖНИЙ угол последней ячейки")
time.sleep(5)
x2, y2 = pyautogui.position()
print(f"Вторая точка: {x2}, {y2}")

cell_width = (x2 - x1) / COLS
cell_height = (y2 - y1) / ROWS

print("Старт через 5 секунд...")
time.sleep(5)

data = [
    ["№","Обозначение","Откуда","Куда","Тип кабеля","Длина","Кол."],
    ["1","ПК-0.1","ODF U13 порт FO-2","Коммутатор ядра Edimax GS-5424LX U03 SFP1","Оптический LC-LC","0,5 м","1"],
    ["2","ПК-0.2","ODF U13 порт FO-3","Коммутатор ядра Edimax GS-5424LX U03 SFP2","Оптический LC-LC","0,5 м","1"],
    ["3","ПК-0.3","ODF U13 порт FO-4","Коммутатор ядра Edimax GS-5424LX U03 SFP3","Оптический LC-LC","0,5 м","1"],
    ["4","ПК-0.4","ODF U13 порт FO-5","Коммутатор ядра Edimax GS-5424LX U03 SFP4","Оптический LC-LC","0,5 м","1"],
    ["5","ПК-0.5","ODF U13 порт FO-6","Коммутатор ядра Edimax GS-5424LX U03 SFP5","Оптический LC-LC","0,5 м","1"],
    ["6","ПК-0.6","Маршрутизатор Router порт LAN","Коммутатор ядра Edimax GS-5424LX U03 порт 1","UTP Cat.5e","0,5 м","1"],
    ["7","ПК-0.7","Коммутатор ядра Edimax GS-5424LX U03 порт 2","Сервер (файловый)","UTP Cat.5e","0,5 м","1"],
    ["8","ПК-0.8","Коммутатор ядра Edimax GS-5424LX U03 порт 3","Сервер (контроллер)","UTP Cat.5e","0,5 м","1"],
    ["9","ПК-0.9","Коммутатор ядра Edimax GS-5424LX U03 порт 4","Сервер (архив/резерв)","UTP Cat.5e","0,5 м","1"],
    ["10","ПК-0.10","Коммутатор ядра Edimax GS-5424LX U03 порт 5","Рабочая станция Admin","UTP Cat.5e","0,5 м","1"],
    ["11","ПК-0.11","Коммутатор ядра Edimax GS-5424LX U03 порт 6","IP-камера CAM 1.1","UTP Cat.5e","0,5 м","1"],
    ["12","ПК-0.12","Router порт WAN","Интернет-шлюз","UTP Cat.5e","1,0 м","1"],
    ["13","ПК-0.13","Коммутатор ядра Edimax GS-5424LX U03","PDU U21","Кабель питания C13","1,5 м","1"],
    ["14","ПК-0.14","Сервер (файловый)","PDU U21","Кабель питания C13","1,5 м","1"],
    ["15","ПК-0.15","Сервер (контроллер)","PDU U21","Кабель питания C13","1,5 м","1"],
    ["16","ПК-0.16","Сервер (архив/резерв)","PDU U21","Кабель питания C13","1,5 м","1"],
    ["17","ПК-0.17","ИБП SRT3000RMXLI U46–48","PDU U21","Кабель питания C13","1,0 м","1"]
]

for r in range(len(data)):
    for c in range(len(data[r])):

        if keyboard.is_pressed('esc'):
            print("Остановлено.")
            exit()

        x = int(x1 + c * cell_width + cell_width / 2)
        y = int(y1 + r * cell_height + cell_height / 2)

        # Двойной клик чтобы войти в ячейку
        pyautogui.doubleClick(x, y)
        time.sleep(0.3)

        # Выделить всё и удалить старый текст
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.05)
        pyautogui.press('backspace')
        time.sleep(0.05)

        # Написать текст
        keyboard.write(data[r][c], delay=0.02)
        time.sleep(0.1)

        # Escape фиксирует без смещения курсора
        pyautogui.press('escape')
        time.sleep(0.1)

print("Таблица заполнена.")