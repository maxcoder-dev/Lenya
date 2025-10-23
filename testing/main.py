import os
import sys
import time
import threading
import datetime
import speech_recognition as sr
from fuzzywuzzy import fuzz
import win32com.client
import random
import psutil
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from docx import Document

# ==============================
# Настройки голосового помощника
# ==============================
opts = {
    'alias': ('леня', 'лёня', 'лёнь', 'леонид', 'леша', 'лёнечка', 'линь', 'лёшаня', 'лешаня'),
    'tbr': ('скажи', 'покажи', 'расскажи', 'сколько', 'произнеси', 'похвастайся', 'сделай', 'расскажи'),
    'cmds': {
        'ctime': (
            'текущее время', 'сейчас времени', 'который час', 'сколько сейчас часов', 'время сейчас'
        ),
        'youtube': (
            'включи ютуб', 'го ютуб', 'ютуб', 'чекнуть ютуб', 'ютюбик', 'го ютубчик',
            'чекнуть ютубчик', 'ютубчик', 'youtube', 'открой youtube'
        ),
        'lichess': (
            'го шахматы', 'играть в шахматы', 'играть в чесс', 'играть в чес', 'погнали на личесс',
            'го на личес', 'шахматы', 'шахматюга', 'игра шахматы', 'шахматы играть',
            'открой шахматы', 'открой личес', 'открой личесс', 'открой ш'
        ),
        'explorer': (
            'проводник', 'файлы', 'открой проводник', 'открой файлы', 'открой файл'
        ),
        'browser': (
            'firefox', 'открой браузер', 'браузер', 'го браузер', 'вруби браузер', 'вруби firefox'
        ),
        'music': (
            'вруби музон', 'музыка', 'музон', 'go музон', 'врубай музон', 'давай музыку', 'музыку', 'вырубай музон'
        ),
        'github': (
            'го гитхаб', 'гитхаб', 'открой гитхаб', 'открой github', 'го github', 'go github', 'github'
        ),
        'talk': (
            'как дела', 'как ты', 'как у тебя дела', 'как ты сегодня', 'как у тебя дела сегодня',
            'как делишки', 'как сегодня у тебя дела?'
        ),
        'exit': (
            'выход', 'выйти', 'отмена', 'отменить', 'выходить', 'выйти из программы', 'выход из программы', 'отменить программу'
        ),
        'weather': (
            'погода', 'погода в москве', 'погода сегодня', 'сегодня погода', 'прогноз погоды', 'прогноз погода', 'прогноз погоды в москве',
            'погода в городе', 'погода сейчас'
        ),

        'hey': (
            'привет', 'ку', 'пр', 'здорово', 'здарова', 'привет леня', 'привет леонид', 'ку-ку'
        ),
        'comment_hard': (   #Будет работать не только на hard.
            'ты тупой', 'ты ужасный', 'ты плохой', 'плохой', 'ужасный', 'тупой', 'дибил', 'ты дибил', 'ты ужастен', 'ты не красивый',
            'ты очень плохой', 'ты очень ужасный', 'ты очень тупой', 'ты очень дибил', 'ты слишком ужастный', 'ты слишком плохой', 'ты слишком плох',
            'ты слишком тупой', 'ты слишком туп', 'ты очень туп', 'ты слишком дибильный'
        ),
        'chessbase': (
            'открой chessbase', 'открой чесбэйз', 'открой базу', 'го базу', 'давай чесбэйз', 'давай chessbase',
            'открой чесбэйс', 'давай чесбэйс'
        ),
        'dispz': (
            'открой диспетчер задач', 'диспетчер задач', 'го диспетчер задач', 'открой диспетчер задач, пожаоуйста',
            'диспетчер задач открыть', 'отрыть диспетчер задач'
        ),
        'pycharm': (
            'открой вскод', 'открой vscode', 'го vscode', 'go vscode', 'go вскод', 'го вскод',
            'vscode', 'вскод', 'го кодить', 'хочу кодить', 'кодинг', 'код', 'го кодинг', 'питон', 'пайтон',
            'python'
        ),
        'nomer': (
            'как решить номер', 'как решить тот номер', 'реши номер 235', 'го решим номер 235'
        ),
        'zoom': (
            'зум', 'zoom', 'открой зум', 'открой zoom', 'давай zoom', 'давай зум', 'го zoom', 'го зум',
            'открыть zoom', 'открыть зум'
        ),
        'discord': (
            'дискорд', 'discord', 'открой дискорд', 'открой discord', 'го дискорд', 'го discord', 'го в дискорд', 'го в discord',
            'дс', 'ds', 'открой дс', 'открой ds', 'го дс', 'го ds', 'го в дс', 'го в ds'
        ),
        'calc': (
            'калькулятор', 'сколько будет', 'а сколько', 'а сколько будет', 'сколько'
        ),
        'timer': (
            'установить таймер', 'установи таймер', 'таймер', 'го таймер'
        ),
        'thank_u': (
            'пасиб', 'пасиба', 'спасибо', 'спасибо, Бест!', 'спс', 'спасибки'
        ),

        'strax': (
            'страх', 'страшно', 'маме страшно', 'страшно маме'
        ),

        'gg_browser': (
            'закрой браузер', 'закрой файрфокс', 'закрыть броузер'
        ),
        'docx': (
            'прочитай', 'прочитай ворд', 'прочитай word', 'прочитай world', 'го читать', 'читать', 'читай', 'прочти', 'прочти word', 'прочти ворд'
        )
    }
}

# ==============================
# Реплики
# ==============================
bot_rand_replic_talk = [
    'У меня все хорошо. Как я могу помочь тебе сегодня?',
    'У меня все отлично, как всегда. Чем тебе помочь?',
    'У меня все как обычно, скучаю на AMD Ryzen 5. Чем могу быть полезен?',
    'У меня все как всегда... Вирусы, антивирусы, браузеры... А как у тебя?',
    'Все хорошо. Сижу тут, жду команды.'
]

bot_rand_replic_hey = [
    'Привет, кожанный',
    'Привет, Макс! Как дела?',
    'Привет, Бест!',
    'Хэй, кожанный',
    'Привет, человек!'
]

#Отвечаем на "Лёня, ты..."

bot_rand_replic_hard = [
    'Очень тонкое замечание, сэр',
    'Очень тонкое замечание, Макс'
]

bot_rand_replic_pycharm = [
    'Открываю pycharm. Приятного кодинга!',
    'Хорошего кодинга!',
    'Открываю pycharm, хорошо покодить!',
    'Хорошо, удачи в кодинге!'
]

# ==============================
# Инициализация голосового движка с поддержкой OneCore (Pavel)
# ==============================
speaker = win32com.client.Dispatch("SAPI.SpVoice")

tokenizer = win32com.client.Dispatch("SAPI.SpObjectTokenCategory")
tokenizer.SetId(r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech_OneCore\Voices", False)

tokens = tokenizer.EnumerateTokens()
pavel_found = False

for i in range(tokens.Count):
    token = tokens.Item(i)
    desc = token.GetDescription()
    if "Pavel" in desc:
        speaker.Voice = token
        pavel_found = True
        print(f"✅ Используется голос: {desc}")
        break

if not pavel_found:
    print("⚠️ Голос Pavel не найден, используется стандартный голос Irina.")

speaker.Rate = 2
speaker.Volume = 100

def speak(text):
    print(text)
    speaker.Speak(text)

def run_timer(seconds):
    time.sleep(seconds)
    speak(f'Таймер на {seconds} секунд сработал')

# ==============================
# Распознавание команд
# ==============================
def recognize_cmd(cmd):
    result = {'cmd': '', 'percent': 0}
    for c, variants in opts['cmds'].items():
        for v in variants:
            score = fuzz.ratio(cmd, v)
            if score > result['percent']:
                result['cmd'] = c
                result['percent'] = score
    return result

# ==============================
# Выполнение команд
# ==============================
def execute_cmd(cmd):
    if cmd == 'ctime':
        now = datetime.datetime.now()
        speak(f'Сейчас {now.hour}:{now.minute:02d}')

    elif cmd == 'youtube':
        speak('Открываю YouTube')
        os.system('start https://youtube.com')

    elif cmd == 'lichess':
        speak('Удачи в игре')
        os.system('start https://lichess.org/')

    elif cmd == 'explorer':
        speak('Секунду')
        os.system('start explorer')

    elif cmd == 'browser':
        speak('Открываю браузер')
        os.system('start https://example.com')

    elif cmd == 'music':
        speak('Ща будет пушка')
        os.system('start https://www.youtube.com/watch?v=uttVf8QPiN0')

    elif cmd == 'github':
        speak('Гитхаб открыт, только давай там норм коммиты делай')
        os.system('start https://github.com')

    elif cmd == 'talk':
        speak(random.choice(bot_rand_replic_talk))
    elif cmd == 'hey':
        speak(random.choice(bot_rand_replic_hey))

    elif cmd == 'exit':
        speak('Хорошо, пока!')
        stop_listening(wait_for_stop=False)
        sys.exit(0)
    elif cmd == 'strax':
        speak('Спакуха, Мариша, Так и должно быть. (доказано Лёней)')
    elif cmd == 'weather':
        import requests

        city = 'Москва'
        api_key = '79d1ca96933b0328e1c7e3e7a26cb347'
        url = f'https://api.openweathermap.org/data/2.5/weather?q={city}&units=metric&lang=ru&appid={api_key}'

        weather_data = requests.get(url).json()
        temperature = round(weather_data['main']['temp'])
        temperature_feels = round(weather_data['main']['feels_like'])

        speak(f'Сейчас в городе {city} {temperature} градусов по Цельсию')
        speak(f'Ощущается как {temperature_feels} градусов по Цельсию')

        # Рекомендации по одежде
        if temperature >= 30:
            speak('На улице очень жарко! Лучше надеть шорты, футболку и взять воду.')
        elif 20 <= temperature < 30:
            speak('Тепло и комфортно. Можно надеть шорты или лёгкую одежду.')
        elif 10 <= temperature < 20:
            speak('На улице прохладно. Лучше надеть штаны и лёгкую куртку или худи.')
        elif 0 <= temperature < 10:
            speak('Прохладно, советую надеть тёплую куртку, штаны и возможно шапку.')
        elif -10 <= temperature < 0:
            speak('Холодно! Лучше утеплиться: зимняя куртка, шапка, перчатки.')
        else:
            speak('Очень сильный мороз! Надень тёплую зимнюю одежду, шарф и перчатки. Будь аккуратнее на улице.')

    elif cmd == 'comment_hard':
        speak(random.choice(bot_rand_replic_hard))

    elif cmd == 'chessbase':
        speak('Открываю ChessBase')
        path_chessbase = r'C:\Program Files\ChessBase\CBase17\CBase17.exe'
        os.system(f'"{path_chessbase}"')
    elif cmd == 'dispz':
        speak('Открываю Диспетчер задач')
        os.system('start taskmgr')
    elif cmd == 'pycharm':
        speak('Открываю ВСкод')
        path_pycharm = r'C:\Users\Max\AppData\Local\Programs\Microsoft VS Code\Code.exe'
        os.system(f'"{path_pycharm}"')
    elif cmd == 'nomer':
        speak('Да хз я, че пристали со своими номерами')
    elif cmd == 'zoom':
        speak('Хорошо, хороших занятий и конференций!')
        path_zoom = r'C:\Users\Max\AppData\Roaming\Zoom\bin\Zoom.exe'
        os.system(f'"{path_zoom}"')
    elif cmd == 'discord':
        speak('Минутку')
        path_ds = r'C:\Users\Max\AppData\Local\Discord\app-1.0.9211\Discord.exe'
        os.system(f'"{path_ds}"')
    # Калькулятор по голосу
    elif cmd == 'calc':
        speak('Скажи пример для вычисления')
        recognizer = sr.Recognizer()
        with sr.Microphone() as source:
            audio = recognizer.listen(source)
        try:
            voice_expr = recognizer.recognize_google(audio, language='ru-RU').lower()
            # Заменим слова на символы
            voice_expr = voice_expr.replace('плюс', '+')
            voice_expr = voice_expr.replace('минус', '-')
            voice_expr = voice_expr.replace('умножить на', '*')
            voice_expr = voice_expr.replace('умножить', '*')
            voice_expr = voice_expr.replace('разделить на', '/')
            voice_expr = voice_expr.replace('разделить', '/')
            voice_expr = voice_expr.replace(' ', '')

            result = eval(voice_expr)
            speak(f'Результат вычисления: {result}')
        except Exception:
            speak('Не получилось распознать выражение')

    # Таймер по голосу
    elif cmd == 'timer':
        speak('Скажи количество секунд для таймера')
        recognizer = sr.Recognizer()
        with sr.Microphone() as source:
            audio = recognizer.listen(source)
        try:
            voice_seconds = recognizer.recognize_google(audio, language='ru-RU').lower()
            seconds = int(''.join(ch for ch in voice_seconds if ch.isdigit()))
            if seconds > 0:
                speak(f'Таймер запущен на {seconds} секунд')
                threading.Thread(target=run_timer, args=(seconds,), daemon=True).start()
            else:
                speak('Нулевой таймер не имеет смысла')
        except Exception:
            speak('Не удалось распознать число')
    elif cmd == 'thank_u':
        speak('Всегда рад помочь!')
    elif cmd == 'gg_browser':
        speak('Ок!')
        process_browser = ['firefox.exe', 'chrome.exe', 'yandex.exe', 'browser.exe', 'opera.exe']
        for process in psutil.process_iter():
            if process.name().lower() in process_browser:
                process.kill()
    elif cmd == 'docx':
        speak("Хорошо")
        speak("Выбери документ, который нужно прочитать в открывшимся проводнике.")
        def read_word_file():
            Tk().withdraw()  # скрыть окно Tkinter
            file_path = askopenfilename(
                title='Выберите Word-документ',
                filetypes=[('Word files', '*.docx')]
            )

            if not file_path:
                speak('Файл не выбран.')
                return

            try:
                doc = Document(file_path)
                full_text = []

                for para in doc.paragraphs:
                    if para.text.strip():
                        full_text.append(para.text.strip())

                if not full_text:
                    speak('Документ пустой.')
                    return

                speak('Начинаю чтение документа.')
                time.sleep(0.5)

                for paragraph in full_text:
                    speak(paragraph)
                    time.sleep(0.3)

                speak('Чтение документа завершено.')

            except Exception as e:
                print(f"Ошибка: {e}")
                speak('Не удалось прочитать документ.')

        read_word_file()    #вызываем функцию

    else:
        speak('Команда не распознана')

# ==============================
# Обработка звука
# ==============================
def callback(recognizer, audio):
    try:
        voice = recognizer.recognize_google(audio, language='ru-RU').lower()
        if any(a in voice for a in opts['alias']):
            cmd_text = voice
            for a in opts['alias']:
                cmd_text = cmd_text.replace(a, '').strip()
            for t in opts['tbr']:
                cmd_text = cmd_text.replace(t, '').strip()

            cmd = recognize_cmd(cmd_text)

            if cmd['percent'] > 50:
                threading.Thread(target=execute_cmd, args=(cmd['cmd'],), daemon=True).start()
            else:
                speak('Команда не распознана')

    except sr.UnknownValueError:
        pass
    except sr.RequestError:
        print('Ошибка доступа к сервису распознавания речи')

# ==============================
# Микрофон
# ==============================
r = sr.Recognizer()
r.energy_threshold = 300
r.dynamic_energy_threshold = True

print("Доступные микрофоны:")
for i, name in enumerate(sr.Microphone.list_microphone_names()):
    print(f"[{i}] {name}")

m = sr.Microphone(device_index=None)
with m as source:
    r.adjust_for_ambient_noise(source, duration=1)
    print("Порог шума установлен:", r.energy_threshold)

# ==============================
# Запуск прослушивания
# ==============================
speak('Привет, Максим. Леня слушает.')
stop_listening = r.listen_in_background(m, callback)
print("Леня запущен. Говорите команды.")

# ==============================
# Основной цикл
# ==============================
try:
    while True:
        time.sleep(0.1)
except KeyboardInterrupt:
    stop_listening(wait_for_stop=False)
    print('Программа завершена.')