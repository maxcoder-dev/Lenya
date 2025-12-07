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
import replics as rep
import requests

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
    for c, variants in rep.opts['cmds'].items():
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
        speak(random.choice(rep.bot_rand_replic_talk))
    elif cmd == 'hey':
        speak(random.choice(rep.bot_rand_replic_hey))

    elif cmd == 'exit':
        speak('Хорошо, пока!')
        stop_listening(wait_for_stop=False)
        sys.exit(0)
    elif cmd == 'strax':
        bot_rand_name2 = random.choice(rep.bot_rand_name)
        speak(f'Спакуха, {bot_rand_name2}, Так и должно быть. (доказано Лёней)')   #Рандомное имя
    elif cmd == 'weather':

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
        speak(random.choice(rep.bot_rand_replic_hard))

    elif cmd == 'chessbase':
        speak('Открываю ChessBase')
        path_chessbase = r'C:\Program Files\ChessBase\CBase17\CBase17.exe'
        os.system(f'"{path_chessbase}"')
    elif cmd == 'dispz':
        speak('Открываю Диспетчер задач')
        os.system('start taskmgr')
    elif cmd == 'info':
        speak('Я - голосовой ассистент Лёня и могу много всего. Прочитать войну и мир, открыть ютубчик и пожелать приятного просмотра, дать советы как одется сегодня и многое другое.')
    elif cmd == 'code':
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
    elif cmd == 'zapros':
        # Запрос голосом -> поиск в интернете -> чтение результатов
        def listen_for_query(prompt='Скажи, что ищешь'):
            speak(prompt)
            recognizer = sr.Recognizer()
            with sr.Microphone() as source:
                recognizer.adjust_for_ambient_noise(source, duration=0.6)
                audio = recognizer.listen(source, phrase_time_limit=8)
            try:
                q = recognizer.recognize_google(audio, language='ru-RU').lower()
                return q
            except sr.UnknownValueError:
                return ''
            except sr.RequestError:
                speak('Ошибка сервиса распознавания речи.')
                return ''

        def search_duckduckgo(query):
            """Используем DuckDuckGo Instant Answer API (без ключа, с нормальным User-Agent)."""
            try:
                url = 'https://api.duckduckgo.com/'
                params = {
                    'q': query,
                    'format': 'json',
                    'no_html': 1,
                    'skip_disambig': 1,
                    'kl': 'ru-ru'
                }
                headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
                r = requests.get(url, params=params, headers=headers, timeout=8)
                data = r.json()
            except Exception as e:
                print('DDG error:', e)
                return []

            results = []
            # Основной ответ (Abstract)
            abstract = data.get('AbstractText') or ''
            abstract_url = data.get('AbstractURL')
            if abstract:
                results.append({
                    'title': data.get('Heading') or query,
                    'snippet': abstract,
                    'url': abstract_url or ''
                })

            # RelatedTopics
            for item in data.get('RelatedTopics', []):
                if isinstance(item, dict):
                    text = item.get('Text', '')
                    first_url = item.get('FirstURL', '')
                    if text:
                        results.append({'title': text.split(' - ')[0], 'snippet': text, 'url': first_url})
            return results

        def search_wikipedia(query):
            """Запрос к Wikipedia API с User-Agent (иначе часто блокирует)."""
            try:
                url = 'https://ru.wikipedia.org/w/api.php'
                params = {
                    'action': 'opensearch',
                    'search': query,
                    'limit': 5,
                    'namespace': 0,
                    'format': 'json',
                    'utf8': 1
                }
                headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
                r = requests.get(url, params=params, headers=headers, timeout=8)
                r.raise_for_status()  # выбросит исключение, если код != 200
                arr = r.json()
                titles, descs, urls = arr[1], arr[2], arr[3]
                results = []
                for t, d, u in zip(titles, descs, urls):
                    if t or d:
                        results.append({'title': t, 'snippet': d, 'url': u})
                return results
            except Exception as e:
                print('Wiki error:', e)
                return []


        def fetch_page_first_paragraphs(url, max_paragraphs=3):
            """Пытаемся получить первые абзацы HTML-страницы простым способом (без bs4)."""
            try:
                r = requests.get(url, timeout=8, headers={'User-Agent': 'Mozilla/5.0'})
                html = r.text
                # грубая очистка: извлечь теги <p>
                paragraphs = []
                for part in html.split('<p'):
                    if '</p>' in part:
                        inner = part.split('>', 1)[1].split('</p>', 1)[0]
                        # удалить теги внутри
                        plain = inner
                        # простая очистка тегов
                        while '<' in plain and '>' in plain:
                            a = plain.find('<')
                            b = plain.find('>', a)
                            if a == -1 or b == -1:
                                break
                            plain = plain[:a] + plain[b+1:]
                        plain = plain.strip()
                        if plain and len(plain) > 30:
                            paragraphs.append(plain)
                        if len(paragraphs) >= max_paragraphs:
                            break
                return paragraphs
            except Exception as e:
                print('Fetch page error:', e)
                return []

        # --- основной рабочий поток ---
        query = listen_for_query('Что ищем в интернете?')
        if not query:
            speak('Не расслышал запрос. Попробуй ещё раз.')
            return

        speak(f'Ищу в интернете: {query}')
        results = []

        # 1) быстрый ответ DDG
        ddg = search_duckduckgo(query)
        if ddg:
            results.extend(ddg)

        # 2) Wikipedia как надёжный источник
        wiki = search_wikipedia(query)
        # добавляем те, которых ещё нет (по url или title)
        seen = set((r.get('url') or r.get('title')).lower() for r in results)
        for item in wiki:
            key = (item.get('url') or item.get('title')).lower()
            if key not in seen:
                results.append(item)
                seen.add(key)

        # Обрежем до 2-х результатов
        results = results[:2]

        if not results:
            speak('Ничего не найдено. Попробуй переформулировать запрос.')
            return

        # Читаем кратко топ-результаты
        speak(random.choice(rep.random_replic_search))
        for i, ritem in enumerate(results, start=1):
            title = ritem.get('title') or 'без названия'
            snippet = ritem.get('snippet') or ''
            speak(f'Результат {i}: {title}. {snippet[:250]}')
            time.sleep(0.2)

        # Спросим пользователя, какой результат открыть / прочитать
        speak('Хочешь, чтобы я прочитал полную статью одного из этих результатов? Скажи номер (или "нет").')
        answer = listen_for_query('Говори номер или "нет"')
        if not answer:
            speak('Не расслышал. Отмена.')
            return

        # Попробуем извлечь цифру
        chosen = None
        for token in answer.split():
            if token.isdigit():
                n = int(token)
                if 1 <= n <= len(results):
                    chosen = n - 1
                    break
        if chosen is None:
            # может сказать "один", "два" на русском — простая мапа
            mapping = {'ноль':0,'один':1,'два':2,'три':3,'четыре':4,'пять':5}
            for word in mapping:
                if word in answer:
                    n = mapping[word]
                    if 1 <= n <= len(results):
                        chosen = n - 1
                        break

        if chosen is None:
            if 'нет' in answer or 'не надо' in answer or 'отмена' in answer:
                speak('Хорошо, не буду открывать.')
                return
            else:
                speak('Не понял номер. Отмена.')
                return

        sel = results[chosen]
        sel_url = sel.get('url') or ''
        sel_title = sel.get('title') or 'статья'
        speak(f'Открываю и читаю: {sel_title}')

        if sel_url:
            paras = fetch_page_first_paragraphs(sel_url, max_paragraphs=4)
            if paras:
                for p in paras:
                    speak(p)
                    time.sleep(0.4)
                speak('Чтение статьи завершено.')
                # Предложим открыть в браузере
                speak('Открыть страницу в браузере? Скажи "да" или "нет".')
                yn = listen_for_query()
                if 'да' in yn or 'открыть' in yn:
                    os.system(f'start {sel_url}')
                else:
                    speak('Хорошо.')
            else:
                speak('Не удалось извлечь текст с этой страницы. Открою ссылку в браузере.')
                os.system(f'start {sel_url}')
        else:
            # Если URL отсутствует, просто повторим сниппет
            speak(sel.get('snippet') or 'Подробностей нет.')

       

    else:
        speak('Команда не распознана')

# ==============================
# Обработка звука
# ==============================
def callback(recognizer, audio):
    try:
        voice = recognizer.recognize_google(audio, language='ru-RU').lower()
        if any(a in voice for a in rep.opts['alias']):
            cmd_text = voice
            for a in rep.opts['alias']:
                cmd_text = cmd_text.replace(a, '').strip()
            for t in rep.opts['tbr']:
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
