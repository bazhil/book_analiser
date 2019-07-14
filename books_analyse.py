# -*- coding: utf-8 -*-
import collections
import os
import docx2txt
import math
import json
import pymorphy2

class Books_analiser:
    """
    Класс, который содержит функционал для анализа книг.
    В качестве результата своей работы возвращает JSON.
    """
    def __init__(self):
        self.books_info = []
        self.book_split = None

    def preprocess_text(self, text):
        """
        Служебная Функция, которая убирает из текста все лишние символы
        :param text: строка
        :return:
        """
        text = text.replace('%', '')
        text = text.replace('(', '')
        text = text.replace(')', '')
        text = text.replace('...', '.')
        text = text.replace('\n', ' ')
        text = text.replace('[', ' ')
        text = text.replace(']', ' ')
        text = text.replace('"', ' ')
        text = text.replace('  ', ' ')
        return text

    def text_prepare(self, book):
        """
        Подготавливает книгу к анализу
        :param book: книга в фломате .docx
        :return: массив слов из книги
        """
        book_str = docx2txt.process(book)
        book_str = self.preprocess_text(book_str)
        global book_split
        book_split = book_str.split(' ')
        return book_split, len(book_split)

    def unique_words_counter(self, book_split):
        """
        Функция, которая подсчтывает количество уникальных слов
        :param book_split: массив слов из книги
        :return: количество уникальных слов в книге, словарь с количеством частей речи, пять и десять наиболее встречаемых слов
        """
        book_counter = collections.Counter(book_split)
        words = list(book_counter)
        normal_words = []
        m = pymorphy2.MorphAnalyzer()

        for word in words:
            lemmas = m.parse(word)[0]
            normal_word = lemmas.normal_form
            normal_words.append(normal_word)
        words_count = len(words)
        # Отсекаем повторяющиеся слова из словаря
        unique_words = list(set(normal_words))
        # Определяем количество слов
        quantity_of_unique_words = len(unique_words)
        return words_count, quantity_of_unique_words

    def most_common_words(self, book_split, ind=int):
        """
        Функция, которая возвращает заданное количество наиболее встречаемых в книге слов (длинной от 4-х символов)
        :param book_split: массив слов из книги
        :param ind: количество наиболее встречаемых в книге слов
        :return: словарь
        """
        filtred_book = []
        m = pymorphy2.MorphAnalyzer()
        # Отбираем все существительные, глаголы и прилагательные для дальнейшего анализа самых распространенных слов
        for word in book_split:
            lemmas = m.parse(word)[0]
            normal_word = lemmas.normal_form
            if 'NOUN' in lemmas.tag:
                filtred_book.append(normal_word)
            if 'ADJF' in lemmas.tag:
                filtred_book.append(normal_word)
            if 'ADJS' in lemmas.tag:
                filtred_book.append(normal_word)
            if 'VERB' in lemmas.tag:
                filtred_book.append(normal_word)
            if 'INFN' in lemmas.tag:
                filtred_book.append(normal_word)
        book_counter = collections.Counter(filtred_book)
        most_common_words = book_counter.most_common(ind)
        return most_common_words

    def tags_counter(self, book_split):
        """
        Функция, которая преобразует слова в нормальную форму и подсчитывает количество частей речи
        :param book_split: массив слов из книги
        :return: словарь с количеством частей речи
        """
        tags_counter = {}
        book_counter = collections.Counter(book_split)
        words = list(book_counter)
        m = pymorphy2.MorphAnalyzer()

        # look transcripts here: http://opencorpora.org/dict.php?act=gram
        # счетчик частей речи
        post = 0
        # счетчик существительных
        noun = 0
        # счетчик прилагательных (полных)
        adjf = 0
        # счетчик прилагательных (кратких)
        adjs = 0
        # счетчик компративов
        comp = 0
        # счетчик глаголов (личная форма)
        verb = 0
        # счетчик глаголов (инфинитив)
        infn = 0
        # счетчик причастий (полных)
        prtf = 0
        # счетчик присчатий (кратких)
        prts = 0
        # счетчик деепричастий
        grnd = 0
        # счетчик числительных
        numr = 0
        # счетчик наречий
        advb = 0
        # счетчик местоимений-существительных
        npro = 0
        # счетчик предикативов
        pred = 0
        # счетчик предлогов
        prep = 0
        # счетчик союзов
        conj = 0
        # счетчик частиц
        prcl = 0
        # счетчик междометий
        intj = 0

        for word in words:
            lemmas = m.parse(word)[0]
            # Считаем части рчи
            if 'POST' in lemmas.tag:
                post += 1
            if 'NOUN' in lemmas.tag:
                noun += 1
            if 'ADJF' in lemmas.tag:
                adjf += 1
            if 'ADJS' in lemmas.tag:
                adjs += 1
            if 'COMP' in lemmas.tag:
                comp += 1
            if 'VERB' in lemmas.tag:
                verb += 1
            if 'INFN' in lemmas.tag:
                infn += 1
            if 'PRTF' in lemmas.tag:
                prtf += 1
            if 'PRTS' in lemmas.tag:
                prts += 1
            if 'GRND' in lemmas.tag:
                grnd += 1
            if 'NUMR' in lemmas.tag:
                numr += 1
            if 'ADVB' in lemmas.tag:
                advb += 1
            if 'NPRO' in lemmas.tag:
                npro += 1
            if 'PRED' in lemmas.tag:
                pred += 1
            if 'PREP' in lemmas.tag:
                prep += 1
            if 'CONJ' in lemmas.tag:
                conj += 1
            if 'PRCL' in lemmas.tag:
                prcl += 1
            if 'INTJ' in lemmas.tag:
                intj += 1
        # Собираем словарь с данными по количеству частей речи
        tags_counter['ЧастиРечи'] = post
        tags_counter['Существительные'] = noun
        tags_counter['ПрилагательныеПолные'] = adjf
        tags_counter['ПрилагательныеКраткие'] = adjs
        tags_counter['Компративы'] = comp
        tags_counter['ГлаголыЛичнаяФорма'] = verb
        tags_counter['ГлаголыИнфинитив'] = infn
        tags_counter['ПричастияПолные'] = prtf
        tags_counter['ПричастияКраткие'] = prts
        tags_counter['Деепричастия'] = grnd
        tags_counter['Числительные'] = numr
        tags_counter['Наречия'] = advb
        tags_counter['МестоименияСуществительные'] = npro
        tags_counter['Предикативы'] = pred
        tags_counter['Предлоги'] = prep
        tags_counter['Союзы'] = conj
        tags_counter['Частицы'] = prcl
        tags_counter['Междометия'] = intj

        return tags_counter


    def unique_words_in(self, book_split, index=int):
        """
        Функция, которая возвращает количество уникальных слов на заданный объем текста
        :param index: объем текста на которое опредедяется число уникальных слов
        :param book_split: массив слов книги
        :return: количество уникальных слов в заданном интервале index
        """
        unique_words = []
        # Считаем части речи
        words = list(book_split[:index])
        m = pymorphy2.MorphAnalyzer()

        for word in words:
            lemmas = m.parse(word)[0]
            normal_word = lemmas.normal_form
            unique_words.append(normal_word)
        unique_words_counter = collections.Counter(unique_words)
        unique_words = list(set(unique_words_counter))
        quantity_of_unique_words = len(unique_words)
        return quantity_of_unique_words

    def count_centencions(self, book):
        """
        Функция, которая выводит количество предложений в тексте и среднюю длину предложения
        :param book: файл книги
        :return: количество предложений в тексте и среднюю длину предложения
        """
        book_str = docx2txt.process(book)
        centencions_list = book_str.split('.')
        centencions_count = len(centencions_list)
        centencion_sum_length = 0
        for centencion in centencions_list:
            centencion_sum_length += len(centencion)
        average_centencion_length = math.ceil(centencion_sum_length/centencions_count)
        return centencions_count, average_centencion_length

    def sequense(self, book, character=str):
        """
        Возвращает частоту вхождения определенного символа в текст произведения
        :param book: файл книги
        :param character: символ
        :return:
        """
        book_str = docx2txt.process(book)
        count = 0
        for char in book_str:
            if char == character:
                count += 1
        return count

    def brut_dialog_procents(self, book):
        """
        Функция, которая грубо высчитывает количество диалогов в книге
        :param book: файл книги
        :return:
        """
        # Принимаем, что 70% длинных тире или тире стоят в диалогах
        dialogs = self.sequense(book, '—') * 0.7
        if dialogs == 0:
            dialogs = self.sequense(book, '–') * 0.7
        cents = self.count_centencions(book)[0]
        dialogs_procents = math.floor((dialogs / cents) * 100)
        return dialogs_procents

    def split(self, a, n):
        """
        Вспомогательная функция, которая помогает разделить список на подсписки заданной длины
        :param a: целевой список
        :param n: длина подсписка
        :return:
        """
        k, m = divmod(len(a), n)
        return (a[i * k + min(i, m):(i + 1) * k + min(i + 1, m)] for i in range(n))

    def dinamical_change_asz(self, book_split, ind=3000):
        """
        Функция, которая определяет динамику изменения количества уникальных слов на 3000 слов от начала и до конца книги
        :param book_split: список слов книги
        :param ind: 3000 - число слов, для которого динамически определяется уникальное количество слов
        :return: список словарей с данными о динамике изменения количества уникальных слов
        """
        sublist_length = len(book_split) // ind
        fragments = list(self.split(book_split, sublist_length))
        x = list(range(len(fragments)))
        y = [self.unique_words_counter(i)[0] for i in fragments]
        change_asz_data = dict(zip(x, y))
        return change_asz_data

    def analyze(self, book):
        """
        Функция, которая анализирует книгу по ряду параметров и результат записывает в json
        :param book: файл книги
        :return: json-объект
        """
        global books_info
        book_data = {}
        split_book = self.text_prepare(book)[0]
        head, tail = os.path.split(book)
        book_data['Автор'] = tail.split(' - ')[0]
        book_data['НазваниеКниги'] = tail.split(' - ')[1].replace('.docx', '')
        book_data['КоличествоСловГрубо'] = self.text_prepare(book)[1]
        book_data['КоличествоСловТочно'] = self.unique_words_counter(split_book)[0]
        book_data['КоличествоУникальныхСлов'] = self.unique_words_counter(split_book)[1]
        book_data['КоличествоЧастейРечи'] = self.tags_counter(split_book)
        book_data['ПятьСамыхРаспространенныхСлов'] = self.most_common_words(split_book, 5)
        book_data['ДесятьСамыхРаспространенныхСлов'] = self.most_common_words(split_book, 10)
        book_data['ПроцентУникальныхСлов'] = (book_data['КоличествоУникальныхСлов'] / book_data['КоличествоСловТочно']) * 100
        book_data['КоличествоУникальныхСловНа3000'] = self.unique_words_in(split_book, 3000)
        book_data['ПроцентУникальныхСловНа3000'] = (book_data['КоличествоУникальныхСловНа3000'] / 3000) * 100
        book_data['КоличествоУникальныхСловНа10000'] = self.unique_words_in(split_book, 10000)
        book_data['ПроцентУникальныхСловНа10000'] = (book_data['КоличествоУникальныхСловНа10000'] / 10000) * 100
        book_data['КоличествоПредлолжений'] = self.count_centencions(book)[0]
        book_data['СредняяДлинаПредлолжения'] = self.count_centencions(book)[1]
        book_data['ПроцентСуществительных'] = (book_data['КоличествоЧастейРечи']['Существительные'] / book_data['КоличествоСловТочно']) * 100
        book_data['ПроцентПрилагательных'] = ((book_data['КоличествоЧастейРечи']['ПрилагательныеПолные'] +
                                              book_data['КоличествоЧастейРечи']['ПрилагательныеКраткие']) / book_data['КоличествоСловТочно']) * 100
        book_data['ПроцентГлаголов'] = ((book_data['КоличествоЧастейРечи']['ГлаголыЛичнаяФорма'] +
                                              book_data['КоличествоЧастейРечи']['ГлаголыИнфинитив']) / book_data['КоличествоСловТочно']) * 100
        book_data['ПроцентОстальныхЧастейРечи'] = 100 - (book_data['ПроцентСуществительных'] + book_data['ПроцентПрилагательных'] + book_data['ПроцентГлаголов'])
        book_data['КоличествоТочек'] = self.sequense(book, '.')
        book_data['КоличествоЗапятых'] = self.sequense(book, ',')
        book_data['КоличествоВосклицательныхЗнаков'] = self.sequense(book, '!')
        book_data['КоличествоВопросительныхЗнаков'] = self.sequense(book, '?')
        book_data['КоличествоДлинныхТире'] = self.sequense(book, '—')
        book_data['КоличествоТире'] = self.sequense(book, '–')
        book_data['КоличествоДефисов'] = self.sequense(book, '-')
        book_data['ПроцентДиалогов'] = self.brut_dialog_procents(book)
        book_data['АЗС3000'] = self.dinamical_change_asz(split_book, 3000)
        self.books_info.append(book_data)
        return self.books_info

    def books_analyze(self):
        """
        Функция, которая анализирует все книги, которые лежат в одной папке с ней, и результат сохраняет в JSON-файл.
        :return: books_info
        """
        dir = os.listdir(os.getcwd())
        for file in dir:
            if file.endswith('.docx'):
                print('Анализируется: {}'.format(file))
                self.analyze(file)
        with open('books_information.json', 'w', encoding='utf-8') as f:
            json.dump(self.books_info, f, ensure_ascii=False, indent=2)
        return self.books_info

if __name__ == '__main__':
    analiser = Books_analiser()
    analiser.books_analyze()
