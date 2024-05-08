import flet as ft
import openpyxl
import datetime





def main(page: ft.Page):
    page.title = "Анкета"
    #page.theme_mode = "light"
    #page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.window_width = 630
    page.window_height = 875
    page.window_resizable = False

    #запись ответов в xls файл
    def write_xlsx(e):
        txt2 = [fio.label, born_date.label,job.label,how_become_it.label,but_why.label,hobbies.label,music.label,movies.label,do_you_like_games.label,favorite_games.label,do_you_like_anime.label,favorite_anime.label]


        txt = [fio.value, born_date.value,job.value,how_become_it.value,but_why.value,hobbies.value,music.value,movies.value,do_you_like_games.value,favorite_games.value,do_you_like_anime.value,favorite_anime.value]
        tmn = ft.Text(str(datetime.date.today()))
        #tm = str(datetime.datetime.date())
        book = openpyxl.Workbook()
        sheet = book.active
        sheet["A1"] = "Вопросы:"  # запись в ячейку А1
        sheet["B1"] = "Ответы:"  # запись в ячейку B1
        row = 2  # начать запись со строки 2(для цикла записи в колонку A)
        crow = 2  # начать запись со строки 2(для цикла записи в колонку B)
        for i in txt2:
            sheet[row][0].value = i  # запись в столбец А
            row += 1  # переход на следующую строку
        for j in txt:
            sheet[crow][1].value = j  # запись в стобец B
            crow += 1  # переход на следующую строку
        book.save(tmn.value+"_"+fio.value+".xls")  # сохранение всех ранее записанных изменений в файл
        book.close()  # закрытие файла

        page.snack_bar.open = True
        fio.value = ""
        born_date.value = ""
        job.value = ""
        how_become_it.value = ""
        but_why.value = ""
        hobbies.value = ""
        music.value = ""
        movies.value = ""
        do_you_like_games.value = ""
        favorite_games.value  = ""
        do_you_like_anime.value = ""
        favorite_anime.value = ""
        btn.disabled = True

        page.update()


    #проверка ввода всех полей анкеты
    def validate(e):
        if all([fio.value, born_date.value,job.value,how_become_it.value,but_why.value,hobbies.value,music.value,movies.value,do_you_like_games.value,favorite_games.value,do_you_like_anime.value,favorite_anime.value]):
            btn.disabled= False


        else:
            btn.disabled = True

        page.update()



    fio = ft.TextField(label= "ФИО:", width=600, on_change=validate)
    born_date = ft.TextField(label="Дата рождения:", width=600, on_change=validate,input_filter=ft.InputFilter(allow= True, regex_string= r"[0-9-]",replacement_string=""))
    job = ft.TextField(label="Род занятий:", width=600, on_change=validate)
    how_become_it = ft.TextField(label="Как Вы попали в IT индустрию?:", width=600, on_change=validate)
    but_why = ft.TextField(label="Почему решили заняться именно тем чем занимаетесь сейчас?:", width=600, on_change=validate)
    hobbies = ft.TextField(label="Какие Ваши увлечения?:", width=600,)# on_change=validate)
    music = ft.TextField(label='''Какую музыку предпочитаете слушать?(жанры\группы\композиции):''', width=600, on_change=validate)
    movies = ft.TextField(label='''Какие фильмы предпочитаете смотреть?(жанры\конкретные наименования):''', width=600, on_change=validate)
    do_you_like_games = ft.TextField(label='''Любите ли Вы играть в видеоигры?Если нет то почему?:''', width=600, on_change=validate)
    favorite_games = ft.TextField(label='''Какие Ваши любимые видеоигры?:''', width=600, on_change=validate)
    do_you_like_anime = ft.TextField(label="Смотрите ли Вы аниме? если нет то почему?:", width=600, on_change=validate)
    favorite_anime = ft.TextField(label='''Какое аниме предпочитаете смотреть?:''', width=600, on_change=validate)

    btn = ft.OutlinedButton(text="нажать после заполнения всех полей", width= 600 , disabled= True,on_click= write_xlsx)
    page.snack_bar = ft.SnackBar(content=ft.Text("Результаты анкетирования записаны в файл"), open=False)


    page.add(ft.Row([ft.Column([fio,born_date,job,how_become_it,but_why,hobbies,music,movies,do_you_like_games,favorite_games,do_you_like_anime,favorite_anime,btn])]))



if __name__ == "__main__":
    ft.app(target=main)