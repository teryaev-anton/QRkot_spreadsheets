# QRkot_spreadseets
Добавлено формирование отчета в гугл-таблице

Учебный проект:
## Приложение для Благотворительного фонда поддержки котиков **QRKot**

В проекте используется база данных SQLite



## Инструкция по установке и запуску проекта:
1. Клонировать репозиторий:

```
git clone https://github.com/teryaev-anton/cat_charity_fund.git
```

2. Создать и активировать виртуальное окружение:

```
python -m venv venv

source venv/scripts/activate
```

3. Обновить PIP и установить зависимости из файла requirements.txt:

```
python3 -m pip install --upgrade pip
pip install -r requirements.txt
```

4. В корневой папке cоздать и заполнить файл `.env` (образец в файле `.env_example`)

5. Создать базу данных и выполнить миграции:
```
alembic upgrade head
```

6. Запустить проект командой:
```
uvicorn app.main:app
```

После запуска проекта документация swager будет доступна по адресу http://127.0.0.1:8000/docs/


## Проект подготовил Антон Теряев