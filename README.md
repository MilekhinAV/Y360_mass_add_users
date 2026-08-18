# Импорт сотрудников и подразделений в Яндекс 360 из XLSX

Скрипт `import_yandex360_users.py` читает лист `Шаблон`, создаёт недостающую
иерархию подразделений и затем создаёт сотрудников в самых нижних указанных
подразделениях.

Скрипт предназначен только для создания. Если сотрудник с таким логином уже
есть в организации, строка будет пропущена и отмечена в CSV-отчёте. Данные
существующего сотрудника не изменяются.

## Что делает скрипт

- полностью проверяет структуру и значения XLSX до первого изменяющего запроса;
- читает 13 фиксированных колонок `A–M`;
- считает колонки `N+` последовательными уровнями подразделений;
- получает все существующие подразделения с учётом пагинации;
- повторно использует подразделение, если совпали его название и родитель;
- создаёт отсутствующие уровни сверху вниз;
- создаёт сотрудника в последнем подразделении строки;
- добавляет рабочий и мобильный телефоны в `contacts`;
- добавляет `personal_email` как ручной контакт типа `email`;
- соблюдает паузу между API-запросами;
- повторяет запросы после `429`, `500`, `502`, `503` и `504`;
- формирует CSV-отчёт с результатом каждой строки и `x-request-id` ошибки;
- поддерживает безопасный режим `--dry-run` без POST-запросов.

## Требования

- Python 3.10 или новее;
- доступ к `https://api360.yandex.net`;
- OAuth-токен администратора организации;
- пакет `openpyxl` версии 3.1 или новее.

OAuth-приложению нужны права:

- `directory:read_users`;
- `directory:write_users`;
- `directory:read_departments`;
- `directory:write_departments`.

Права на группы и чтение организации этому скрипту не требуются, поскольку он
не вызывает методы групп и организаций.

## Установка на macOS или Linux

Перейдите в каталог с файлами скрипта и выполните:

```bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
```

## Установка на Windows PowerShell

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
```

## Передача OAuth-токена и ID организации

### macOS с zsh

Команды запрашивают значения интерактивно и не сохраняют токен в истории:

```zsh
read -s "OAUTH_TOKEN?OAuth-токен: "; echo
read "ORG_ID?ID организации: "
export OAUTH_TOKEN ORG_ID
```

### Linux с bash

```bash
read -rsp "OAuth-токен: " OAUTH_TOKEN; echo
read -rp "ID организации: " ORG_ID
export OAUTH_TOKEN ORG_ID
```

### Windows PowerShell

```powershell
$secureToken = Read-Host "OAuth-токен" -AsSecureString
$tokenPointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureToken)
$env:OAUTH_TOKEN = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($tokenPointer)
[Runtime.InteropServices.Marshal]::ZeroFreeBSTR($tokenPointer)
$env:ORG_ID = Read-Host "ID организации"
```

Если переменные окружения не заданы, скрипт сам запросит оба значения в
интерактивном терминале. Также поддерживается файл `.env` в UTF-8 с переменными
`OAUTH_TOKEN` и `ORG_ID`.

Не рекомендуется передавать токен через `--token`: значение может попасть в
историю команд и список процессов.

## Проверка без изменений

Сначала обязательно выполните пробный запуск:

```bash
python3 import_yandex360_users.py new_mail_users_import_template.xlsx --dry-run
```

На Windows:

```powershell
py import_yandex360_users.py new_mail_users_import_template.xlsx --dry-run
```

В режиме `--dry-run` скрипт выполняет только GET-запросы: получает существующие
подразделения и логины, но не создаёт объекты. Поэтому OAuth-токен и `ORG_ID`
всё равно обязательны.

## Реальный импорт

После проверки отчёта запустите ту же команду без `--dry-run`:

```bash
python3 import_yandex360_users.py new_mail_users_import_template.xlsx
```

На Windows:

```powershell
py import_yandex360_users.py new_mail_users_import_template.xlsx
```

## Структура XLSX

На листе `Шаблон` строка 1 должна содержать строго такие колонки `A–M`:

1. `last_name` — фамилия, обязательное поле;
2. `first_name` — имя, обязательное поле;
3. `middle_name` — отчество;
4. `login` — логин, обязательное поле;
5. `password` — пароль;
6. `must_change_password` — требование сменить пароль;
7. `position` — должность;
8. `gender` — пол;
9. `birthday` — дата рождения в формате `YYYY-MM-DD`;
10. `language` — язык;
11. `work_phone` — рабочий телефон;
12. `mobile_phone` — мобильный телефон;
13. `personal_email` — дополнительный контактный email.

С колонки `N` начинаются уровни подразделения. Например:

- `N` — `Школы`;
- `O` — `Школа №1`;
- `P` — `Класс 5А`;
- `Q` — `Парта 2`;
- `R` — `Правый стул`.

Пустой уровень допустим только после последнего заполненного уровня. Конструкция
`Школы → пусто → Класс 5А` считается ошибкой. Если все колонки `N+` пустые,
сотрудник создаётся в корневом подразделении с ID `1`.

Одинаковые названия под разными родителями считаются разными подразделениями.
Например, два отдела `Бухгалтерия` в разных филиалах не конфликтуют.

## Правила обработки полей

- `login`, `first_name` и `last_name` должны быть заполнены;
- повтор одного логина внутри XLSX останавливает импорт до изменений;
- существующий в организации логин не создаётся повторно, но не останавливает
  обработку остальных строк;
- пустой `password` не передаётся в API;
- `must_change_password` принимает `true/false`, `1/0`, `да/нет`;
- пустой `must_change_password` не передаётся в API;
- телефоны лучше хранить в Excel как текст, особенно если нужен знак `+`;
- `personal_email` проверяется и передаётся в `contacts` так:

```json
{
  "type": "email",
  "value": "mail@ya.ru",
  "label": "Personal"
}
```

Это контакт в карточке сотрудника, а не алиас почтового ящика. Письма на такой
адрес не начинают поступать в ящик Яндекс 360.

Формат XLSX не имеет внешней текстовой кодировки: файл представляет собой ZIP с
XML внутри, поэтому автоопределение кодировки для него не требуется. `.env`
читается как UTF-8, включая UTF-8 с BOM. CSV-отчёт записывается в UTF-8 с BOM,
чтобы корректно открываться в Excel.

## Итоговый отчёт

Рядом с исходным XLSX создаётся файл:

- `yandex360_dry_run_report_ГГГГММДД_ЧЧММСС.csv` для пробного запуска;
- `yandex360_import_report_ГГГГММДД_ЧЧММСС.csv` для реального импорта.

Основные статусы:

- `planned` — объект будет создан после запуска без `--dry-run`;
- `created` — сотрудник создан;
- `skipped` — логин уже существует;
- `error` — API вернул ошибку или не удалось создать нужное подразделение.

Если API вернул `x-request-id`, он записывается в последнюю колонку отчёта. Этот
идентификатор нужен при обращении в поддержку Яндекс 360.

Скрипт завершает работу с кодом `0`, если ошибок создания сотрудников нет; с
кодом `1`, если отдельные строки завершились ошибкой; с кодом `2`, если не
пройдена проверка файла, параметров или стартовых запросов к API.

## Дополнительные параметры

Показать все параметры:

```bash
python3 import_yandex360_users.py --help
```

Увеличить паузу между запросами до одной секунды:

```bash
python3 import_yandex360_users.py new_mail_users_import_template.xlsx --request-delay 1
```

Записать отчёт в конкретный файл:

```bash
python3 import_yandex360_users.py new_mail_users_import_template.xlsx --report result.csv
```

Включить подробный журнал:

```bash
python3 import_yandex360_users.py new_mail_users_import_template.xlsx --verbose
```

## Используемые методы API

- получение подразделений: `GET /directory/v1/org/{orgId}/departments`;
- создание подразделения: `POST /directory/v1/org/{orgId}/departments`;
- получение сотрудников: `GET /directory/v1/org/{orgId}/users`;
- создание сотрудника: `POST /directory/v1/org/{orgId}/users`.

Официальная документация:

- https://yandex.ru/dev/api360/doc/ru/ref/DepartmentService/DepartmentService_List
- https://yandex.ru/dev/api360/doc/ru/ref/DepartmentService/DepartmentService_Create
- https://yandex.ru/dev/api360/doc/ru/ref/UserService/UserService_List
- https://yandex.ru/dev/api360/doc/ru/ref/UserService/UserService_Create
