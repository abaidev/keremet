## Описание
Проект работает на основе .Net Framework 4.7.2 (Целевая рабочая среда) <br/>
Приложение представляет из себя Windows Form которая ищет клиента по его ИНН, и в случае нахождения предлагает
Сохранить его в формате .xlsx <br/>
Шаблон .xlsx файла будет находиться в папке `Template` <br/>
Сохранение происходит в папку `Result`, с указанием ИНН клиента в названии файла. <br/>
В качестве Базы данных используется `PostgreSQL`.<br/>


## Запуск проекта
Перед запуском проекта необходимо создать переменные среды:
	`"DB_HOST"` (localhost или 127.0.0.1, если на локальной машине), 
	`"DB_NAME"`, `"DB_USER"` и `"DB_PASSWORD"` 
и указать их значения для подключения к Серверу базы данных. <br/>

Других действий не требуется. Запуск в обычном режиме (<b>`F5`</b>).

## Дополнительная информация
Больше информации о [`PostgreSQL`](https://www.postgresql.org/).<br/>
Для подключения к серверу PostgreSQL использовался [`Data Provider Npgsql`](https://www.postgresql.org/).<br/>
Для работы с Excel документом был выбран [`IronXL`](https://ironsoftware.com/csharp/excel/).<br/>


