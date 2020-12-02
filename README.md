# Parse_Kasp_blist (дальнейшая разработка прекращена).
Небольшой скрипт для парсинга черных списков "Kaspersky Security для Microsoft Exchange Servers" из файлов формата blist в Microsoft Excel. **Обрабатывает все файлы найденные в указанной папке.**

# Минимальные требования
1. Наличие установленного Microsoft Excel на ПК, на котором происходит обработка черных списков. Результатом работы скрипта будет открытое окно Microsoft Excel c обработанной информацией.
2. Powershell 5.1 (тестирование на более поздних версиях не проводилось, на более ранних - не заработает).

# Как использовать
1. Выгрузите черный список (в формате blist). Дополнительную информацию о том, как это сделать можно почитать [здесь](https://support.kaspersky.com/KS4Exchange/9.6/ru-RU/127325.htm).
2. Скопируйте на свой ПК репозиторий с скриптом. Для работы сркипта минимально необходимы следующие файлы:
- Parse_Kasp_blist.ps1 (сам скрипт).
- Black_list.xlsx (Excel шаблон результирующего файла).
2. Запустите скрипт Parse_Kasp_blist.ps1.
3. На запрос скрипта "Enter the path to the folder containing the files with the extension * .blist:" введите полный путь к каталогу содержащиму файлы с расширением "blist" и нажмите Enter. Далее запустится процесс обьработки файлов.
4. Итоговый результат обработки вы увидите в отдельном окне Microsoft Excel.

<p align="center">
  <img width="800" height="474" src="https://github.com/blademoon/Parse_Kasp_blist/blob/master/img/Result_window.png">
</p>
<p align="center">Пример окна Microsoft Excel с результатами работы скрипта</p>

5. ***Обязательно сохраните результаты работы скрипта (переданные в окно Microsft Excel) в файл. Файл не сохраняется автоматически.***

# Дополнительная информация.
В папке "BLACKLIST_EXAMPLE" находится несколько файлов-примеров черных списков. Данные файлы взяты с реального, работающего в продакшене сервера и подвергнуты [анонимизации](https://ru.wikipedia.org/wiki/%D0%90%D0%BD%D0%BE%D0%BD%D0%B8%D0%BC%D0%B8%D0%B7%D0%B0%D1%86%D0%B8%D1%8F).

P.S. Всем отличного дня. 
<p align="center">
  <img width="320" height="357" src="https://github.com/blademoon/Parse_Kasp_blist/blob/master/img/black_cat.png">
</p>
 
