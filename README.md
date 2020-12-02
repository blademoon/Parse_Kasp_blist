# Parse_Kasp_blist
Небольшой скрипт для парсинга черных списков "Kaspersky Security для Microsoft Exchange Servers".

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
  <img width="500" height="313" src="https://github.com/blademoon/Coursera_Mathematics_and_Python_for_Data_Analysis/blob/main/Week%201/Lecture%201-5/Python%20operating%20modes.png">
</p>
*Пример окна Microsoft Excel с результатами работы скрипта*

***Обязательно сохраните результаты обработки переданный в Microsft Excel в файл. Он не сохраняется автоматически.***


