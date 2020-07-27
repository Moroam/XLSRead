# XLSRead
Simple and powerfull XLS reader.
Is a wrapper over PhpSpreadsheet.

Main functions:

1. is_xls_file - Static function for checking file type is supported
2. xls_to_array - Static function for reading XLS file in ARRAY
  Parameters:
  $File — the file to read
  $B - is the last column for reading, the default value is calculated
  $sheet - index of the sheet to read, the active sheet is read by default
  $ExplicitColumns - list of column names to read in text format, convenient for reading numbers — disables automatic type conversion

Additional functions:

1. array_delete_col - deletes a column in the associative array by key
2. column_name_by_number - returns the column name by index
3. column_number_by_name - returns the index of the column by name
4. file_by_num - returns information on a single file from the global array $_FILES when loading multiple files

Простой и мощный инфструмент для чтения XLS файлов.
Является обверткой над PhpSpreadsheet.

Основные функции:

1. is_xls_file — статическая функция для проверка поддержки работы с выбранным типом файлов
2. xls_to_array — статическая функция для чтение файла XLS в массив,
	параметры:
		$FILE — файл для чтения
		$B — последняя колонка, по умолчанию значение вычисляется автоматически
		$sheet — индекс листа для чтения, по умолчанию читается активный лист
		$ExplicitColumns — список имен колонок для чтения в формате текста, удобно для чтения чисел — отключает автоматическое преобразование типов

Вспомогательные функции:

1. array_delete_col — удаляет колонку в ассоциативном массиве
2. column_name_by_number — возвращает имя колонки по индексу
3. column_number_by_name — возвращает индекс колонки по имени
4. file_by_num — возвращает информацию по одному файлу из глобального массива $_FILES при загрузке множества файлов
