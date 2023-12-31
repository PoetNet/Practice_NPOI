# 📑 Пример работы с NPOI для создания и загрузки Excel файлов в ASP.NET Core
Этот проект представляет собой пример использования библиотеки NPOI для создания и загрузки Excel файлов в ASP.NET Core. В коде представлен контроллер, который создает Excel файл на основе данных модели и предоставляет его для загрузки.

## 💬 Терминология
[См. статью](https://habr.com/ru/articles/452094/)

***Для начала ознакомимся с терминами, которые могут быть полезными на старте (хоть и не все из них затронуты в этом примере кода):***

**Процесс (Process)** — объект ОС, изолированное адресное пространство, содержит потоки.

**Поток (Thread)** — объект ОС, наименьшая единица выполнения, часть процесса, потоки делят память и другие ресурсы между собой в рамках процесса.

**Многозадачность** — свойство ОС, возможность выполнять несколько процессов одновременно.

**Многоядерность** — свойство процессора, возможность использовать несколько ядер для обработки данных.

**Многопроцессорность** — свойство компьютера, возможность одновременно работать с несколькими процессорами физически.

**Многопоточность** — свойство процесса, возможность распределять обработку данных между несколькими потоками.

**Параллельность** — выполнение нескольких действий физически одновременно в единицу времени.

**Асинхронность** — выполнение операции без ожидания окончания завершения этой обработки, результат же выполнения может быть обработан позднее.

## 🚀 Примеры кода

1.  **Создание заголовков**
```C#
IRow headerRow = sheet.CreateRow(0);
for (int i = 0; i < properties.Length; i++)
{
    ICell cell = headerRow.CreateCell(i);
    cell.SetCellValue(properties[i].Name);
}
```

2.  **Заполнение данными**
```C#
for (int i = 0; i < data.Count; i++)
{
    TableModel record = data[i];
    IRow row = sheet.CreateRow(i + 1);
    
    for (int j = 0; j < properties.Length; j++)
    {
        ICell cell = row.CreateCell(j);
        object value = properties[j].GetValue(record);
        cell.SetCellValue(value?.ToString() ?? string.Empty);
    }
}
```
3.  **Запись в поток**
   ```C#
workbook.Write(stream);
var content = stream.ToArray();
```
4. **Отправка файла для загрузки**
```C#
return File(content, "application/vnd.ms-excel", "История обслуживания.xls");
```

## 📚 Описание
В коде используется библиотека NPOI для работы с Excel файлами.

1. Создаются заголовки и заполняются данными. Для этого происходит создание объектов IRow и ICell из библиотеки NPOI, а затем устанавливается значение ячейки методом SetCellValue.

2. Данные записываются в поток типа MemoryStream, что позволяет создать Excel файл в памяти.

3. Файл отправляется клиенту с использованием метода File контроллера ASP.NET Core.

Эти фрагменты кода отвечают за создание структуры Excel файла и его отправку клиенту для загрузки. Изучив их, можно лучше понять, как NPOI интегрируется в проект и как формируется Excel файл на основе данных модели.

## ▶️ Запуск проекта

Чтобы запустить этот обучающий проект, выполните следующие шаги:

1. Склонируйте репозиторий на свой компьютер.

2. Откройте проект в вашей среде разработки ASP.NET Core (например, Visual Studio).

3. Запустите проект, затем перейдите по URL http://localhost:<номер_порта>/api/downladexcel в браузере, чтобы скачать Excel файл.

## 📚 Учебный материал

Этот пример может служить учебным материалом для начала работы с библиотекой NPOI в контексте ASP.NET Core. Экспериментируйте с кодом, вносите изменения и углубляйтесь в изучении возможностей NPOI. Удачи 🥰
