---
title: "Как прочитать абзац Word, прежде чем просить локальную модель процитировать его"
date: 2026-11-22T08:45:00+03:00
author: "Ranis Galiaskarov"
description: "На примере трёх абзацев и обработчика на C# разбираем, почему выделение жирным ломает поиск слова и как сохранить номер абзаца."
tags: ["dotnet", "csharp", "openxml", "local-ai"]
draft: false
---

Коллега спрашивает у автономного помощника дату поставки из документа Word. Слово «delivery» видно на странице, а поиск по отдельным текстовым узлам XML даёт ноль совпадений. Чтобы повторить опыт, достаточно выделить жирным несколько букв в середине слова. Для человека слово остаётся целым, а в файле оно уже разбито на части.

## Возьмите документ из трёх абзацев

Скачайте [учебный файл Word](delivery.docx). В его первом абзаце написано «The delivery date is 18 May.», причём в слове «delivery» выделены только буквы **liv**. Второй абзац пуст. В третьем указано условие оплаты «Payment is due on 30 June.». Обе фразы вымышлены.

DOCX — это ZIP-пакет. Основной текст этого файла лежит в `word/document.xml`. Ниже первый абзац без объявления пространства имён и части оформления.

```xml
<w:p>
  <w:r><w:t>The de</w:t></w:r>
  <w:r><w:rPr><w:b/></w:rPr><w:t>liv</w:t></w:r>
  <w:r><w:t>ery date is 18 May.</w:t></w:r>
</w:p>
```

Слова «delivery» нет ни в одном отдельном `w:t`. Если поставить пробел на каждой границе форматирования, получится «de liv ery». Чтобы сохранить цитату, соедините текстовые узлы **одного абзаца** подряд. Номер абзаца увеличивайте и для пустых строк. Тогда срок оплаты получит номер 3, и читатель найдёт его в исходнике. [Устройство документа WordprocessingML в документации Microsoft](https://learn.microsoft.com/en-us/office/open-xml/word/structure-of-a-wordprocessingml-document)

## Прочитайте абзацы на C#

Создайте консольный проект и замените содержимое его `Program.cs` кодом ниже. Скачайте учебный файл в папку проекта. При установленном .NET SDK выполните `dotnet run -- delivery.docx`. Для чтения ZIP и XML используются встроенные библиотеки.

```csharp
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

class ReadParagraphs
{
    static void Main(string[] args)
    {
        if (args.Length != 1) throw new ArgumentException("Pass one DOCX file.");
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        using (var zip = ZipFile.OpenRead(args[0]))
        {
            var part = zip.GetEntry("word/document.xml");
            if (part == null) throw new InvalidDataException("Missing main document.");
            var settings = new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = 24000000
            };
            using (var stream = part.Open())
            using (var reader = XmlReader.Create(stream, settings))
            {
                var document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
                int paragraph = 0;
                foreach (var p in document.Descendants(w + "p"))
                {
                    paragraph++;
                    var text = string.Concat(p.Descendants(w + "t").Select(t => t.Value));
                    if (!string.IsNullOrWhiteSpace(text))
                        Console.WriteLine("{0} | {1}", paragraph, text);
                }
            }
        }
    }
}
```

## Проверьте, откуда помощник взял ответ

Ниже результат консольного примера. Этот же DOCX мы передали обработчику из Windows-выпуска OSTUM Desktop 5.7.0. Он вернул те же две цитаты с метками `абз. 1` и `абз. 3`.

```text
1 | The delivery date is 18 May.
3 | Payment is due on 30 June.
```

Теперь спросите помощника: «Когда ожидается поставка? Приведи точную фразу на языке оригинала и номер абзаца». Дату, цитату и ссылку на абзац можно оценить по отдельности. Если слово исчезло при чтении файла, этот опыт покажет, на каком этапе искать ошибку.

Пример обрабатывает обычный текст абзацев основной части документа. Колонтитулы, сноски, явные табуляции и переносы строк требуют отдельных правил. Номер абзаца относится к конкретной версии документа. Если во второй редакции текст переместили, импортируйте её заново. [Обработчик DOCX в исходниках соответствующей версии OSTUM](https://github.com/Ranis2023/ostum-desktop/blob/v5.7.0/src/ResearchDocuments.cs)

Я разрабатываю OSTUM Desktop. Консольный пример и обработчик программы проверены на приложенном файле. Текст подготовлен с помощью ИИ.
