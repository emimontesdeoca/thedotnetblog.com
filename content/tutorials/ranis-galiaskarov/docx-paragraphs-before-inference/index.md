---
title: "Read Word paragraphs before asking the local model for a quotation"
date: 2026-11-22T08:45:00+03:00
author: "Ranis Galiaskarov"
description: "A three-paragraph DOCX and a C# reader reveal how a bold fragment can break a word search and how to keep the original paragraph number."
tags: ["dotnet", "csharp", "openxml", "local-ai"]
draft: false
---

A colleague asks an offline assistant for the delivery date in a Word document. The search term is present on the page. Yet a search over individual XML text nodes finds zero matches. You can reproduce this case by making the middle of a word bold. The visible word stays intact while its representation changes.

## Try it with one short document

Download the [synthetic Word example](delivery.docx). The first paragraph reads “The delivery date is 18 May.” Only the letters **liv** in “delivery” are bold. The second paragraph is empty. In the third, “Payment is due on 30 June.” The document contains no personal data.

A DOCX file is a ZIP package. The example stores its main text in `word/document.xml`. This is the first paragraph with the namespace and formatting details omitted:

```xml
<w:p>
  <w:r><w:t>The de</w:t></w:r>
  <w:r><w:rPr><w:b/></w:rPr><w:t>liv</w:t></w:r>
  <w:r><w:t>ery date is 18 May.</w:t></w:r>
</w:p>
```

Searching one `w:t` at a time misses “delivery”. Adding a space at every formatting boundary produces “de liv ery”. Concatenating the text **within the paragraph** reconstructs the visible word. Advance the paragraph number before skipping an empty paragraph so the following reference still points to paragraph 3. [Microsoft's introduction to WordprocessingML](https://learn.microsoft.com/en-us/office/open-xml/word/structure-of-a-wordprocessingml-document)

## Read the paragraphs in C#

Place the following code in a console project, replacing its sample `Program.cs`. Download the fixture to the project directory and run `dotnet run -- delivery.docx` using a .NET SDK. The code uses the framework ZIP and XML libraries.

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

## Give the assistant evidence you can check

The console example printed the two lines below. We also passed the same DOCX to the document reader from the Windows release of OSTUM Desktop 5.7.0. That reader returned the same quotations with its Russian paragraph labels, `абз. 1` and `абз. 3`.

```text
1 | The delivery date is 18 May.
3 | Payment is due on 30 June.
```

Now ask your chosen assistant “When is delivery due? Give the exact sentence and its paragraph number.” A test can check the date, quotation and source number separately. If extraction fails, you can fix the reader before spending time changing the model or its prompt.

The example covers ordinary paragraph text in the main part. Headers, footnotes, explicit tabs and line breaks require their own handling. In this example the paragraph number is a position in the imported document, so moving the text in a revised copy calls for importing it again. [Reader in the tagged OSTUM source](https://github.com/Ranis2023/ostum-desktop/blob/v5.7.0/src/ResearchDocuments.cs)

I maintain OSTUM Desktop. Both the console example and the release reader were run against the attached fixture. This article was prepared with AI assistance.
