---
lang: en-US
title: Placeholder Replacement
description: Guide to use Parse():IList<OpenXmlCompositeElement> in conjunction with placeholder
llmstxt: false
---

## Replace placeholder with generated output

This guide will demonstrate the recommended and safest way to replace some content located in.

The recommend and safe way relies on **Content Control** (`SdtBlock`).

> **💡 Pro Tip:** Whilst it seems easier to write raw text `{{ MyPlaceholder }}` in a Word document body, you should not trust the underlying persisted format. MS Word can break this text in multiple `run` chunks, leaving your code searching for a keyword that will remain unmatched.

## 1. Create the Placeholder in Word

Open MS Word. Ensure you have enabled the `Developer` tab from the Ribbon.

In the `Developer` tab, click on the **rich text content** (the `Aa` icon).

[Create a new StdBlock](images/sdtcontent-ribbon.png)

Name your content control. Click on the **Properties** tab from the ribbon.
In the dialog box that appears, enter `MyPlaceholder` in the **Tag** field.

[Assign a tag to a SdtBlock](images/sdtcontent-tag.png)

### Enable the Developer tab

Follow the instructions on [Microsoft Support](https://support.microsoft.com/en-us/word/show-the-developer-tab-in-word) or execute the following steps:

On the **File** tab, go to **Options** > **Customize Ribbon**.
Under **Customize the Ribbon** and under **Main Tabs**, select the ``Developer`` check box.

## 2. Use HtmlConverter with an existing template

```csharp
await using var generatedDocument = new MemoryStream();
using var templateBuffer = ResourceHelper.GetStream("Resources.template.docx");
await templateBuffer.CopyToAsync(generatedDocument);
using var package = WordprocessingDocument.Open(generatedDocument, true);
HtmlConverter converter = new(package.MainDocumentPart!);
await converter.ParseBody(html);
```

## 3. Use Parse() / ParseAsync()

```csharp
var std = mainPart.Document!.Body!.Descendants<SdtBlock>()
   .FirstOrDefault(b => b.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value == "MyPlaceholder");
if (std is null)
    throw new InvalidOperationException("Placeholder missing from template");

HtmlConverter converter = new(mainPart);
var elements = await converter.ParseAsync(html);

foreach (var el in elements)
    std.InsertAfterSelf(el);
std.Remove();
```
