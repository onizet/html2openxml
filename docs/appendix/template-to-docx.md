---
lang: en-US
title: Convert .dotx to .docx
description: Convert .dotx to .docx
llmstxt: false
---

# The Conversion Process

While renaming a file extension (`.dotx` to `.docx`) is insufficient, this section details the programmatic process of converting a template file into a fully structured, ready-to-use DOCX document.

The conversion involves two critical steps: changing the file type and establishing the template relationship.

1. **Change Document Type:** You must explicitly inform the parser that the document is being converted into a final, executable document format.

2. **Establish Relationship:** You add an external relationship referencing the template's original location, which helps maintain document integrity.

## Code sample

```csharp
using (WordprocessingDocument template = WordprocessingDocument.CreateFromTemplate(documentStream))
{
    var converter = new HtmlConverter(template);
    ...
}
```
