---
lang: en-US
title: Restrict Editing
description: Common post-generation operation
llmstxt: false
---

While `HtmlToOpenXml` handles the conversion from HTML to OpenXml, advanced controls like locking or password protection are features of the underlying DocumentFormat.OpenXml library itself.
These properties allow you to set metadata or protection flags on the final .docx package object.

## Restrict Editing

If you ship a document with only some Forms input control, you can prevent unwanted accidental edits.
User's input will only be accepted in the forms controls.

```csharp
private static void RestrictDocumentEdition(WordprocessingDocument document, string password)
{
    var settingsPart = document.MainDocumentPart!.DocumentSettingsPart
        ?? document.MainDocumentPart.AddNewPart<DocumentSettingsPart>();

    settingsPart.Settings ??= new();

    var protection = settingsPart.Settings.GetFirstChild<DocumentProtection>()
        ?? settingsPart.Settings.AppendChild(new DocumentProtection());

    protection.Edit = DocumentProtectionValues.Forms;
    protection.Enforcement = OnOffValue.FromBoolean(true);

    // you can enforce the protection with a password
    const uint spinCount = 10_000;
    var salt = RandomNumberGenerator.GetBytes(16);
    var hash = SHA1.HashData(salt.Concat(Encoding.UTF8.GetBytes(password)).ToArray());
    for (int i = 0; i < spinCount; i++)
    {
        hash = SHA1.HashData(hash);
    }
    protection.CryptographicProviderType = CryptProviderValues.RsaFull;
    protection.CryptographicAlgorithmClass = CryptAlgorithmClassValues.Hash;
    protection.CryptographicAlgorithmType = CryptAlgorithmValues.TypeAny;
    protection.CryptographicAlgorithmSid = 4; // SHA-1
    protection.CryptographicSpinCount = spinCount;
    protection.Salt = Convert.ToBase64String(salt);
    protection.Hash = Convert.ToBase64String(hash);
    settingsPart.Settings.Save();
}
```
