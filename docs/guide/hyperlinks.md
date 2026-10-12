---
lang: en-US
title: Hyperlinks and Internal Navigation
description: Guide to linking external resources and navigating within the generated document using anchors and bookmarks.
---

## External Hyperlinks

The parser supports converting HTML links to valid targets in the generated DOCX file. Only fully qualified or relative URI schemes are processed.
The converter renders all content inside the anchor tags faithfully. However, if the `href` URI is invalid or unresolvable, the text will appear, but no functional hyperlink object will be created.

**Supported Schemes:**

* `http://www.site.com` (Absolute URL)
* `file:///../reference.docx` (Local URL)
* `://www.site.com` (`http` scheme will be used)
* `www.site.com` (`http` scheme will be used)

**Ignored Schemes:**

* `javascript:void(0)` or similar client-side scripts.

## Internal Anchors and Bookmarks

To create navigable links within the document (e.g., Table of Contents or cross-referencing), you must define targets using either explicit IDs (which the library maps to internal anchors) or the `data-bookmark` attribute.

By default, **Anchor Link Support is enabled**. To disable this feature globally:

```csharp
converter.SupportsAnchorLinks = false; 
```

### Targeting Specific Anchors

To allow a link to point directly to an element within the document, apply a unique `id` attribute. The library handles the low-level conversion to a searchable document anchor.

```html
<!-- Defining an Anchor Target -->
<h1 id="chapter-one">Chapter 1: Introduction</h1>

<!-- Linking to the Target -->
<p>Read more in our next chapter on <a href="#chapter-one">Chapter 1</a>.</p>
```

### Explicit Bookmarking (`data-bookmark`)

When appending content to an existing Word document or template, use `data-bookmark` to explicitly target structural elements that reside in the destination DOCX but are not part of your current input HTML.
This functionality is critical when working with master templates or merging multiple conversion outputs.

```html
<!-- Using data-bookmark for any content type (H1, P, DIV, etc.) -->
<h1 data-bookmark="my_topic">Detailed Analysis</h1>
```

### Built-In Document Targets

The following targets are automatically recognized and supported by the library:

* **`#_top`**: Links to the beginning of the document body. It is always respected and serves as a universal fallback target, even when `SupportsAnchorLinks` is set to `false`.
* **`#top`**: An alias for `#_top`, provided for user convenience.

## Embed an image inside an anchor

Image can be clickable and redirecting to a link or anchor.

```html
<a href="http://www.wikipedia.org">
   <img alt="Wikipedia, the Free Encyclopedia" src="wikipedia.png" />
</a>
```
