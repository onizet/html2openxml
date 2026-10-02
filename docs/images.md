---
lang: en-US
title: Image and Asset Processing Guide
description: Guide for image processing.
---

This guide details how `HtmlToOpenXml` converts image references into native OpenXML resources within your target Word document. The library is designed to handle everything from inline Base64 strings to complex network downloads and dynamic format conversions.

## Quick Decision Guide

Before diving into the conversion code, determine your source type.

| Scenario | Goal | Implementation Detail |
| :--- | :--- | :--- |
| Default Fetch | Simple conversion; images are inline or on a network. | Use `DefaultWebRequest`. This handles Base64 data URIs, HTTP(S), and file:/// protocols out of the box. |
| Authenticated Fetch | Retrieving remote content requiring credentials. | Instantiate `DefaultWebRequest(myHttpClient)` and configure the client object to handle proxies, API keys, or JWT tokens. |
| Local Mapping | Retrieving content from a known local path structure. | Set the `BaseImageUrl` property on the `DefaultWebRequest` to map relative paths (see code example below). |
| Custom Protocol | Retrieving resources from non-standard sources (DB, API). | Implement your custom class against the `IWebRequest` interface to manage unique retrieval protocols. |

* **Inline (Data URI)**: → Use Base64 string. Fastest, zero network dependencies, reliable for simple assets.
* **Relative Path**: → Use `BaseImageUrl` on the converter. Simplest network retrieval.
* **Absolute URI**
* **Custom Format or Protocol**: → Requires implementing IWebRequest. Most robust, highest complexity.

## Image Processing Modes

The library's behavior regarding downloaded images is controlled by the `ImageProcessing` property on the `HtmlConverter`. This setting dictates whether a downloaded image is bundled into the document or merely linked to an external source.

| ImageProcessingMode | Behavior | Use Case |
| :--- | :--- | :--- |
| `Embed` (default) | Downloads and embeds all referenced images into the package. Creates a self-contained, offline snapshot of the document. | Most use cases; ensures portability and reliability. |
| `LinkExternal` | Downloads the image into the document but maintains an external link. The image is not bundled, making the final document size small, but it requires network access to view or print. | Internal Corporate documents. |
| `EmbedDataUriOnly` | Skips all external network requests. Only inline Base64/Data URI images are processed and included. | Security or testing, or when you *only* want to use inline assets. |

## Image Source and Retrieval Strategies

The library supports three ways to provide image data: inline, relative path (local server), or absolute URI (remote server).

The library natively supports the following formats within the OpenXML standard: `bmp`, `emf`, `gif`, `ico`, `jp2`, `jpe`, `jpeg`, `pcx`, `png`, `svg`, `tif`, and `tiff`. If your source image format is outside of this list (e.g., `.webp`), it must be converted to a supported format before or during the library's download process (see sample below).

### 1. Inline Assets (Base64 Data URI)

For maximum performance and zero external dependency, images can be included directly in the HTML using a Data URI scheme.

```html
<!-- Example: Base64 encoded PNG data -->
<img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUA..." alt="Red dot" />
```

### 2. External Downloads (Network/Disk)

When the image path is external (`<img>` tag with `src="http://..."` or `/images/pic.gif`), the conversion process delegates retrieval to an `IWebRequest` implementation.

#### Standard Approach

`DefaultWebRequest` provides native support for all standard protocols: `http`, `https`, and local file system paths (`file://`).

**Setting a Base URL:**

If your HTML uses relative paths (e.g., `/_layouts/images/pic.gif`), you must provide a `BaseImageUrl` to resolve the paths correctly:

```csharp
// The converter will combine this base URL with the relative path.
HtmlConverter converter = new(mainPart, new DefaultWebRequest() {
    BaseImageUrl = "http://myserver:8080/"
}); 
```

#### Customizing Downloads with `IWebRequest`

If you require specific behavior—such as authentication headers, proxy usage, or monitoring the download status—you must implement `HtmlToOpenXml.IO.IWebRequest` and pass it to the converter constructor:

```csharp
// Dependency Injection of a custom downloader is the most flexible approach.
HtmlConverter converter = new(mainPart, myCustomDownloader); 
```

### Logging

For monitoring or advanced troubleshooting of I/O access during downloads, you may provide a custom `ILogger` implementation to the `DefaultWebRequest` constructor. This enables you to log connection status and resource retrieval attempts without modifying the core conversion logic:

```csharp
// Pass the logger for detailed I/O troubleshooting. The null indicates you are not providing a custom HttpClient.
HtmlConverter converter = new(mainPart, new DefaultWebRequest(null, myLoggerInstance)); 
```

## Advanced Image Processing & Customization

The power of the library lies in its ability to intervene during the conversion pipeline.
Use these advanced features when you need more than a simple download.

### Dynamic Format Transformation (WebP Example)

If the referenced image format is not natively supported by OpenXML (`.webp`, for example), you can use a custom `IWebRequest` implementation to download the file and convert its format into a supported stream (`png`, `jpg`) before passing it to the converter.

This is a perfect use case for inheriting from `DefaultWebRequest`:

```csharp
class WebPWebRequest : DefaultWebRequest
{
    protected override async Task<Resource> DownloadHttpFile(Uri requestUri, CancellationToken cancellationToken)
    {
        var resource = await base.DownloadHttpFile(requestUri, cancellationToken);

        if (resource.Content is not null && requestUri.OriginalString.EndsWith(".webp"))
        {
            // Load, convert to PNG stream, and replace the original resource content.
            using var img = await Image.LoadAsync(resource.Content, cancellationToken); 
            var convertedStream = new MemoryStream();
            await img.SaveAsPngAsync(convertedStream, cancellationToken);
            
            // Clean up original resource and set the new PNG stream.
            resource.Content.Dispose();
            resource.Content = convertedStream; 
        }
        return resource;
    }
}

// Usage: Pass the custom retriever to the converter.
HtmlConverter converter = new HtmlConverter(mainPart, new WebPWebRequest());
await converter.ParseBody(@"<img src='https://www.gstatic.com/webp/gallery/1.webp'>");
```

### Metadata and Sizing Controls

* **`alt` attribute** The value provided in the alt attribute serves as a tooltip on hover and is crucial for screen readers, fulfilling accessibility requirements.

* **Dimensions:** Specify `width` and `height` attributes in pixels (`px`) or percentages (`%`). The converter respects the aspect ratio if only one attribute is provided. If no dimension attributes are present, the library attempts auto-detection.

* **Captions/Legends:** Use the `<figure>` and `<figcaption>` HTML tags to include a legend that appears above or below the image in Word.

* **SVG Detail:** When using `<svg>`, placing descriptions inside the `<title>` or `<desc>` tags ensures these metadata fields are accurately reflected in the resulting Word document properties.

## Styling and Layout Support

The library interprets several HTML attributes that control the visual rendering in Word:

* **Borders:** Supported via `border`, `border-style`, `border-width`, and `border-color`.

  * `border-style`: Supports `dotted`, `dashed`, `solid`, `double`, `inset`, `outset`, or `none`.
  * `border-width`: Supports `px` units, or keywords like `medium`, `thick`, or `thin`.

* **Scaling:** You can command an image to extend past the page boundaries using `width:100%` on a constrained layout.
