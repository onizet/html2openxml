---
lang: en-US
title: Page Layout
description: Guide for advanced page layout.
---

To manage the visual flow and chapter structure of your generated Word document, use these CSS-like page layout attributes to control breaks and orientation.

## Page Break

Page breaks are controlled via the `page-break-before` or `page-break-after` attributes, which accept the value **always**.
These directional controls apply to various block elements (`p`, `div`, `pre`, `span`) or the document body.

```html
<div style="page-break-after:always">
    Lorem ipsum dolor sit amet,
    consectetur adipiscing elit.
</div>

Duis dictum leo quis ipsum tempor nec ultrices sapien elementum.
```

## Page Orientation

Since page orientation is not a standard HTML property, the library uses the non-standard `page-orientation` attribute.
Currently, this controls the orientation of the entire document body (`body`).
Future versions will expand support to control individual page sections via the `div` tag

```html
<body style="page-orientation: landscape">
    Lorem ipsum dolor sit amet,
    consectetur adipiscing elit.
</body>
```
