---
lang: en-US
title: Numbering List
description: Simple and advanced list support
---

The library supports deeply nested lists, translating up to the physical limit of MS Word's outlining structure at 8 levels.
Lists can be defined using `<ol>` or `<ul>` tags.

## Customise List Styles

The visual style is controlled via the `list-style-type` CSS attribute or the `type` attribute (`<ol type="1|a|A|i|I">`).

Supported values for `list-style-type` include:

* `decimal`, `disc`, `square`, `circle`
* `dash` (non standard, used for convenience)
* `upper-alpha` / `lower-alpha`, `upper-roman` / `lower-roman`, `upper-greek` / `lower-greek`.
* Custom values (e.g., `list-style-type: '+'`)

Supported values for `type` attribute:

* `1`: Decimal Numbers (1, 2, 3...)
* `a`: Lowercase Alphabetical (`a`, `b`, `c`...)
* `A`: Uppercase Alphabetical (`A`, `B`, `C`...)
* `i`: Lowercase Roman Numerals (`i`, `ii`, `iii`...)
* `I`: Uppercase Roman Numerals (`I`, `II`, `III`...)

## Controlling Numbering Flow

By default, the converter assumes continuity between sequential lists. Use these properties to define or override the flow:

- **Continuation:** Force the list to reset a brand new sequence (e.g., to treat two separate lists as distinct), set `ContinueNumbering = false`.

- **Initial Value:** Use the `start` attribute on `<ol>` tags to define an initial value for a list:

    ```html
    <!-- Starts at 50, not 1 -->
    <ol start="50">...</ol>
    ```

## Heading Numbering Detection

The library intelligently detects when a list acts as a chapter outline. If your list structure aligns with document hierarchies (e.g., `1.`, `1.1`, `1.2`), the library automatically maps this structure to a dedicated Word list style (`decimal-tiered`) and applies the appropriate heading numbering.

To disable this automatic mapping and rely purely on HTML structure, set:

```csharp
converter.SupportsHeadingNumbering = false;
```

## Nested Tables for Alignment

For complex layouts where a table must appear as part of a list item, use the following pattern. The converter automatically handles the necessary indentation to maintain visual hierarchy.

```html
<ol>
   <li>Main Item</li>
   <li>
        <!-- The table is naturally indented as part of the list item -->
        <table>...</table>
   </li>
</ol>
```
