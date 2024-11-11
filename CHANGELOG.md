# Changelog

## 3.2.2

- Supports a feature to disable heading numbering #175
- Support center image with margin auto #171
- Support deprecrated align attribute for block #171
- Fix parsing of style attribute with a key with no value
- Improve parsing of style attribute to avoid an extra call to HtmlDecode
- Extend support of nested list for non-W3C compliant html #173

## 3.2.1

- Fix indentation of numbering list #166
- Bordered container must render its content with one bordered frame #168
- Fix serialisation of the "Harvard" style for lower-roman list
- Fix ParseHeader/Footer where input with multiple paragraphs output only the latest
- Ensure to apply default style for paragraphs, to avoid a paragraph between 2 list is mis-guessed

## 3.2.0

- Add new public API to allow parsing into Header and Footer #162. Some API methods as been flagged as obsolete with a clear message of what to use instead.
  This is not a breaking changes as it keep existing behaviour.
- Add support for `SVG` format (either from img src or the SVG node tag)
- Automatically create the `_top` bookmark if needed
- Fix a crash when a hyperlink contains both `img` and `figcation`
- Fix a crash when `li` is empty #161

## 3.1.1

- Fix respecting layout with `div`/`p` ending with line break #158
- Prevent crash when header/footer is incomplete and parsing image #159
- Fix combining 2 runs separated by a break, 2nd line should not be prefixed by a space

## 3.1.0

- Fix table Cell borders are wrongly applied on the run #156
- Correctly handle RTL layout for text, list, table and document scope #86 #66
- Support property line-height #52
- Fallback to `background` style attribute as many users use this simplified attribute version
- In `HtmlDomExpression.CreateFromHtmlNode`, use the correct casting to `IElement` rather than `IHtmlElement`, to prevent crash if `svg` node is encountered

## 3.0.1

- Ensure to count existing images from header and footer too #113
- Preserve line break pre for OSX/Windows
- Prevent a crash when the provided style is missing its type
- Defensive code to avoid 2 rowSpan+colSpan with a cell in between to crash #59

## 3.0.0

- AngleSharp is now the backend parser for Html
- Refactoring to use the Interpreter/Composite design pattern, which ease the code maintenance
- Lots of new unit test cases (190+)
- Rewriting of `list` (correct handling of nested style, restarting numbers and consecutive)
- Rewriting of `table` (row span, col span, col tags driving styles)
- Parallel download of images at early stage of the parsing.

## 2.4.2

- Fix signing the assembly
- Enable Nullable reference types
- support latest version of OpenXML SDK (3.1.0) which introduces breaking changes, but also support embedding SVG and JPEG2000 files.
- fix caching the provisioned images
- drop support for .Net Standard 1.3

## 2.4.0 and 2.4.1

do not use as the signing assembly was in failure #138

## 2.3.0

- better table border style
- keep processing html even if downloading image generates an error
- support for styling OL, UL and LI elements

## 2.2.0

- support latest version of OpenXML SDK (2.12.0) which introduces an API to add an OpenXmlElement to the correct XSD order
- restore support for .NET 4.6+, Net Standard 1.3+
- use cleaner name for base-64 images description

## 2.1.0

- support latest version of OpenXML SDK (2.11.0+) which fix fatal issue
- drop support for .NET 4.0, .Net Standard 1.4

## 2.0.3

- optimize number of nested list numbering (thanks to BenGraf)
- fix an issue where some styles weren't being applied
- fix reading JPEG images with SOF2 progressive DCT encoding

## 2.0.2

- fix nested list numbering

## 2.0.1

- fix manual provisioning of images
- img respect both border attribute and border style attribute

## 2.0.0

This brings .Net Core support:

- better inline styling
- numbering list with nested list is more stable
- allow parsing unit with decimals
- color can be either rgb(a), hsl(a), hex or named color.
- parser is more stable

## Pre 1.6.0

- imported from codeplex.com
