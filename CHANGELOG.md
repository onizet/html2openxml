# Changelog

## 3.0.0 (Next major release)

- rely on HtmlAgilityPack for the parsing

## 2.4.0

- fix caching the provisioned images
- drop support for .Net Standard 1.3

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
