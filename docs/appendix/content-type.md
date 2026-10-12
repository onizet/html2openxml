---
lang: en-US
title: Content Type MIME Header
description: Advanced Usage and Deployment
---

# MIME Header

When deploying the DOCX file via web server (ASP.NET, IIS, etc.), you must serve the correct MIME type to prevent browsers from defaulting to an incompatible legacy format (e.g., .doc).

`application/vnd.openxmlformats-officedocument.wordprocessingml.document.`
