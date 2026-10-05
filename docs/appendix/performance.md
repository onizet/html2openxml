---
lang: en-US
title: Performance
description: Performance and reliability
llmstxt: false
---

# Performance and reliability

Recent versions include several internal improvements designed for large-scale document generation:

* HTML parsing has been rewritten to use `Span<char>` in critical code paths, reducing allocations and
improving parsing throughput by approximatively 25%.

* All remaining regular expressions are executed with explicit timeouts to protect against
catastrophic backtracking and potential denial-of-service scenarios when processing untrusted input.

* The overall parsing is done via [AngleSharp](https://www.nuget.org/packages/AngleSharp), and processed in accordance with the **Interpreter** design pattern.
