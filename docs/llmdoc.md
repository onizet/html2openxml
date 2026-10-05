---
lang: en-US
title: LLM Friendly Docs
description: Selecting the correct document context optimizes LLM performance.
llmstxt: false
---

# AI Pairing Guide

This repository offers two views of the library: the **Technical Contract** and the **Usage Guide**. Choosing the right context ensures efficient LLM processing.

## Usage Guide: [llms-full.txt](llms-full.txt)

**[Primary Context for Most Tasks]**

This is the curated, human-readable instruction manual. It translates code concepts into usage scenarios. Include this when:

* You need a conceptual walkthrough (e.g., "How do I handle footnotes?").
* You need runnable code examples and best practices.
* You are performing onboarding or high-level design discussions.

## Technical Contract: [nuspec-ai.txt](nuspec-ai.txt)

**[Deep Dive / Debugging Context]**

This library leverages [NuSpec.AI](https://www.nuget.org/packages/NuSpec.AI) to structure the API documentation, including property constraints and enumeration values. Include this when:

* You need to debug a subtle rendering issue (e.g., "Why is the table not spanning 2 columns?").
* You need to validate low-level constraints (e.g., "Does the library support `colspan`?").
* The task requires precise, implementation-level detail.

---

> **💡 Pro Tip:** For maximum context fidelity, include both documents; the Wiki handles the "How," and the NuSpec handles the "Why" it works that way.
