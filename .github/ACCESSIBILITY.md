# Accessibility

Excelize is an open source library for reading and writing Microsoft Excel&trade; (XLAM / XLSM / XLSX / XLTM / XLTX) spreadsheets. It is a developer library, not an end-user application, so accessibility for this project has two sides:

1. **The spreadsheets your programs produce.** Can people who use assistive technology read and work with the files you generate?
2. **The project itself.** Can everyone use our documentation, contribute to the code, and take part in discussions?

We care about both, and this document describes what we commit to, what we support, where we fall short today, and how to tell us about barriers.

## Our commitment

- We want developers to be able to build accessible spreadsheets with Excelize, and we want every person, including those who use screen readers, keyboard-only navigation, magnification, or other assistive technology, to be able to learn about, use, and contribute to the project.
- We treat accessibility barriers reported by users as real bugs, not as feature requests that can be ignored.
- We will keep making incremental improvements instead of waiting for a perfect solution. Where we cannot fix something right away, we will say so openly and document a workaround where one exists.
- We welcome contributions that improve accessibility, including documentation fixes, new API options for accessibility-related spreadsheet features, and test coverage.

## Scope

|Area|In scope|
|---|---|
|Library API (`github.com/xuri/excelize/v2`)|Exposing spreadsheet features that make generated files accessible|
|Documentation (API reference, examples, README)|Readable by screen readers, usable by keyboard, clear language|
|GitHub repository (issues, pull requests, discussions)|Templates and contribution guides that do not create unnecessary barriers|
|Generated workbook files|Excelize writes what your code tells it to; responsibility is shared (see [Known limitations](#known-limitations))|

## Supported environments

### Spreadsheet output

Excelize generates files following the Office Open XML (OOXML) spreadsheet format. Files are intended to open in current versions of Microsoft Excel, and in other applications that implement OOXML spreadsheets (for example LibreOffice Calc and Google Sheets). Accessibility behavior of an opened file depends on the application and its assistive technology integration, which we do not control.

Assistive technology commonly used with spreadsheet applications includes:

- Screen readers such as JAWS, NVDA, Narrator, and VoiceOver
- Screen magnifiers and high-contrast / forced-color modes
- Keyboard-only navigation and switch devices
- Speech recognition software

We do not currently run automated, release-gated tests against specific assistive technologies. Reports from people who do are especially valuable to us.

### Documentation and repository

- The documentation is published as web pages and is intended to work with current versions of major browsers and screen readers
- The README, contribution guides, and this file are plain Markdown, which renders accessibly in GitHub and works well with text-based tools and screen readers
- For supported Go versions and platforms, see the README and `go.mod`

## Building accessible spreadsheets with Excelize

Excelize exposes a number of features that help produce accessible files. Which ones apply depends on your content, but these are good starting points. See the [documentation](https://xuri.me/excelize) for the exact API in the version you use.

- **Alternative text for images.** When adding pictures, set the alternative text through the graphic options so screen readers can describe the image
- **Document properties.** Set a meaningful title, subject, and language with `SetDocProps` so assistive technology and file browsers can identify the workbook
- **Meaningful sheet names.** Use descriptive sheet names rather than defaults such as `Sheet1`, and avoid blank sheets used only for spacing
- **Table structure.** Use real tables with header rows (`AddTable`) rather than simulating them with formatting, so header cells are announced with data cells
- **Avoid merged cells where possible.** Merged cells can make navigation and reading order harder for screen reader users. Prefer a simple, regular grid
- **Hyperlinks with clear text and screen tips.** Use descriptive link text instead of "click here", and set the display and tooltip values when calling `SetCellHyperLink`
- **Cell comments and notes.** Use comments to provide extra context and keep important information in cells, not only in comments
- **Data validation with messages.** Provide input and error messages so people entering data understand what is expected
- **Color and contrast.** Do not use color as the only way to convey meaning, and choose font and fill colors with sufficient contrast through the style APIs
- **Readable fonts and sizes.** Use clear fonts at reasonable sizes, and avoid relying on very small text
- **Reading direction and view settings.** Use sheet view options (such as zoom and right-to-left layout) to match your audience's needs
- **Charts.** Give charts clear titles, axis titles, and legends, and consider also providing the underlying data in cells so it can be read without the chart

Microsoft publishes guidance on making Excel documents accessible, and running Excel's built-in Accessibility Checker on generated files is a worthwhile step in your own testing.

## Known limitations

We aim to be transparent about the gaps we know about.

- **No automatic accessibility.** Excelize writes exactly what your program specifies. It does not infer alternative text, reading order, header rows, or contrast, and it does not warn when a generated file would fail an accessibility check
- **No built-in accessibility checker.** The library does not validate output against accessibility guidelines. Use Excel's Accessibility Checker or other tools on your generated files
- **Incomplete coverage of Excel accessibility features.** Not every accessibility-related feature that Excel offers is exposed through the API, and some are only partially supported. If you need one that is missing, please open an issue
- **Application differences.** The same file may behave differently in Excel, LibreOffice, Google Sheets, or on mobile and web versions, especially with screen readers. We cannot fix behavior inside those applications
- **Language.** Documentation is available in a limited set of languages. Machine translation or missing translations may reduce clarity for some readers
- **Testing.** We have not performed a formal accessibility audit of the documentation site or the project's generated output, and we do not claim conformance to a specific standard such as WCAG or EN 301 549

## Reporting an accessibility barrier

If you run into a problem, please tell us. You do not need to be a developer or to know the cause.

### How to report

- Open an issue in the [GitHub repository](https://github.com/xuri/excelize/issues) and include the word **"accessibility"** in the title. If you prefer not to share details publicly, say so in a short issue and we will find another way to follow up
- For general questions or ideas, use the repository's Discussions
- Pull requests that fix an accessibility issue are very welcome

### Helpful details to include

- What you were trying to do (for example, "read the API reference for `AddPicture`" or "open a file generated by my program")
- What got in the way
- The page URL, function name, or a minimal code sample, if applicable
- The Excelize version, Go version, and operating system
- The application and version used to open a file (for example, Microsoft Excel&trade; 365 or LibreOffice 24.x), and the assistive technology and version, if any
- A suggested fix, if you have one (optional)

### What to expect from us

- We will acknowledge accessibility reports as quickly as we reasonably can. Excelize is maintained by maintainers and community contributors rather than a dedicated full-time support team, so we cannot promise a fixed response time, but we will prioritize these reports
- We will tell you whether we can reproduce the problem, what we plan to do, and whether a workaround exists
- We will credit reporters in release notes if they wish, and we will not penalize anyone for reporting

Please note that security vulnerabilities should **not** be reported as public issues. Follow the instructions in the project's security policy (`SECURITY.md`) instead.

## Contributing to accessibility

Ways you can help:

- Improve documentation: clearer wording, descriptive link text, proper heading structure, and text alternatives for any images
- Add examples that demonstrate accessible spreadsheet practices
- Propose or implement API support for accessibility features of the OOXML format
- Test generated files with your assistive technology and report your findings
- Review pull requests for accessibility impact

Please follow the project's contribution guidelines when sending changes, and mention accessibility in the pull request description so reviewers can give it proper attention.
