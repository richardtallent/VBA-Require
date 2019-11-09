# What is VBA-Require?

VBA-Require is a simple **module manager** for Visual Basic for Applications (VBA) modules. Think "npm/NuGet for Excel macros". The goal of this module is to facilitate sharing of common code across an organization or on the open web by making it easy to import and update modules.

## Why?

Because most world's line-of-business logic and data are still in Excel, and will be for years to come. Sure, Excel will become more connected to external data sources over time, and other tools like PowerBI may replace some use cases for Excel, but the core idea of an infinitely-malleable spreadsheet isn't going away.

However, developing robust Excel-based applications is a challenge in part because historically, macros are copied and pasted from one workbook to another (or, let's be honest, from StackOverflow to Excel), like a virtual game of telephone. This means the more popular a macro is, the _more_ difficult it is to squash bugs or improve performance.

## Status

This is a **rough draft**. There is very little code so far.

## Concept

It _is_ possible, using the VBA Extensibility object model, to allow VBA code to read, add, and modify other VBA code modules. However, this functionality requires a reference to the extensibility library to do so, requires special macro permissions, and can be seen as "suspicious" behavior by antivirus programs.

So, rather than giving every Excel workbook "live update" capabilities (my original plan), the current vision is to create a single, open-source Excel workbook that developers can use as a tool to manage the modules in _other workbooks_. This tool would permit them to:

- Find modules with outdated code and upgrade them with a single click
- Add modules by URL (which will in turn add dependencies as needed)
- Determine which modules are dependencies of which others

## Hosting Modules

To publish a module for others to use, a developer simply needs to create a _durable URL_ (one that does not change as versions are updated). The URL can point to either a plain-text `.bas` file, or to an Excel `.xlsm` file containing the module (by name--only the bespoke module is imported).

The advantage of publishing an Excel workbook is that the Excel file can contain dependencies, unit tests, examples, etc., easing the cost of maintaining the module.

Since the publication URL can be private or public, VBA-Require can be easily used to manage both open-source and proprietary modules.

_(Note: my primary use case is Excel. I realize other developers may be creating VBA macros in Access, Word, PowerPoint, etc. I'm not against providing support for other Office document types, if there's a demonstrated need.)_

## Naming Modules

I have no strong opinions about how modules should be named. My personal convention is to use camelCase, with an "m" prefix for workbook-specific modules and "mc" for "common" modules reusable by many workbooks. YMMV.

## Metadata

Since VBA doesn't provide a native metadata structure for modules (where we can declare the version, license, dependencies, etc.), VBA-Require provides a _convention-based_ mechanism using code comments in the header of a module to provide this information.

This header must part of a list of comments at the TOP of the module (before Option Explicit, etc., and with no blank lines at the top). The module header lines here MAY be mixed with other comments not conforming to this format. Basically, once the parser hits the first non-comment line, it should stop looking for Require-VBA headers. Comments may use the "'" form or the "REM" form

Required headers:

- ' MODULE_VERSION: (the version of the module, in semver X.Y.Z format)
- ' MODULE_URL: (where the current version is hosted)
- ' MODULE_NAME: (name of the module, not assumed from the URL)

Recommended/common headers:

- ' MODULE_DEPENDENCY: (url of a module this one depends on. This may appear 0+ times. See next section.)
- ' MODULE_HOMEPAGE: (same as npm, url where the code is hosted with documentation, etc.)
- ' MODULE_LICENSE: (same as npm)
- ' MODULE_COPYRIGHT: (copyright statement)
- ' MODULE_COMPATIBILITY: (notes about compatibility issues with various versions of Excel or platforms)
- ' MODULE_AUTHOR: (name and/or email of the person responsible)
- ' MODULE_DESCRIPTION: (short description of the module's purpose)
- ' MODULE_NOTES: (additional notes, continues for as many lines as needed)
- ' MODULE_USAGE: (instructions for integrating the module with your workbook, continues for as many lines as needed)
- ' MODULE_HISTORY: (multi-line log of version numbers and notes about changes made, with a colon between the version number and note, and the note itself may be multi-line. The URL of the old version can be provided in [brackets] after the note (or directly after the colon if there is no note))
- ' MODULE_SCOPE_METHODS_NEEDED: Comma-delimited list of names of methods that are required to be defined. This differs from dependencies -- where dependencies are looking for specific functionality in a known module, "methods needed" would generally be custom callbacks to be implemented by the spreadsheet.
- ' MODULE_SCOPE_VARIABLES_NEEDED: Comma-delimited list of names of globally-scoped variables the module expects to exist. as with METHODS_NEEDED, this isn't looking for a specific module, it's more for settings, etc. that should be defined by the spreadsheet using the module.
- ' MODULE_SCOPE_RANGES_NEEDED: Comma-delimited list of named ranges required. The "\*" character can be used as a placeholder if the module looks for named ranges matching a certain prefix or suffix.

Lines may wrap into additional lines based on the author's word-wrapping preferences. Additional lines still need the comment marker ("'" or "REM"), but do not need intentional line continuation marks. A header's value stops when a line is reached that is either:

- not a comment line
- a comment line starting with "MODULE\_"
- a blank comment line (whitespace only)
- a comment line beginning with 3 ore more "\*" or "-" characters

## Versioning

VBA-Require defines two version styles: ISO dates or a _very_ simplified "semver" variant.

### Semver-Light

Versions can be numbered using a _highly_ simplified "semver" format -- "major.minor.patch", where major, minor, and patch are whole numbers. Letters, hyphens, qualifiers like "beta", etc. are not supported. Leading zeros in numbers are ignored -- 2.0001, 2.001, and 2.1 are the same version, and "2.10" indicates version two-point-ten (after 2.9), not two-point-one. Numbering should follow the usual semver logic -- patches fix errors, minor versions add new functionality, and major versions indicate breaking changes.

### ISO Date Versioning

Alternatively, module authors may express their versions using the ISO 8601 "yyyy-mm-dd" date format. This may be far more practical for small teams, where trying to use proper semver numbers is more difficult and where date-based versions allow users to reason about the changes in a more meaningful way. Same-day patch numbers are not posisble using this approach -- if a second release is made the same day, it should simply overwrite the original (if that presents a serious risk for you, you should use semver numbering).

## Declaring Dependencies

Despite its lackluster speed, one reason for npm's success is the idea that it's better to have small modules each that do One Thing Well and depend on one another than it is to create monolith packages that reinvent every wheel. VBA-Require is built on the same idea.

Dependencies are declared as URLs, so there is no need for a central "official" repository of modules. Dependencies can be _optionally_ limited to a specific version by putting the requested version **before** the URL, in brackets. Example:

```VB
' MODULE_DEPENDENCY: [2] https://foo.com/myAwesomeModule.xlsm
```

Note that the URL above is still the durable URL of the _current_ version, which may be outside the 2.x range. We are relying on the current version to point us to the old versions using its `MODULE_URL_oldversion` headers.

The number in brackets can be either a specific version or a _partial_ version. This supports both semver and date numbering, and the partial numbering works the same way -- [2.1] includes any version greater than or equal to 2.1 and less than 2.2, and [2019-08] covers any dated version released during August 2019. Explicit ranges and "<" semantics are not supported.

If VBA-Require finds two incompatible dependencies on the same module (say, `1.3+` and `2`), the most recent matching version wins. The same happens if a historical version of a module is not available. While imperfect, this resolves the issues of fighting VBA's global scope. Preferring newer versions improves chances of getting the most bug-free and secure version, and also encourages developers to keep their code up to date with its dependencies.

## Version Hosting

The URL should always host the **most recent** version of the module. This module in return can have _directives_ at the top of the module that point to past versions of the same module still being hosted.

This approach is far simpler than trying to embed version numbers in specific places in the URL. As you deploy a new version of a module, simply add a URL for the old version to the header, post the old file at that location (which need not even be on the same host), and you're done.

This also allows automatic failure -- if an old version should not be used anymore, it can be "wiped from the record" and would no longer be available to be restored. This is especially useful if a particular version has a serious security or other bug and you want to ensure that any code that doesn't support a version newer than the buggy one can at least hopefully regress to an older one.

## Naming Modules, methods, and variables

**NOTE: I'm looking into recommending using `PublicNotCreateable` class modules over standard modules, which would solve most global scope pollution issues. Just need to play with that idea before suggesting it here.**

Where possible, variables and methods should be declared `Private`. This helps prevent collisions, and for Subs, avoids unnecessarily polluting the user's Run Macro dialog with code the module only calls internally.

Public members should be named something explicit enough to prevent collisions with other modules (_e.g._, don't make a `Public Sub Initialize()`).

When you refer to methods in your dependencies, consider using the `ModuleName.MethodName` form so VBA knows precisely which method you're calling, even if there is an accidental collision.

For proprietary modules, it is recommended that module names begin with an organization prefix.

Any public function intended to be exposed as a UDF for formulas should be in `ALLCAPS`.

## Example Headers

These are actual examples of headers from some of the code I've written and maintain, with some tweaks to provide more examples.

```Visual Basic
' ******************************************************************************************************************
' MODULE_NAME: mcGUID
' MODULE_VERSION: 2019-03-25
' MODULE_DESCRIPTION: Provides UDF and VBA function for creating GUIDs using the system provider.
' MODULE_URL: https://-------------/static/vba/common/mcGUID.bas
' MODULE_COMPATIBILITY: Not compatible with Excel for Mac
' MODULE_HISTORY:
' 2019-03-25: Renamed primary method, cleaned up code, added "short" method, set as volatile
' ******************************************************************************************************************
```

```Visual Basic
' ******************************************************************************************************************
' MODULE_NAME: mcHttp
' MODULE_VERSION: 2019-03-25
' MODULE_DESCRIPTION: Pulls information from the web
' MODULE_URL: https://-------------/static/vba/common/mcHttp.bas
' MODEULE_DEPENDENCIES: none
' MODULE_HISTORY:
' 2019-01-01: Original version [https://-------------/static/vba/common/mcHttp-2019-01-01.bas]
' 2019-03-25: Updated to use native worksheet function.
' ******************************************************************************************************************
```

```Visual Basic
' ******************************************************************************************************************
' MODULE_NAME: mcStrings
' MODULE_VERSION: 2017-12-08
' MODULE_DESCRIPTION: Utility UDFs for strings
' MODULE_URL: https://-------------/static/vba/common/mcStrings.bas
' ******************************************************************************************************************
```

```Visual Basic
' ******************************************************************************************************************
' MODULE_NAME: mcChainOfCustody
' MODULE_VERSION: 2017-12-08
' MODULE_URL: https:://-------------/static/vba/common/mcChainOfCustody.bas
' MODULE_DEPENDENCY: https://-------------/static/vba/common/mcRanges.bas
' MODULE_DEPENDENCY: https://-------------/static/vba/common/mcEnvironment.bas
' MODULE_DESCRIPTION: Updates the named range "chainOfCustody" to add the name of the user opening the workbook.
' ******************************************************************************************************************
```

## History

---

2017-10
2018-01 Renamed by suggestion in issue #1, rewrite based on current information
2019-11 Updated to match how I'm currently using this approach IRL
