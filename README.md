# What is VBA-Require?
VBA-Require is a simple **module manager** for Visual Basic for Applications (VBA) modules. Think "npm/NuGet for Excel macros". The goal of this module is to facilitate sharing of common code across an organization or on the open web by making it easy to import and update modules.

## Why?
Because most world's line-of-business logic and data are still in Excel, and will be for years to come. Sure, Excel will become more connected to external data sources over time, and other tools like PowerBI may replace some use cases for Excel, but the core idea of an infinitely-malleable spreadsheet isn't going away.

However, developing robust Excel-based applications is a challenge in part because historically, macros are copied and pasted from one workbook to another (or, let's be honest, from StackOverflow to Excel), like a virtual game of telephone. This means the more popular a macro is, the *more* difficult it is to squash bugs or improve performance.

## Status
This is a **rough draft**. There is no code yet.

## Concept
It *is* possible, using the VBA Extensibility object model, to allow VBA code to read, add, and modify other VBA code modules. However, this functionality requires a reference to the extensibility library to do so, requires special macro permissions, and can be seen as "suspicious" behavior by antivirus programs.

So, rather than giving every Excel workbook "live update" capabilities (my original plan), the current vision is to create a single, open-source Excel workbook that developers can use as a tool to manage the modules in *other workbooks*. This tool would permit them to:

- Find modules with outdated code and upgrade them with a single click
- Add modules by URL (which will in turn add dependencies as needed)
- Determine which modules are dependencies of which others

## Hosting Modules
To publish a module for others to use, a developer simply needs to create a *durable URL* (one that does not change as versions are updated). The URL can point to either a plain-text `.bas` file, or to an Excel `.xlsm` file containing the module (by name--only the bespoke module is imported).

The advantage of publishing an Excel workbook is that the Excel file can contain dependencies, unit tests, etc., easing the cost of maintaining the module.

Since the publication URL can be private or public, VBA-Require can be easily used to manage both open-source and proprietary modules.

*(Note: my primary use case is Excel. I realize other developers may be creating VBA macros in Access, Word, PowerPoint, etc. I'm not against providing support for other Office document types, if there's a demonstrated need.)*

## Metadata
Since VBA doesn't provide a native metadata structure for modules (where we can declare the version, license, dependencies, etc.), VBA-Require provides a *convention-based* mechanism using code comments in the header of a module to provide this information.

Required headers:
  - ' MODULE_VERSION:		(the version of the module, in semver X.Y.Z format)
  - ' MODULE_URL: 			(where the current version is hosted)
  - ' MODULE_NAME: 			(name of the module, not assumed from the URL)

Recommended/common headers:
  - ' MODULE_DEPENDENCY:	(url of a module this one depends on. This may appear 0+ times. See next section.)
  - ' MODULE_HOMEPAGE:		(same as npm, url where the code is hosted with documentation, etc.)
  - ' MODULE_LICENSE:		(same as npm)
  - ' MODULE_AUTHOR:		(name and/or email of the person responsible)
  - ' MODULE_DESCRIPTION:	(description of the module's purpose)
  - ' MODULE_URL_*oldversion*: (where a particular old version [in X.Y.Z form] may be found). There may be any number of these lines for different historical versions.

## Declaring Dependencies
While I have to remind myself of the benefits when I'm waiting on `npm install`, one of the reasons for npm's success has been the idea that it's better to depend on another package than to create monolith packages that reinvent every wheel. I'd like VBA-Require to encourage the same behavior.

Dependencies are URL-based, so there is no need for a central "official" repository of modules. Dependencies can be limited to a specific version by putting the requested version **before** the URL, in brackets. Example:

```VB
' MODULE_DEPENDENCY: [2.x] https://foo.com/myAwesomeModule.xlsm
```

Note that the URL above is still the durable URL of the *current* version, which may be outside the 2.x range. We are relying on the current version to point us to the old versions using its `MODULE_URL_oldversion` headers.

Versions should be numbered using a *highly* simplified "semver" format -- "major.minor.patch", where major, minor, and patch are decimal numbers. Letters, hyphens, qualifiers like "beta", etc. are not supported.

Dependencies are also specified using a format similar to `packages.json`, but far more simplified. To declare a specific version of a dependency, use the same "major.minor.patch" form, but *leave off* the words (right to left) that you aren't concerned about. For example, if you wish to rely on version 2.0.0 or above but not 3.0.0, simply say `2`. To rely specifically on version 2.1 but any patch thereof, use `2.1`. To ask for version 2.1 or higher, add a `+` to the end. The `+` *only applies* to the word it follows. Ranges (e.g., "2.1 - 2.3") are not supported, nor are "<" semantics.

If VBA-Require finds two incompatible dependencies on the same module (say, `1.3+` and `2`), the most recent matching version wins. While imperfect, this resolves the issues of fighting VBA's global scope. Preferring newer versions improves chances of getting the most bug-free and secure version, and also encourages developers to keep their code up to date with its dependencies.

## Version Hosting
The URL should always host the **most recent** version of the module. This module in return can have *directives* at the top of the module that point to past versions of the same module still being hosted.

This approach is far simpler than trying to embed version numbers in specific places in the URL. As you deploy a new version of a module, simply add a URL for the old version to the header, post the old file at that location (which need not even be on the same host), and you're done.

This also allows automatic failure -- if an old version should not be used anymore, it can be "wiped from the record" and would no longer be available to be restored. This is especially useful if a particular version has a serious security or other bug and you want to ensure that any code that doesn't support a version newer than the buggy one can at least hopefully regress to an older one.

## Naming Modules, methods, and variables

Where possible, variables and methods should be declared `Private`. This helps prevent collisions, and for Subs,  avoids unnecessarily polluting the user's Run Macro dialog with code the module only calls internally.

Public members should be named something explicit enough to prevent collisions with other modules (*e.g.*, don't make a `Public Sub Initialize()`).

When you refer to methods in your dependencies, you should prefer the `ModuleName.MethodName` form so VBA knows precisely which method you're calling, even if there is an accidental collision.

For proprietary modules, it is recommended that module names begin with an organization prefix.

Any public function intended to be exposed as a UDF for formulas should be in `ALLCAPS`.

## Class Modules
I don't believe the first version of this tool will support class modules, since my use cases for those are rare, but I'm open to the idea.

## History
--------------------------------
2017-10
2018-01 Renamed by suggestion in issue #1, rewrite based on current information

