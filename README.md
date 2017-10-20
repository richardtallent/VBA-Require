# What is RequireVBA?
RequireVBA is a very simple **module manager** for Visual Basic for Applications (VBA) modules. Think "npm for Excel macros".

The goal of this module is to facilitate sharing of common code across an organization or on the open web by making it easy to import and update modules.

## Why?
Because most world's line-of-business logic and data are still in Excel, and will be for years to come. Sure, Excel will become more connected to external data sources over time, and other tools like PowerBI may replace some use cases for Excel, but the core idea of an infinitely-malleable spreadsheet isn't going away.

However, it's far too difficult to develop robust Excel-based applications, in no small part because historically, macros are copied and pasted from one workbook to another (or, let's be honest, from StackOverflow to Excel), like a virtual game of telephone. This means the more popular a macro is, the *more* difficult it is to squash bugs or make performance or code improvements across the board.

I'm hoping that this tool not only makes it easier to share code within an organization, but that it also creates a surge in well-crafted, open-source Excel modules.

## Status

This is a germ of an idea. I'm still sketching out the API, **there's no code yet**. I've created code to perform some of the essential operations, such as replacing a VBA module with a new file off the web, but only for ad-hoc needs.

**Again, this is a very rough draft... just an idea that's been in my head for awhile and a late-night session of writing this up. I'm definitely looking for ideas, blind spots, etc., so please file issues for discussion if you're interested in seeing this idea happen!**

## API

### Function RequireVBA(ByVal *url* As String, Optional ByVal *version* As String = "") AS Boolean

When this procedure is called, it checks the workbook to ensure that the module from the specified **url** is available as a module in the workbook. If not, it makes a request to said URL. If successful, it automatically creates the module and populates the code (and returns "True"), thus allowing your code afterward to make use of what it needs.

These calls should be at the *top* of your module, outside of any other procedures and just after any *Option* directives or header comments.

This procedure is recursive--if a module created has its own dependencies, they will also be pulled, etc. Modules required by more than one dependency path will only be pulled once.

Note that the URL can point to either a ".vba" file, in which case the module is treated as text, or to an ".xslm" file, in which case the Excel file is loaded in the background, the module is extracted from the file, and the Excel file is closed. (This is slower for restoring a package, but makes it far easier to manage module packages, because the Excel file wrapper is easily edited in Excel and can contain unit tests, dependencies, etc.)

The optional **version** argument provides a way to restrict the version of the module to be used. If not provided, the most recent version is used.

### Function RequireVBA_Outdated() As String

This procedure traverses all RequireVBA-managed modules, pulls the current code from their URLs, and determines whether or not the online version is up to date. It returns a list of packages installed, the installed versions, and the current version online. (still working out if this should return something easy like a string for Immediate use, or a 2D array to make it easier to display to the common user).

### Function RequireVBA_Update(ByVal *url* As String) As Boolean

This function replaces an existing module with the current version online. If *url* is "*", it will traverse and update all modules in the workbook. It returns **True** if all updates were successful, or **False** if the module at the dependency or any of its own dependencies failed to be updated.

For performance, it is *not* recommended that this be wired to your *Workook_Open()* event. Instead, this could be done based on user initiation, or perhaps on a scheduled basis (say, only check if a date stored on the spreadsheet somewhere is more than a month old, and update that date after calling this and getting a successful response).

### Function RequireVBA_Log() As String()

This returns a log of all activity since the workbook was loaded. You can use this to display any issues to the user. The log is a 2-D array, where the rows represent log entries, and the columns are:
  1. Status (OK, Warning, Error, Info, or Debug)
  2. Timestamp
  3. Method
  4. Module
  5. Message

## What about conflicting version limits?

If a call to RequireVBA or RequireVBAUpdate results in two incompatible dependencies to the same package (say, ">2.1" and "<1.5"), it will pull the highest version within the dependency chain. While imperfect, this is needed because of Excel's global namespace, and preferring newer versions improves chances of fixing bugs and uncovering incompatibility.

## Version Numbering

RequireVBA() implements a *simplified* "semver"-style mechanism for describing module versions and dependencies. Supported grammar includes:
  - Use of major.minor.patch version numbers
  - Use of "X", "x", or "*" placeholders
  - Use of "<", ">", "<=", ">=", or "=" operators
  - Use of "~" and "^" ranges
  - Use of multiple operators (e.g., ">=1.3 < 4.0")
  - Whitespace is ignored

Not supported:
  - Use of qualifiers like "-beta"
  - Logical OR ("||")
  - Hyphens

Note that RequireVBA itself has a dependency on a package that parses X.Y.Z version numbers and compares them to other versions and to this grammar.

## Version Hosting

The URL should always host the **most recent** version of the module. This module in return can have *directives* at the top of the module that point to past versions of the same module still being hosted.

This approach is far simpler than trying to embed version numbers in specific places in the URL. As you deploy a new version of a module, simply add a URL for the old version to the header, post the old file at that location (which need not even be on the same host), and you're done.

This also allows automatic failure -- if an old version should not be used anymore, it can be "wiped from the record" and would no longer be available to be restored. This is especially useful if a particular version has a serious security or other bug and you want to ensure that any code that doesn't support a version newer than the buggy one can at least hopefully regress to an older one.

## Directives

For *RequireVBA* to function, each module needs to have some special comments to relay metadata.

Required directives:
  - ' RVBA_VERSION:			(the version of the module, in semver X.Y.Z format)
  - ' RVBA_URL: 			(where the current version is hosted)
  - ' RVBA_NAME: 			(name of the module, not assumed from the URL)
  - ' RVBA_URL_{oldversion}: (where a particular old version [in X.Y.Z form] may be found). There may be any number of these lines for different historical versions.

Recommended directives:
  - ' RVBA_HOMEPAGE:		(same as npm, url where the code is hosted with documentation, etc.)
  - ' RVBA_LICENSE:			(same as npm)
  - ' RVBA_AUTHOR:			(name and/or email of the person responsible)
  - ' RVBA_DESCRIPTION:		(description of the module's purpose)

Dependencies are specified by calling *RequireVBA()*, ideally near the top of your module, so there are no comment-driven directives for them.

## Class Modules

I don't believe the first version of this tool will support class modules, since my use cases for those are rare, but I'm open to the idea.

## Namespacing

In your modules, it is *recommended* that any public members be prefixed with your module name and an underscore. If the module has only a single method, the module name itself can suffice.

If the intent is to create functions that are exposed as UDFs, holding to this approach militantly may be impractical, but good naming is still recommended to avoid name collisions.

Speaking of which, it is recommended that module names (and thus their public members) begin with an organization prefix, to help alleviate collisions. Otherwise, you're just asking for problems when three separate modules call themselves "Utilities", making it impossible for RequireVBA() to properly manage them.

