# Excel VBA type Definitions Package (excel-types)
 
 - This package provides a way to use auto-complete for Excel VBA Class
 - You can Create your own vba type definition packages our update this.
 - Type definition packages are named like: name-types 
 - The packages are store in www.xvba.dev
 - Ech excel vba class has your own file definition 
 - The files extension has to be filename.d.vb
 - Auto-complete just expose Public types


## Create,install and share VBA Packages With Xvba-cli and Xvba Repository:

- Xvba Repository : <a href="https://www.xvba.dev"> www.xvba.dev</a>
- XVBA-CLI Command Line Interface for XVBA VSCode extension <a href="https://www.npmjs.com/package/@localsmart/xvba-cli">@localsmart/xvba-cli </a>

## Install

- For instal excel-types just use XVBA-CLI install command

```
 npx xvba install excel-types
```
## Comments Block

- Use comments blocks below for documenting class/methods/Subs/Functions/Properties
- The comments blocs has to start with '/* and ends with '*/

```

'/*
'Represents the entire Microsoft Excel application.
'
'
'*/
Public Class Application()

```

```
'/*
'Returns a Range object that represents the active cell in the active window 
'(the window on top) or in the specified window. If the window isn't displaying 
'a worksheet, this property fails. Read-only.
'
'@type {Object.<Range>}
'
'*/
Public Property ActiveCell As Range

```