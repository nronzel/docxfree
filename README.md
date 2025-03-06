# Docxfree

## Overview

Removes password protection from `.docx` files.
Forgot the password to your `.docx` file?
This utility will create a copy of the original file with
the password protection removed, append `unlocked-` to the beginning
of the filename, and saves them into a `docxfree_unlocked` directory.
A copy is created to avoid potential corruption of the original file.

> [!WARNING]
> Only works on `.docx` files. (Office 2007 or newer)
> Will NOT work with `.doc` files.

## Installation

## Usage

### Arguments

| Arg       | Description                        |
| --------- | ---------------------------------- |
| -f (file) | Specify a single .docx file        |
| -p (path) | Specify a directory of .docx files |

Example:

```sh
docxfree -f document.docx
# or
docxfree -p documents/
```

## How It Works

Office 2007 and newer `.docx` files are actually just a zip archive with
multiple XML files. When you apply password protection to the document,
a `<w:documentProtection>` node is added to the `settings.xml` file in the archive.

This utility creates a copy of all of the XML files in the `.docx` archive,
and creates a modified copy of the `settings.xml` file with the protection node
removed.

Older Office files (`.doc`) are a single binary file and not an XML archive,
so this method of removing password protection will not work.
