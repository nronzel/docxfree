# Docxfree

## Overview

Removes password protection from `.docx` files.
Forgot the password to your `.docx` file?
This utility will create a copy of the original file with
the password protection removed, append `unlocked-` to the beginning
of the filename, and saves them into a `docxfreed/` directory.
A copy is created to avoid potential corruption of the original file.

> [!WARNING]
> Only works on `.docx` files. (Office 2007 or newer)
> Will NOT work with `.doc` files.

## Installation

### Manual Installation

Clone the repo:

```sh
git clone https://github.com/nronzel/docxfree.git

cd docxfree
```

Build the program:

```sh
go build
```

### Release

Download the correct binary from the latest [release](https://github.com/nronzel/docxfree/releases).

## Usage

You can specify a single filename, or a path to a directory.

If you specify a path, you can also specify an optional depth argument if there are
sub-directories that contain `.docx` files that you need processed.

### Arguments

| Arg        | Description                              |
| ---------- | ---------------------------------------- |
| -f (file)  | Specify a single .docx file.             |
| -p (path)  | Specify a directory of .docx files.      |
| -d (depth) | \*Optional - Recursion depth. Default: 1 |

Example:

```sh
docxfree -f document.docx
# or
docxfree -p documents/
docxfree -p documents/ -d 2
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
