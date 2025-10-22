# PyPolyglot

A tool for creating polyglot files that function as both their original file type and as executable Python archives.

## How It Works

PyPolyglot exploits two key properties:

1. Python can execute scripts or modules from a ZIP archives that [contain a `__main__.py` file](https://peps.python.org/pep-0441/).
1. ZIP archives store their central directory structure at the *end* of the file.  
   Most ZIP readers (including Python!) ignore any data *before* the start of the ZIP structure.

By combining these properties, PyPolyglot can:
- Append a properly structured ZIP archive to the end of any file
- Add a `__main__.py` payload to existing ZIP archives (including Microsoft Office documents!)

This works great on images, documents, audio, and any file type identified by markers at the *beginning* of the file.

### Office Document Handling

Microsoft Office documents (`.docx`, `.xlsx`, `.pptx`) are actually ZIP archives with strict validation rules.  
If Office detects any extra data before or after the ZIP archive, it will throw an "unreadable content" error.  
If Office detects any unnecessary or unreferenced files within the ZIP archive, it will throw an "unreadable content" error.  
You get the idea.

PyPolyglot handles these checks by patching the `[Content_Types].xml` file within Office document archives
to create a generic reference to the added Python payload.  This technique works for all Office documents,
but the payload will be lost if the document is later modified and saved.

## Installation

No installation required!  
Just download [`pypolyglot.py`](pypolyglot.py) and run it with Python 3.6+.

## Usage

```bash
python pypolyglot.py  INPUT  PAYLOAD  OUTPUT
```

**Arguments:**
- `INPUT`: The file to convert into a polyglot (any file type)
- `PAYLOAD`: A Python script that will be executed when the polyglot is run
- `OUTPUT`: Where to save the resulting polyglot file

### Examples

Create a polyglot Word document:
```bash
python pypolyglot.py document.docx payload.py output.docx
```

Create a polyglot image:
```bash
python pypolyglot.py image.png payload.py output.png
```

Create a polyglot from an existing ZIP file:
```bash
python pypolyglot.py archive.zip payload.py output.zip
```

### Running the Polyglot

Once created, you can:
- Open the file normally as a document, image, etc
- Execute it as a Python script: `python output.docx`

### Extracting the Payload

You can extract the embedded Python script from any polyglot: `unzip output.docx __main__.py`

## Limitations

- The payload must be a single script and can't use any third-party libraries the user doesn't already has them installed
- The output file will be slightly larger than the original (due to the embedded Python payload)
- Some file tools may flag polyglots as corrupted
- Most tools will remove the payload if the polyglot is later edited

### Possible Improvements

PyPolyglot doesn't support it, but it is possible to package an entire Python virtual environment into the ZIP archive.  
When Python executes from a ZIP archive, the archive is added to sys.path/PYTHONPATH meaning
other files and libraries from the archive can be imported and used.  In fact, this is
basically what [pex](https://github.com/pex-tool/pex) and [zipapp](https://docs.python.org/3/library/zipapp.html)
use to create executable archives.  
However, even if the shebang is omitted, there are still some major limitations:
- It still depends on the user having the correct version of Python available
- Third-party libraries must be *pure* Python or exactly match the user's Python implementation, ABI, and platform

Given that PyPolyglot has no way to check these things, it currently only supports a single payload script,
not a payload script and virtual environment.

Furthermore, it might be possible to make Office document polyglots persistent across file edits.  
The current method for preventing the "unreadable content" error is generic across all Office products,
but it might be possible to add application-specific references that retain the payload after edits.
