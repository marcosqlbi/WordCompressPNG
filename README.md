# WordCompressPNG

WordCompressPNG is a command-line tool that compresses PNG images embedded inside Microsoft Word documents.

## Purpose

The tool helps reduce the size of Word documents by optimizing PNG files stored within them. This is useful when documents contain large or unoptimized images and you want to shrink the file size without manually extracting and recompressing every image.

## Features

- Processes Microsoft Word documents from the command line
- Finds embedded PNG images inside the document
- Compresses PNG images to reduce document size
- Preserves the document structure while updating image content
- Helps simplify bulk optimization of image-heavy Word files

## How it works

1. Open a Word document.
2. Scan the document for embedded PNG images.
3. Compress each image found.
4. Save the updated document with smaller image data.

## Typical use cases

- Reducing the size of documents before sharing them
- Optimizing reports, proposals, and manuals with many screenshots
- Cleaning up documents that grew large due to pasted PNG images

## Getting started

### Prerequisites

- .NET / C# build environment
- A Microsoft Word document containing embedded PNG images

### Build

Clone the repository and build it with your preferred .NET tooling:

```bash
dotnet build
```

### Run

Run the tool against a Word document from the command line. Refer to the project’s source code or command-line help for the exact arguments supported by the current version.

```bash
dotnet run -- <your-document.docx>
```

## Example workflow

1. Make a copy of the original `.docx` file.
2. Run WordCompressPNG on the copy.
3. Compare the file size before and after.
4. Open the optimized document in Word to verify the result.

## Notes

- Always keep a backup of important documents before modifying them.
- Results may vary depending on the images contained in the document.

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.
