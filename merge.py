import sys
from pathlib import Path

from docx import Document
from docxcompose.composer import Composer


def merge_folder(folder_path, output_file):
    folder = Path(folder_path)

    # Find all DOCX files
    files = sorted(
        folder.glob("*.docx"),
        key=lambda x: x.name.lower()
    )

    if not files:
        print("No .docx files found!")
        return

    print(f"Found {len(files)} DOCX files:\n")

    for i, file in enumerate(files, 1):
        print(f"{i:02d}. {file.name}")

    # First document = master
    master = Document(files[0])
    composer = Composer(master)

    # Append remaining documents
    for file in files[1:]:
        doc = Document(file)

        # Page break before next document
        master.add_page_break()

        composer.append(doc)

    # Save
    composer.save(output_file)

    print("\nDone!")
    print(f"Merged file: {output_file}")


def main():

    if len(sys.argv) != 2:
        print("Usage:")
        print("python merge_docx.py <folder>")
        print()
        print("Example:")
        print("python merge_docx.py invoices")
        return

    folder = Path(sys.argv[1])

    if not folder.is_dir():
        print(f"Folder not found: {folder}")
        return

    output_file = folder / "merged.docx"

    merge_folder(folder, output_file)


if __name__ == "__main__":
    main()