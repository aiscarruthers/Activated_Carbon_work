import sys
import os
from PyPDF2 import PdfReader, PdfWriter


def rotate_pdf(input_path, output_path, rotation, clockwise=True):
    if not os.path.exists(input_path):
        print(f"Error: File '{input_path}' not found.")
        sys.exit(1)

    reader = PdfReader(input_path)
    writer = PdfWriter()

    # Determine rotation direction
    angle = rotation if clockwise else -rotation

    for page in reader.pages:
        page.rotate(angle)  # PyPDF2 handles 90, 180, 270 etc.
        writer.add_page(page)

    with open(output_path, "wb") as f:
        writer.write(f)

    print(f"Saved rotated PDF to: {output_path}")


def main():
    if len(sys.argv) < 4:
        print("Usage:")
        print("python rotate_pdf.py <input> <rotation> <direction> [output]")
        print("Example:")
        print("python rotate_pdf.py file.pdf 90 cw output.pdf")
        sys.exit(1)

    input_file = sys.argv[1]
    rotation = int(sys.argv[2])
    direction = sys.argv[3].lower()

    if direction not in ["cw", "ccw"]:
        print("Direction must be 'cw' or 'ccw'")
        sys.exit(1)

    clockwise = direction == "cw"

    # Default output name
    if len(sys.argv) >= 5:
        output_file = sys.argv[4]
    else:
        base, ext = os.path.splitext(input_file)
        output_file = f"{base}_rotated{ext}"

    rotate_pdf(input_file, output_file, rotation, clockwise)


if __name__ == "__main__":
    main()