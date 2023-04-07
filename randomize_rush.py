import sys
import random
import PyPDF2


def randomize_pdf(input_pdf, output_pdf, k):
    # Read the input PDF
    pdf_reader = PyPDF2.PdfFileReader(input_pdf)

    # Get the total number of pages
    num_pages = len(pdf_reader.pages)

    # Generate lists of page indices for the first k-1 pages (excluding the first 2 pages) and the remaining n-k pages
    first_page_indices = list(range(2, k-1))
    last_page_indices = list(range(k, num_pages))

    # Shuffle the first k-1 pages (excluding the first 2 pages) and the remaining n-k pages
    random.shuffle(first_page_indices)
    random.shuffle(last_page_indices)

    # Create a new PDF
    pdf_writer = PyPDF2.PdfFileWriter()

    # Add the first 2 pages to the new PDF
    pdf_writer.addPage(pdf_reader.getPage(0))
    pdf_writer.addPage(pdf_reader.getPage(1))

    # Add the shuffled pages within the scope of 2 and k-1 to the new PDF
    for index in first_page_indices:
        pdf_writer.addPage(pdf_reader.getPage(index))

    # Add the kth page to the new PDF
    pdf_writer.addPage(pdf_reader.getPage(k-1))

    # Add the shuffled remaining n-k pages within the scope of n-k and n to the new PDF
    for index in last_page_indices:
        pdf_writer.addPage(pdf_reader.getPage(index))

    # Write the output PDF
    with open(output_pdf, 'wb') as output_file:
        pdf_writer.write(output_file)


if __name__ == '__main__':
    if len(sys.argv) != 4:
        print("Usage: python randomize_pdf.py input_pdf output_pdf k")
        sys.exit(1)

    input_pdf = sys.argv[1]
    output_pdf = sys.argv[2]
    k = int(sys.argv[3])

    randomize_pdf(input_pdf, output_pdf, k)
    print(
        f"Randomized PDF with fixed first 2 pages and {k}th page saved to {output_pdf}")
