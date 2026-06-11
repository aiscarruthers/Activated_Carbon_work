import os
import sys
from PyPDF2 import PdfReader, PdfWriter
import argparse

def split_pdf(pdf_path, pages_per_group):
    try:
        # Open the PDF file
        pdf_reader = PdfReader(pdf_path)
        total_pages = len(pdf_reader.pages)
        
        # Determine the number of groups
        num_groups = (total_pages + pages_per_group - 1) // pages_per_group
        
        # Split the PDF into groups
        for i in range(num_groups):
            start_page = i * pages_per_group
            end_page = min(start_page + pages_per_group, total_pages)
            
            pdf_writer = PdfWriter()
            for page_num in range(start_page, end_page):
                pdf_writer.add_page(pdf_reader.pages[page_num])
            
            output_filename = f"{os.path.splitext(pdf_path)[0]}_part_{i+1}.pdf"
            with open(output_filename, 'wb') as output_pdf:
                pdf_writer.write(output_pdf)
            
            print(f"Created: {output_filename},\n")
    
    except Exception as e:
        print(f"Error: {e}")

def split_by_range(pdf_path, page_ranges):
    
    pdf_reader = PdfReader(pdf_path)

    for i, (start_page, end_page) in enumerate(page_ranges):
        pdf_writer = PdfWriter() 

        for page_num in range(start_page - 1, end_page):
            pdf_writer.add_page(pdf_reader.pages[page_num])
        
        output_file_name = f"{os.path.splitext(pdf_path)[0]}_part{i+1}.pdf"
        with open(output_file_name, "wb") as output_pdf:
            pdf_writer.write(output_file_name)
        
        print(f"Created: {output_file_name},\n")


if __name__ == "__main__":
    # if len(sys.argv) != 2:
    #     print("Usage: python split_pdf.py <pdf_path> <pages_per_group>")
    #     sys.exit(1)
    
    parser = argparse.ArgumentParser(description="Split PDF files.")
    parser.add_argument("input_pdf_path", type=str, help="Path tot the input PDF file.")
    parser.add_argument("--ranges",type=str,help="Comma-separated list of page ranges (e.g, '1-3,4-6,7-10')")
    parser.add_argument("--division", type=int, help="Number of pages per group")
    args = parser.parse_args()

    if args.ranges:
        page_ranges = [tuple(map(int, r.split('-'))) for r in args.ranges.split(',')]
        split_by_range(args.input_pdf_path,page_ranges)
    
    elif args.division:
        split_pdf(args.input_pdf_path, args.division)
    
    # pdf_path = sys.argv[1]
    # pages_per_group = int(sys.argv[2])
    
    if not os.path.exists(args.input_pdf_path):
        print(f"Error: The file {args.input_pdf_path} does not exist.")
        sys.exit(1)
    
    # split_pdf(pdf_path, pages_per_group)
