import easyocr
from pdf2image import convert_from_path
from PIL import Image, ImageEnhance

# Initialize the EasyOCR reader
reader = easyocr.Reader(lang_list=['en'])

# Path to the PDF document
pdf_path = "C:\\Users\\ACarruther\\Downloads\\TT-01-02-01_spec.pdf"

# Convert PDF to images
images = convert_from_path(pdf_path)

# Perform OCR on each image and extract recognized text
recognized_text = []

for image in images:
    
    enhanced_image = ImageEnhance.Contrast(image).enhance(2.0)


    # Save the image as a temporary JPEG file
    temp_image_path = 'temp_image.jpg'
    enhanced_image.save(temp_image_path, 'JPEG')

    # Perform OCR on the saved image
    results = reader.readtext(temp_image_path)
    image_text = [result[1] for result in results]
    recognized_text.extend(image_text)

# Print the recognized text
print("Recognized Text:")
for text in recognized_text:
    print(text)






