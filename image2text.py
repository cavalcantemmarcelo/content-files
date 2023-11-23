import os
from PIL import Image
import pytesseract
from openpyxl import Workbook


def extract_text_from_image(image_path):
    try:
        with Image.open(image_path) as img:
            text = pytesseract.image_to_string(img, lang='eng')
            return text
    except Exception as e:
        print(f"Error processing {image_path}: {str(e)}")
        return None


def process_images_folder(input_folder, output_excel_path):
    # Create a new Excel workbook and get the active sheet
    wb = Workbook()
    ws = wb.active

    # Write header row to the Excel file
    ws.append(['Image Name', 'Content', 'Image Path'])

    # Traverse files in the input folder
    for root, dirs, files in os.walk(input_folder):
        for file_name in files:
            if file_name.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                image_path = os.path.join(root, file_name)

                # Extract text from the image
                image_text = extract_text_from_image(image_path)

                # Write information to the Excel file
                ws.append([file_name, image_text, image_path])

    # Save the Excel file
    wb.save(output_excel_path)


# Example usage
input_image_folder_path = './questions'
output_excel_path = 'text2image.xlsx'
process_images_folder(input_image_folder_path, output_excel_path)
