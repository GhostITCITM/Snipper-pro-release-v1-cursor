import os
import subprocess
import re
from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas


def run_tesseract(image_path):
    result = subprocess.run([
        'tesseract', image_path, 'stdout', '-l', 'eng', '--oem', '1', '--psm', '6'],
        stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True, text=True
    )
    return result.stdout


def create_test_image(path, text):
    img = Image.new('RGB', (200, 60), color='white')
    draw = ImageDraw.Draw(img)
    draw.text((10, 20), text, fill='black')
    img.save(path)


def create_test_pdf(path, text):
    c = canvas.Canvas(str(path))
    c.drawString(100, 750, text)
    c.save()


def test_image_decimal(tmp_path):
    img_path = tmp_path / 'decimal.png'
    create_test_image(img_path, 'Total: 123.45')
    output = run_tesseract(str(img_path))
    assert re.search(r"123[.,]45", output)


def test_pdf_decimal(tmp_path):
    pdf_path = tmp_path / 'decimal.pdf'
    img_out = tmp_path / 'pdf_image'
    create_test_pdf(pdf_path, 'Amount: 987.65')
    subprocess.run(['pdftoppm', '-png', '-singlefile', str(pdf_path), str(img_out)],
                   check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    png_path = str(img_out) + '.png'
    output = run_tesseract(png_path)
    assert re.search(r"987[.,]65", output)
