# This draws an image into Excel

# Excel limit rows 1,048,576
# Excel limit columns 16,384


from models.ExcelModels import EditExcel

# Path to the image
image_path = r'C:\Users\sabax\source\repos\image_to_excel\testImage\Z3VA8oJw_400x400.jpg'

EditExcel(image_path)
