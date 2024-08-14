# This draws an image into Excel

# Excel limit rows 1,048,576
# Excel limit columns 16,384


from models.ImageModels import ImportImage

# Path to the image
image_path = r'\testImage\Z3VA8oJw_400x400.jpg'

ImportImage(image_path)
