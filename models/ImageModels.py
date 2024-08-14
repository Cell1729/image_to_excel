import cv2
import numpy as np
from models.ExcelModels import EditExcel

import cv2
import numpy as np
from models.ExcelModels import EditExcel

class ImportImage():
    """
    Import an image from the given path
    """
    def __init__(self, image_path):
        self.editExcel = EditExcel()
        self.image = cv2.imread(image_path)
        self.GetPixelValue()
     
    def GetPixelValue(self):
        """
        Get the pixel value of the image and convert BGR to RGB
        """
        image_shape = self.image.shape
        pixel_value = np.zeros((image_shape[0], image_shape[1], 3), dtype=int)
        if image_shape[0] < 1048576 and image_shape[1] < 16384:
            for img_x in range(image_shape[0]):
                for img_y in range(image_shape[1]):
                    # Convert BGR to RGB
                    b, g, r = self.image[img_x, img_y]
                    pixel_value[img_x, img_y] = [r, g, b]
                    r_val, g_val, b_val = pixel_value[img_x, img_y]
                    hex_color = f'{r_val:02x}{g_val:02x}{b_val:02x}'
                    self.editExcel.fillColor(img_x, img_y, hex_color)
            self.editExcel.saveExcel()
        else:
            print(f"Image size is too large to fit in Excel ({image_shape[0]}x{image_shape[1]}).")