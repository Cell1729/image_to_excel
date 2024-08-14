import cv2
import numpy as np

class ImportImage():
    """
    Import an image from the given path
    """
    def __init__(self, image_path):
        self.image = cv2.imread(image_path)
        self.image_shape, self.pixel_value = self.GetPixelValue()  # pixel_value を取得
     
    def GetPixelValue(self):
        """
        Get the pixel value of the image and convert BGR to RGB
        """
        image_shape = self.image.shape
        pixel_value = np.zeros((image_shape[0], image_shape[1], 3), dtype=int)

        for img_x in range(image_shape[0]):
            for img_y in range(image_shape[1]):
                # Convert BGR to RGB
                b, g, r = self.image[img_x, img_y]
                pixel_value[img_x, img_y] = [r, g, b]

        return image_shape, pixel_value
