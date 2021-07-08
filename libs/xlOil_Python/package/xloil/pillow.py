from PIL import Image
import xloil as xlo

@xlo.returner(typeof=Image.Image, register=True)
class ReturnImage:
    
    """
        Inserts an image associated with the calling cell. A second call
        removes any image previously inserted by the same calling cell.

        Parameters
        ----------

        size:  
            * A tuple (width, height) in points. 
            * "cell" to fit to the caller size
            * "img" or None to keep the original image size
        pos:
            A tuple (X, Y) in points. The origin is determined by the `origin` argument
        origin:
            * "top" or None: the top left of the calling range
            * "sheet": the top left of the sheet
            * "bottom": the bottom right of the calling range
    """
    def __init__(self, size=None,  pos=(0, 0), origin:str=None):
        self._shape = (size, pos, origin)

    def write(self, val):
        return xlo.insert_cell_image(lambda filename: val.save(filename, format="png"), *self._shape)

