from PIL import Image
import xloil as xlo

@xlo.returner(types=Image.Image, register=True)
class ReturnImage:
    
    def __init__(self, size=None,  pos=(0, 0), origin:str=None):
        self._shape = (size, pos, origin)

    def __call__(self, val):
        return xlo.insert_cell_image(lambda filename: val.save(filename, format="png"), *self._shape)

