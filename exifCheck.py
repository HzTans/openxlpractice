from PIL import Image
from PIL.ExifTags import TAGS
from pathlib import Path

file_path = "./img/mouseV.jpg"
img=Image.open(file_path)
file_name = Path(file_path).stem
print(Path(file_path).parent)
print(file_name)
print(str(Path(file_path).parent)+"/"+file_name+"CHANGE")
ex = img._getexif()
exifinfo = img._getexif()
ret={}
if exifinfo != None:
    for tag, value in exifinfo.items():
        decoded = TAGS.get(tag, tag)
        ret[decoded] = value
if "Orientation" in ret:
    print(ret["Orientation"])
    if ret["Orientation"] == 3:
                img = img.rotate(180, expand=True)
    elif ret["Orientation"] == 6:
        img = img.rotate(270, expand=True)
    elif ret["Orientation"] == 8:
        img = img.rotate(90, expand=True)
    img.save(str(Path(file_path).parent)+"/"+file_name+"CHANGGGGGGGGGGGGGGGGGGGGGGGGGE"+".jpg")
    print(ret["Orientation"])
