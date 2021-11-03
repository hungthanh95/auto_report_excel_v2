from pytesseract import image_to_string
from pytesseract.pytesseract import tesseract_cmd
from PIL import Image

# tesseract_cmd = r"D:\\MyProjects\\auto_report\\Tesseract-OCR\\tesseract.exe"

def set_tesseract_cmd(path):
    tesseract_cmd = r"{}".format(path)

def crop_and_resize_img(img_url, box):
    (left_ratio, right_ratio, top_ratio, bottom_ratio) = box
    img = Image.open(r'{}'.format(img_url))
    width, height = img.size

    left = width * left_ratio
    right = width * right_ratio
    top = height * top_ratio
    bottom= height * bottom_ratio

    img_crop = img.crop((left, top, right, bottom))
    img_resize = img_crop.resize((img_crop.size[0] * 2, img_crop.size[1] * 3))

    # img_resize.show() # only debug
    return img_resize    


def convert_img_to_string(image):
    raw_data = image_to_string(image, config=' --psm 6 --oem 3  -c tessedit_char_whitelist=0123456789.Hzmusz')
    value = 0

    data_split = raw_data.split('\n')
    if 'Hz' in data_split[0]:
        filter = 'Hz'
        scale = 1
    elif 'ms' in data_split[0]:
        filter = 'ms'
        scale = 1
    elif 's' in data_split[0]:
        filter = 's'
        scale = 0.001
    else:
        filter = 'us'
        scale = 1000
    try:
        filter_data = data_split[0].replace(filter, '')
        value = float(filter_data) / scale
    except:
        pass
    return value


def crop_and_resize_duration_img(img_url):
    return crop_and_resize_img(img_url, (0.8345, 1, 0.392, 0.44))

    
def crop_and_resize_freq_img(img_url):
    return crop_and_resize_img(img_url, (0.8345, 1, 0.478, 0.52)), crop_and_resize_img(img_url, (0.8345, 1, 0.36, 0.4))  #crop_and_resize_img(img_url, (0.8345, 1, 0.279, 0.32))



def convert_img_duration_to_string(image_path):
    img_duration  = crop_and_resize_duration_img(image_path)
    return convert_img_to_string(img_duration)


def convert_img_freq_to_string(image_path):
    img_freq1, img_freq2 = crop_and_resize_freq_img(image_path)
    value = convert_img_to_string(img_freq1)
    if value == 0:
        value = convert_img_to_string(img_freq2)
    return value


