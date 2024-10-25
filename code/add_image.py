import os
import glob
import io
from PIL import Image, ImageDraw, ImageOps

def get_image_path(image_name, folder_path):
    path = os.path.join(folder_path, '**', '*.jpeg')
    files = glob.glob(path, recursive=True)
    normalized_image_name = image_name.lower().replace(" ", "")
    for file in files:
        normalized_file_name = os.path.splitext(os.path.basename(file))[0].lower().replace(" ", "")
        if normalized_file_name == normalized_image_name:
            return file
    return None


def image_to_round_rectangle(image_path, size=(330, 220), radius=12, supersample=4):
    
    image = Image.open(image_path).convert("RGBA")
    image = image.resize(size, Image.LANCZOS)
    
    mask_size = (size[0] * supersample, size[1] * supersample)
    mask = Image.new('L', mask_size, 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle([(0, 0), mask_size], radius * supersample, fill=255)
    
    mask = mask.resize(size, Image.LANCZOS)
    
    # Optional: Apply a slight blur for even smoother edges
    #mask = mask.filter(ImageFilter.GaussianBlur(radius=0.5))
    output = ImageOps.fit(image, mask.size, centering=(0.5, 0.5))
    output.putalpha(mask)
    output_bytes = io.BytesIO()
    output.save(output_bytes, format='PNG')
    output_bytes.seek(0)
    
    return output_bytes

def get_qrcode_image_path(product_text, folder_path,language):
    if language == "EN":
        target_name = product_text.lower() + '_en.png'
    elif language == "CN":
        target_name = product_text.lower() + '_cn.png'
    else:
        raise ValueError("Unsupported Language")
    for root, _ , files in os.walk(folder_path):
        for file in files:
            if file.lower() == target_name:
                return os.path.join(root, file)
    return None