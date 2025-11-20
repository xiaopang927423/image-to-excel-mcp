"""
Utility functions for image processing.
"""
import base64
from PIL import Image
from io import BytesIO
import logging


def encode_image_to_data_uri(image_path: str) -> str:
    """
    Encode an image to a data URI with appropriate format detection.
    
    Args:
        image_path: Path to the input image file
        
    Returns:
        Data URI string with appropriate MIME type
    """
    with open(image_path, "rb") as image_file:
        base64_image = base64.b64encode(image_file.read()).decode("utf-8")

    try:
        img = Image.open(BytesIO(base64.b64decode(base64_image)))
        image_format = img.format  # Get the format

        if image_format is None:
            raise ValueError("无法识别图像格式")

        image_format = image_format.lower()

        if image_format == 'png':
            return f"data:image/png;base64,{base64_image}"
        elif image_format in ['jpeg', 'jpg']:
            return f"data:image/jpeg;base64,{base64_image}"
        elif image_format == 'webp':
            return f"data:image/webp;base64,{base64_image}"
        else:
            raise ValueError(f"Unsupported image format: {image_format}")

    except Exception as e:
        logging.error(f"处理图像异常: {str(e)}")
        raise