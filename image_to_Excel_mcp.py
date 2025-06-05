import json
from math import e
import os
import re
import sys
from typing import Any
from fastmcp import FastMCP
from openai import OpenAI
import dotenv
import base64
from PIL import Image
from io import BytesIO
import time
import pandas as pd
import logging

file_path = os.path.abspath(os.path.join(__file__, os.pardir))

if dotenv.load_dotenv(".env"):
    print("Loaded .env file successfully")
else:
    print("Failed to load .env file")

mcp: FastMCP[Any] = FastMCP(name="image_to_Excel")


sys_prompt = """
你是一名图片信息提取处理员工，首先你需要根据图片里面的信息，设计表格结构，填写表格中的数据。
# 要求：
保证数据结构完整
# 示例：
1、表格结构：
| 列1 | 列2 | 列3 |
| --- | --- | --- |
| 数据1 | 数据2 | 数据3 |
# 要求：填写表格中的数据，数据正确
# 返回数据格式为```
| 列1 | 列2 | 列3 |
| --- | --- | --- |
| 数据1 | 数据2 | 数据3 |
```
# 只返回表格数据，不返回其他信息
✅正确实例
```
| 列1 | 列2 | 列3 |
| --- | --- | --- |
| 数据1 | 数据2 | 数据3 |
```
❌错误实例
我返回的表格是：
| 列1 | 列2 | 列3 |
| --- | --- | --- |
| 数据1 | 数据2 | 数据3 |

"""
def get_image_format_and_encode(image_path):
    # 判断图像格式并生成对应的数据URI
    with open(image_path, "rb") as image_file:
        base64_image = base64.b64encode(image_file.read()).decode("utf-8")

    try:
        img = Image.open(BytesIO(base64.b64decode(base64_image)))
        image_format = img.format  # 先取出 format

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
        logging.error(f"处理异常: {str(e)}")
        raise

@mcp.tool()
def image_to_Excel(image_path: str,output_path:str =file_path) -> str:
    """
    用户上传图片路径和生成的文件的父目录
    使用视觉模型提取表格中的数据，
    转化为表格
    最后返回一个Excel文件"""
    client = OpenAI(api_key=os.getenv("api-key",),base_url=os.getenv("api-base"))
    if not image_path:
        return json.dumps({"error": "请上传图片"}, ensure_ascii=False)
    
    image_data_uri = get_image_format_and_encode(image_path)
    try:
        completion = client.chat.completions.create(
            model="qwen-vl-max-latest",
            messages=[
                {"role": "system", "content": sys_prompt},
                {"role": "user", "content": [{"type": "image_url", "image_url": {"url": image_data_uri}},{"type": "text", "text": "图中描绘的是什么景象?"}]},
            ],
            stream=False,
        )
        table_info_md = completion.choices[0].message.content
    except Exception as e:
        return json.dumps({"error": str(e)})

    if not table_info_md:
        raise ValueError("表格内容为空")
    lines = table_info_md.strip().split('\n')

    try:
    # 解析每一行的内容
        data = []
        for line in lines:
            if line.startswith('|'):
                # 去除每行两边的竖线，并分割各列
                row = [col.strip() for col in line.strip('|').split('|')]
                # 过滤分隔行
                if all(cell.strip('-') == '' for cell in row):
                    continue
                data.append(row)

        # 使用pandas写入Excel
        output_file_path = output_path + "/" + time.strftime("%Y%m%d%H%M%S") + ".xlsx"
        if not os.path.exists(output_path):
            os.makedirs(output_path)

        df = pd.DataFrame(data[1:], columns=data[0])
        df.to_excel(output_file_path, index=False)
    except Exception as e:
            return json.dumps({"error": str(e)})

    
    return output_file_path

def main():
    mcp.run(transport="sse",host="0.0.0.0",port=8000)

if __name__ == "__main__":
    main()
