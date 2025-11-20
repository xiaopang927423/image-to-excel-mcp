"""
Service class for extracting tables from images and converting them to Excel.
"""
import json
import os
import sys
import time
import pandas as pd
from openai import OpenAI
import dotenv

# Add the src directory to the Python path to import modules
sys.path.append(os.path.join(os.path.dirname(__file__), os.pardir, os.pardir))

from src.utils.image_utils import encode_image_to_data_uri


class TableExtractor:
    """
    Service class to handle the process of extracting tables from images and converting them to Excel files.
    """
    
    def __init__(self):
        """Initialize the TableExtractor with API client and environment settings."""
        # Load environment variables
        if dotenv.load_dotenv(".env"):
            print("Loaded .env file successfully")
        else:
            print("Failed to load .env file")
        
        # Initialize OpenAI client
        self.client = OpenAI(
            api_key=os.getenv("api-key", ""),
            base_url=os.getenv("api-base")
        )
        
        # System prompt for table extraction
        self.sys_prompt = """
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

    def process_image_to_excel(self, image_path: str, output_path: str) -> str:
        """
        Process an image and convert the table in it to an Excel file.
        
        Args:
            image_path: Path to the input image file
            output_path: Output directory for the Excel file
            
        Returns:
            Path to the generated Excel file
        """
        if not image_path:
            return json.dumps({"error": "请上传图片"}, ensure_ascii=False)
        
        # Encode the image to data URI
        image_data_uri = encode_image_to_data_uri(image_path)
        
        # Extract table data from the image using the vision model
        table_info_md = self._extract_table_from_image(image_data_uri)
        
        if not table_info_md:
            raise ValueError("表格内容为空")
        
        # Parse the markdown table to a pandas DataFrame
        df = self._parse_markdown_table(table_info_md)
        
        # Generate output file path
        output_file_path = self._generate_output_path(output_path)
        
        # Create output directory if it doesn't exist
        os.makedirs(output_path, exist_ok=True)
        
        # Write DataFrame to Excel
        df.to_excel(output_file_path, index=False)
        
        return output_file_path

    def _extract_table_from_image(self, image_data_uri: str) -> str:
        """
        Extract table data from an image using the vision model.
        
        Args:
            image_data_uri: Data URI of the image to analyze
            
        Returns:
            Markdown formatted table string
        """
        try:
            completion = self.client.chat.completions.create(
                model="qwen-vl-max-latest",
                messages=[
                    {"role": "system", "content": self.sys_prompt},
                    {
                        "role": "user", 
                        "content": [
                            {"type": "image_url", "image_url": {"url": image_data_uri}},
                            {"type": "text", "text": "图中描绘的是什么景象?"}
                        ]
                    },
                ],
                stream=False,
            )
            return completion.choices[0].message.content
        except Exception as e:
            return json.dumps({"error": str(e)})

    def _parse_markdown_table(self, table_info_md: str) -> pd.DataFrame:
        """
        Parse markdown formatted table to a pandas DataFrame.
        
        Args:
            table_info_md: Markdown formatted table string
            
        Returns:
            Pandas DataFrame containing the table data
        """
        lines = table_info_md.strip().split('\n')
        
        # Parse each line of the markdown table
        data = []
        for line in lines:
            if line.startswith('|'):
                # Remove leading and trailing pipes and split by pipe
                row = [col.strip() for col in line.strip('|').split('|')]
                # Skip separator rows (lines with only dashes and spaces)
                if all(cell.strip('- ') == '' for cell in row):
                    continue
                data.append(row)
        
        if not data:
            raise ValueError("No valid table data found")
        
        # Create DataFrame with first row as headers
        headers = data[0]
        rows = data[1:] if len(data) > 1 else []
        
        return pd.DataFrame(rows, columns=headers)

    def _generate_output_path(self, output_path: str) -> str:
        """
        Generate a unique output file path with timestamp.
        
        Args:
            output_path: Base directory for the output file
            
        Returns:
            Full path to the output Excel file
        """
        timestamp = time.strftime("%Y%m%d%H%M%S")
        return os.path.join(output_path, f"{timestamp}.xlsx")