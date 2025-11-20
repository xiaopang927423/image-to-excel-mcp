"""
Main application entry point for the Image to Excel MCP service.
"""
import os
import sys
from typing import Any

# Add the src directory to the Python path to import modules
sys.path.append(os.path.join(os.path.dirname(__file__), os.pardir))

from fastmcp import FastMCP
from src.services.table_extractor import TableExtractor


# Initialize the MCP server
mcp: FastMCP[Any] = FastMCP(name="image_to_Excel")


def image_to_Excel(image_path: str, output_path: str = None) -> str:
    """
    Convert table image to Excel file.
    
    Args:
        image_path: Path to the input image file
        output_path: Output directory for the Excel file (optional, defaults to current directory)
    
    Returns:
        Path to the generated Excel file
    """
    if output_path is None:
        output_path = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))
    
    extractor = TableExtractor()
    return extractor.process_image_to_excel(image_path, output_path)


def main():
    """Run the MCP server."""
    mcp.tool(image_to_Excel)
    mcp.run(transport="sse", host="0.0.0.0", port=8000)


if __name__ == "__main__":
    main()