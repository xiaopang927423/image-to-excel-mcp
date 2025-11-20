# Image to Excel MCP

Convert table images to Excel files using the Qwen-VL model and FastMCP server.

## Features

- Extracts tables from images using AI vision models
- Converts extracted data to Excel format
- Provides an MCP (Model Context Protocol) interface
- Modular architecture with clear separation of concerns

## Architecture

The project follows a modular architecture:

- `src/main.py`: Main application entry point
- `src/services/table_extractor.py`: Core business logic for table extraction
- `src/utils/image_utils.py`: Utility functions for image processing
- `src/models/`: (Future) Data models

## Requirements

- Python 3.9 or higher
- Dependencies listed in `pyproject.toml`

## Installation

1. Install dependencies:
   ```bash
   pip install .
   ```

2. Set up environment variables in a `.env` file:
   ```
   api-key=your_api_key
   api-base=your_api_base_url
   ```

## Usage

Run the MCP server:
```bash
python -m src.main
```

Or if installed as a package:
```bash
image-to-excel-mcp
```

The server will start on `http://0.0.0.0:8000` using SSE transport.

## Modules

### Main
- Initializes the MCP server
- Defines the main API endpoint

### Services
- `TableExtractor`: Handles the complete process of image-to-Excel conversion
- Uses OpenAI-compatible API for vision model calls
- Parses markdown tables to pandas DataFrames

### Utils
- `encode_image_to_data_uri`: Encodes images with proper format detection
- Supports PNG, JPEG, and WebP formats