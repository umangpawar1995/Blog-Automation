# LinkedIn Post Generator

This project automates the generation of professional LinkedIn posts and hero images using AI models via the OpenRouter API. It reads topics from an Excel calendar, generates a post and an image for each, and saves the results back to the Excel file.

## Features
- Reads topics, angles, and formats from `linkedin_posting_calendar.xlsx`
- Generates a LinkedIn post using GPT-3.5-turbo (or compatible) via OpenRouter
- Generates a hero image using Stable Diffusion XL (or compatible) via OpenRouter
- Saves generated content and image paths back to the Excel file
- Handles errors and falls back to a local placeholder image if needed

## Requirements
- Python 3.7+
- [openpyxl](https://pypi.org/project/openpyxl/)
- [requests](https://pypi.org/project/requests/)
- [Pillow](https://pypi.org/project/Pillow/)

Install dependencies:
```bash
pip install openpyxl requests pillow
```

## Setup
1. Place your `linkedin_posting_calendar.xlsx` in the project folder.
2. Set your OpenRouter API key in `generate_one_post.py`:
   ```python
   OPENROUTER_API_KEY = "sk-or-v1-..."
   ```
3. (Optional) Adjust other config variables as needed.

## Usage
Run the script to generate a post and image for the next pending topic:
```bash
python generate_one_post.py
```

- The script will update the Excel file with the generated post, image path, status, and timestamp.
- Images are saved in the `images/` directory.

## Excel File Format
Your `linkedin_posting_calendar.xlsx` should have at least these columns:
- `topic` (required)
- `angle` (optional)
- `format` (optional)
- `blog` (auto-filled)
- `image` (auto-filled)
- `status` (auto-filled)
- `generated_at` (auto-filled)

## Example
Suppose your Excel file contains:

| topic                | angle                | format   |
|----------------------|----------------------|----------|
| Data Engineering 101 | For beginners        | Listicle |
| AI in Healthcare     | Real-world examples  | Story    |

After running the script, the `blog`, `image`, `status`, and `generated_at` columns will be filled for the first unprocessed row.

## Notes
- If image generation fails, a local placeholder image is created.
- The script will skip rows already marked as `generated`.
- Make sure your API key is valid and has access to the required models.

## License
MIT License
