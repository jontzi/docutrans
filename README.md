# Document Translator

Coded with the help of Claude.
I'm not professional so it might be total garbage.

A Python script that translates Microsoft Word documents using ChatGPT's translation capabilities via the deep-translator library. The script preserves document formatting while translating both regular text and table contents.

## Features

- Translates Word documents (.docx) from English to Finnish
- Preserves original document formatting
- Handles both paragraphs and table cells
- Includes progress tracking with tqdm
- Creates automatic backups during translation
- Supports debug mode for testing with limited pages
- Provides translation speed and statistics

## Prerequisites

- Python 3.6 or higher
- OpenAI API key

## Installation

1. Clone this repository:

```bash
git clone [your-repo-url]
cd document-translator
```

2. Install required dependencies:

```bash
pip install python-docx deep-translator python-dotenv tqdm
```

3. Create a `.env` file in the project root and add your OpenAI API key:

```
OPENAI_API_KEY=your_api_key_here
```

## Usage

1. Place your input Word document in the desired location
2. Update the input and output file paths in the script:

```python
input_file = "path/to/your/input/document.docx"
output_file = "path/to/your/output/document.docx"
```

3. Run the script:

```bash
python translate/deep-translator.py
```

## Configuration

- `DEBUG_MODE`: Set to `True` to test translation on a limited number of pages
- `DEBUG_PAGES`: Number of pages to translate in debug mode
- Source and target languages can be modified in the translator initialization:

```python
translator = ChatGptTranslator(api_key=api_key, source='en', target='fi')
```

## Features in Detail

### Progress Tracking

- Real-time progress bars for both paragraphs and tables
- Translation speed monitoring
- Total items translated counter

### Backup System

- Automatic saves every 10 translations
- Temporary files stored in `translate/temp` directory
- Prevents data loss during long translation sessions

### Statistics

- Total translation time
- Number of items translated
- Average translation speed

## Error Handling

- Validates OpenAI API key presence
- Creates necessary directories automatically
- Handles empty paragraphs and cells

## Limitations

- Currently supports English to Finnish translation
- Requires valid OpenAI API key
- Processing speed depends on API response times

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Acknowledgments

- Uses OpenAI's ChatGPT for translations
- Built with python-docx for document handling
- Uses deep-translator for translation services

## Token Estimator

The `translate/estimator.py` script helps estimate the token count and potential OpenAI API costs for processing DOCX files. It counts tokens in both paragraphs and table cells using tiktoken, OpenAI's tokenizer.

### Usage
1. Place your DOCX file in the input directory
2. Update the `input_file` path in the script to point to your document
3. Run the script to get:
   - Total text elements count
   - Non-empty text elements count
   - Estimated total tokens needed
   - Estimated cost in USD (based on gpt-3.5-turbo pricing)

### Dependencies
- openai
- python-docx
- tiktoken

The script adds a 50-token overhead for each non-empty text element to account for system messages and formatting.