# pptx-translator

Python script that translates pptx files using Amazon Translate service and generates speaker notes using Amazon Bedrock.

## Installation

```bash
$ virtualenv venv
$ source venv/bin/activate
$ pip install -r requirements.txt
```

## Usage

Basic translation:
```bash
python pptx-translator.py source_language_code target_language_code input_file_path
```

Translation with note generation:
```bash
python pptx-translator.py source_language_code target_language_code input_file_path [--overwrite-notes | --add-missing-notes]
```

Example execution:
```bash
python pptx-translator.py ja en input-file.pptx --overwrite-notes
```

For more information on available options:
```bash
python pptx-translator.py --help
```

## Command-line Arguments

```
usage: Translates pptx files from source language to target language using Amazon Translate service
       [-h] [--terminology TERMINOLOGY] [--overwrite-notes] [--add-missing-notes]
       source_language_code target_language_code input_file_path

positional arguments:
  source_language_code  The language code for the language of the source text.
                        Example: en
  target_language_code  The language code requested for the language of the
                        target text. Example: pt
  input_file_path       The path of the pptx file that should be translated

optional arguments:
  -h, --help            show this help message and exit
  --terminology TERMINOLOGY
                        The path of the terminology CSV file
  --overwrite-notes     
                        Generate notes for all slides, overwriting existing notes
  --add-missing-notes   
                        Generate notes only for slides without existing notes
```

## Features

- Translates PowerPoint (.pptx) files from one language to another using Amazon Translate
- Supports custom terminology for translation
- Optionally generates speaker notes using Amazon Bedrock with Antrophic Claude 3.5 Sonnet
- Can either overwrite existing notes or add notes only to slides without any

## Note Generation Options

- `--overwrite-notes`: Generates new notes for all slides, replacing any existing notes.
- `--add-missing-notes`: Generates notes only for slides that don't have existing notes.

If neither option is specified, the script will only translate existing slide and notes without generating new ones.

## Security

See [CONTRIBUTING](CONTRIBUTING.md#security-issue-notifications) for more information.

## License

This library is licensed under the MIT-0 License. See the LICENSE file.