# pptx-translator

Python script that translates pptx files using Amazon Translate service.

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

Example execution:
```bash
python pptx-translator.py ja en input-file.pptx
```

For more information on available options:
```bash
python pptx-translator.py --help
```

## Command-line Arguments

```
usage: Translates pptx files from source language to target language using Amazon Translate service
       [-h] [--terminology TERMINOLOGY]
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
```

## Features

- Translates PowerPoint (.pptx) files from one language to another using Amazon Translate
- Supports custom terminology for translation

## Security

See [CONTRIBUTING](CONTRIBUTING.md#security-issue-notifications) for more information.

## License

This library is licensed under the MIT-0 License. See the LICENSE file.