# pptx-translator

Python script that translates pptx files using Amazon Translate service.

## Installation

```
$ virtualenv venv
$ source venv/bin/activate
$ pip install -r requirements.txt
```

## Usage
```
$ python pptx-translator.py --help
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

## Security

See [CONTRIBUTING](CONTRIBUTING.md#security-issue-notifications) for more information.

## License

This library is licensed under the MIT-0 License. See the LICENSE file.
