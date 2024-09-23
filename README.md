# pptx-translator

Python script that translates pptx files using Amazon Bedrock (Claude Sonet) or Amazon Translate.

## Installation

```
$ virtualenv venv
$ source venv/bin/activate
$ pip install -r requirements.txt
```

## Usage
```
$ python pptx-translator.py --help
usage: Translates pptx files from source language to target language using Amazon Translate service or Bedrock-based translation
       [-h] [--terminology TERMINOLOGY] [--use-bedrock]
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
  --use-bedrock         Use Bedrock-based translation with Claude Sonet model
```

### Using Amazon Translate
To translate a presentation using Amazon Translate:

```
python pptx-translator.py en es input.pptx
```

### Using Bedrock-based Translation
To translate a presentation using the Bedrock-based translation with Claude Sonet model:

```
python pptx-translator.py en es input.pptx --use-bedrock
```

This will use the Claude Sonet model to perform the translation, which may provide improved results for certain language pairs or content types. Note: Bedrock translation does not use the terminology file

## Security

See [CONTRIBUTING](CONTRIBUTING.md#security-issue-notifications) for more information.

## License

This library is licensed under the MIT-0 License. See the LICENSE file.
