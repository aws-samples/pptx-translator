import argparse
import boto3
import json
from botocore.exceptions import ClientError
from pptx import Presentation
from pptx.enum.lang import MSO_LANGUAGE_ID

LANGUAGE_CODE_TO_LANGUAGE_ID = {
    'af': MSO_LANGUAGE_ID.AFRIKAANS,
    'am': MSO_LANGUAGE_ID.AMHARIC,
    'ar': MSO_LANGUAGE_ID.ARABIC,
    'bg': MSO_LANGUAGE_ID.BULGARIAN,
    'bn': MSO_LANGUAGE_ID.BENGALI,
    'bs': MSO_LANGUAGE_ID.BOSNIAN,
    'cs': MSO_LANGUAGE_ID.CZECH,
    'da': MSO_LANGUAGE_ID.DANISH,
    'de': MSO_LANGUAGE_ID.GERMAN,
    'el': MSO_LANGUAGE_ID.GREEK,
    'en': MSO_LANGUAGE_ID.ENGLISH_US,
    'es': MSO_LANGUAGE_ID.SPANISH,
    'et': MSO_LANGUAGE_ID.ESTONIAN,
    'fi': MSO_LANGUAGE_ID.FINNISH,
    'fr': MSO_LANGUAGE_ID.FRENCH,
    'fr-CA': MSO_LANGUAGE_ID.FRENCH_CANADIAN,
    'ha': MSO_LANGUAGE_ID.HAUSA,
    'he': MSO_LANGUAGE_ID.HEBREW,
    'hi': MSO_LANGUAGE_ID.HINDI,
    'hr': MSO_LANGUAGE_ID.CROATIAN,
    'hu': MSO_LANGUAGE_ID.HUNGARIAN,
    'id': MSO_LANGUAGE_ID.INDONESIAN,
    'it': MSO_LANGUAGE_ID.ITALIAN,
    'ja': MSO_LANGUAGE_ID.JAPANESE,
    'ka': MSO_LANGUAGE_ID.GEORGIAN,
    'ko': MSO_LANGUAGE_ID.KOREAN,
    'lv': MSO_LANGUAGE_ID.LATVIAN,
    'ms': MSO_LANGUAGE_ID.MALAYSIAN,
    'nl': MSO_LANGUAGE_ID.DUTCH,
    'no': MSO_LANGUAGE_ID.NORWEGIAN_BOKMOL,
    'pl': MSO_LANGUAGE_ID.POLISH,
    'ps': MSO_LANGUAGE_ID.PASHTO,
    'pt': MSO_LANGUAGE_ID.BRAZILIAN_PORTUGUESE,
    'ro': MSO_LANGUAGE_ID.ROMANIAN,
    'ru': MSO_LANGUAGE_ID.RUSSIAN,
    'sk': MSO_LANGUAGE_ID.SLOVAK,
    'sl': MSO_LANGUAGE_ID.SLOVENIAN,
    'so': MSO_LANGUAGE_ID.SOMALI,
    'sq': MSO_LANGUAGE_ID.ALBANIAN,
    'sr': MSO_LANGUAGE_ID.SERBIAN_LATIN,
    'sv': MSO_LANGUAGE_ID.SWEDISH,
    'sw': MSO_LANGUAGE_ID.SWAHILI,
    'ta': MSO_LANGUAGE_ID.TAMIL,
    'th': MSO_LANGUAGE_ID.THAI,
    'tr': MSO_LANGUAGE_ID.TURKISH,
    'uk': MSO_LANGUAGE_ID.UKRAINIAN,
    'ur': MSO_LANGUAGE_ID.URDU,
    'vi': MSO_LANGUAGE_ID.VIETNAMESE,
    'zh': MSO_LANGUAGE_ID.CHINESE_SINGAPORE,
    'zh-TW': MSO_LANGUAGE_ID.CHINESE_HONG_KONG_SAR,
}

TERMINOLOGY_NAME = 'pptx-translator-terminology'

translate = boto3.client(service_name='translate')
bedrock = boto3.client(service_name='bedrock-runtime')

def generate_notes(slide_content, target_language_code):
    target_language_id = LANGUAGE_CODE_TO_LANGUAGE_ID.get(target_language_code, 'en')

    prompt = f"""Given the following slide content, generate concise and informative speaker notes. Provide the complete output in {target_language_id}.

Slide Content:
{slide_content}

Please provide your response in the following format, in {target_language_id}:

- [List 2-3 key points from the slide]

[2-3 sentences expanding on the key points, providing context or additional information]

Note: If {target_language_id} uses a non-Latin script, please provide the response in that script."""

    body = json.dumps(
        {
            "anthropic_version": "bedrock-2023-05-31",
            "max_tokens": 2048,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                    ],
                },
                {
                    "role": "assistant",
                    "content": ""
                }
            ],
        }
    )

    try:
        response = bedrock.invoke_model(
            modelId="anthropic.claude-3-5-sonnet-20240620-v1:0", body=body
        )
        response_body = json.loads(response["body"].read())
        return response_body["content"][0]["text"].strip()
    except Exception as e:
        print(f"Error generating notes: {str(e)}")
        return ""

def translate_presentation(presentation, source_language_code, target_language_code, terminology_names, overwrite_notes, add_missing_notes):
    slide_number = 1
    for slide in presentation.slides:
        print(f'Slide {slide_number} of {len(presentation.slides)}')
        slide_number += 1

        # Translate slide content and collect translated text
        translated_slide_content = ""
        for shape in slide.shapes:
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        translated_text = translate_text_frame(cell.text_frame, source_language_code, target_language_code, terminology_names)
                        translated_slide_content += translated_text + " | "
                    translated_slide_content += "\n"
            elif shape.has_text_frame:
                translated_text = translate_text_frame(shape.text_frame, source_language_code, target_language_code, terminology_names)
                translated_slide_content += translated_text + "\n"

        # Handle notes
        if not slide.has_notes_slide:
            slide.notes_slide = slide.add_notes_slide()
        
        notes_slide = slide.notes_slide
        notes_text = notes_slide.notes_text_frame.text.strip()

        if overwrite_notes or (add_missing_notes and not notes_text):
            # Generate new notes based on translated content
            generated_notes = generate_notes(translated_slide_content, target_language_code)
            notes_slide.notes_text_frame.text = generated_notes
        elif notes_text:
            # Just translate existing notes
            translate_text_frame(notes_slide.notes_text_frame, source_language_code, target_language_code, terminology_names)

    print("Translation and note generation completed.")

def translate_text_frame(text_frame, source_language_code, target_language_code, terminology_names):
    translated_text = ""
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            try:
                response = translate.translate_text(
                    Text=run.text,
                    SourceLanguageCode=source_language_code,
                    TargetLanguageCode=target_language_code,
                    TerminologyNames=terminology_names)
                run.text = response.get('TranslatedText')
                translated_text += run.text + " "
            except ClientError as client_error:
                if client_error.response['Error']['Code'] == 'ValidationException':
                    print('Invalid text. Ignoring...')
    return translated_text.strip()

def import_terminology(terminology_file_path):
    print(f'Importing terminology data from {terminology_file_path}...')
    with open(terminology_file_path, 'rb') as f:
        translate.import_terminology(Name=TERMINOLOGY_NAME,
                                     MergeStrategy='OVERWRITE',
                                     TerminologyData={'File': bytearray(f.read()), 'Format': 'CSV'})

def main():
    argument_parser = argparse.ArgumentParser(
            'Translates pptx files from source language to target language using Amazon Translate service')
    argument_parser.add_argument(
            'source_language_code', type=str,
            help='The language code for the language of the source text. Example: en')
    argument_parser.add_argument(
            'target_language_code', type=str,
            help='The language code requested for the language of the target text. Example: pt')
    argument_parser.add_argument(
            'input_file_path', type=str,
            help='The path of the pptx file that should be translated')
    argument_parser.add_argument(
            '--terminology', type=str,
            help='The path of the terminology CSV file')
    argument_parser.add_argument(
            '--overwrite-notes', action='store_true',
            help='Generate notes for all slides, overwriting existing notes')
    argument_parser.add_argument(
            '--add-missing-notes', action='store_true',
            help='Generate notes only for slides without existing notes')
    args = argument_parser.parse_args()

    print(args)

    terminology_names = []
    if args.terminology:
        import_terminology(args.terminology)
        terminology_names = [TERMINOLOGY_NAME]

    print(f'Translating {args.input_file_path} from {args.source_language_code} to {args.target_language_code}...')
    presentation = Presentation(args.input_file_path)
    translate_presentation(presentation,
                           args.source_language_code,
                           args.target_language_code,
                           terminology_names,
                           args.overwrite_notes,
                           args.add_missing_notes)

    output_file_path = args.input_file_path.replace(
            '.pptx', f'-{args.target_language_code}.pptx')
    print(f'Saving {output_file_path}...')
    presentation.save(output_file_path)

if __name__ == '__main__':
    main()