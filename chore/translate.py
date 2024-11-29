import os
import xml.etree.ElementTree as ET
from openai import OpenAI
import json

# Load the API key from environment variable
API_KEY = os.getenv("TEXTCRAFT_API_KEY")

if not API_KEY:
    raise EnvironmentError("Please set the TEXTCRAFT_API_KEY in your environment variables.")

client = OpenAI(api_key=API_KEY)

# Define the list of strings requiring translation
language_mapping = {
    "Afrikaans": "af",
    "Albanian": "sq",
    "Amharic": "am",
    "Arabic": "ar",
    "Armenian": "hy",
    "Assamese": "as",
    "Azerbaijani - Latin script": "az-Latn",
    "Bangla (India)": "bn-IN",
    "Basque": "eu",
    "Belarusian": "be",
    "Bosnian - Latin script": "bs-Latn",
    "Bulgarian": "bg",
    "Catalan": "ca",
    "Chinese (Simplified)": "zh-Hans",
    "Chinese (Traditional)": "zh-Hant",
    "Croatian": "hr",
    "Czech": "cs",
    "Danish": "da",
    "Dutch": "nl",
    "Estonian": "et",
    "Filipino": "fil",
    "Finnish": "fi",
    "French (Canada)": "fr-CA",
    "French (France)": "fr-FR",
    "Galician": "gl",
    "Georgian": "ka",
    "German": "de",
    "Greek": "el",
    "Gujarati": "gu",
    "Hebrew": "he",
    "Hindi": "hi",
    "Hungarian": "hu",
    "Icelandic": "is",
    "Indonesian": "id",
    "Irish Gaelic": "ga",
    "Italian": "it",
    "Japanese": "ja",
    "Kannada": "kn",
    "Kazakh": "kk",
    "Khmer": "km",
    "Konkani": "kok",
    "Korean": "ko",
    "Latvian": "lv",
    "Lithuanian": "lt",
    "Luxembourgish": "lb",
    "Macedonian (North Macedonia)": "mk",
    "Malay": "ms",
    "Malayalam": "ml",
    "Maltese": "mt",
    "Māori": "mi",
    "Marathi": "mr",
    "Nepali": "ne",
    "Norwegian (Bokmål)": "nb",
    "Norwegian (Nynorsk)": "nn",
    "Odia": "or",
    "Persian (Farsi)": "fa",
    "Polish": "pl",
    "Portuguese (Brazil)": "pt-BR",
    "Portuguese (Portugal)": "pt-PT",
    "Punjabi (India)": "pa-IN",
    "Romanian": "ro",
    "Russian": "ru",
    "Scottish Gaelic": "gd",
    "Serbian - Cyrillic script": "sr-Cyrl",
    "Serbian - Cyrillic script (Bosnia and Herzegovina)": "sr-Cyrl-BA",
    "Serbian - Latin script": "sr-Latn",
    "Slovak": "sk",
    "Slovenian": "sl",
    "Spanish (Mexico)": "es-MX",
    "Spanish (Spain)": "es-ES",
    "Swedish": "sv",
    "Tamil (India)": "ta-IN",
    "Tatar": "tt",
    "Telugu": "te",
    "Thai": "th",
    "Turkish": "tr",
    "Ukrainian": "uk",
    "Urdu": "ur",
    "Uyghur": "ug",
    "Uzbek - Latin script": "uz-Latn",
    "Valencian": "ca-ES-valencia",
    "Vietnamese": "vi",
    "Welsh": "cy",
}

# Updated resource files and their corresponding entries
resource_files = {
    "AboutBox": [
        "okButton.Text",
        "[AboutBox()] this.Text",
        "[AboutBox()] this.labelVersion.Text",
        "this.labelCopyright.Text",
        "this.labelCompanyName.Text",
        "$this.AccessibleDescription",
        "$this.Text",
    ],
    "Forge": [
        "this.ForgeTab.Label",
        "this.ToolsGroup.Label",
        "this.GenerateButton.Label",
        "this.GenerateButton.SuperTip",
        "this.DefaultCheckBox.Label",
        "this.AboutButton.Label",
        "this.AboutButton.ScreenTip",
        "this.CancelButton.Label",
        "this.CancelButton.ScreenTip",
        "this.WritingToolsGallery.Label",
        "this.WritingToolsGallery.SuperTip",
        "this.ReviewButton.Label",
        "this.ReviewButton.SuperTip",
        "this.ProofreadButton.Label",
        "this.ProofreadButton.SuperTip",
        "this.RewriteButton.Label",
        "this.RewriteButton.SuperTip",
        "this.SettingsGroup.Label",
        "this.SaveSettingsButton.Label",
        "this.ResetSettingsButton.Label",
        "this.RAGControlButton.Label",
        "this.RAGControlButton.SuperTip",
        "this.ModelListDropDown.Label",
        "this.ModelListDropDown.SuperTip",
        "this.DefaultCheckBox.SuperTip",
        "this.OptionsGroup.Label",
        "this.InfoGroup.Label",
        "(ThisAddIn.cs) [InitializeAddIn] ArgumentException #1",
        "[WritingToolsGallery_ButtonClick] ArgumentOutOfRangeException #1",
        "[ReviewButton_Click] MessageBox #1 (text)",
        "[ReviewButton_Click] MessageBox #1 (caption)",
        "(ModelProperties.cs) [GetContextLength] OllamaMissingContextWindowException #1",
        "(CommonUtils.cs) [GetInternetAccessPermission] MessageBox #1 Text",
        "(CommonUtils.cs) [GetInternetAccessPermission] MessageBox #1 Caption",
        "(WordMarkdown.cs) [ApplyMarkdownFormatting] ArgumentOutofRangeException #1",
        "(WordMarkdown.cs) [GetCodeBlockAtIndex] ApplicationException #1",
        "(WordMarkdown.cs) [ApplyImageFormatting] ArgumentException #1",
        "(WordMarkdown.cs) [ApplyHeadingFormatting] ArgumentException #1",
        "[AnalyzeText] InvalidRangeException #1",
        "this.CommentSystemPrompt",
        "(CommentHandler.cs) [AICommentReplyTask] UserChatMessage #1",
        "(CommentHandler.cs) [AICommentReplyTask] UserChatMessage #2",
        "[ProofreadButton_Click] SystemPrompt",
        "[ProofreadButton_Click] UserPrompt",
        "[RewriteButton_Click] SystemPrompt",
        "[RewriteButton_Click] UserPrompt",
        "[ReviewButton_Click] UserPrompt",
        "[Review] chatHistory #1",
        "(RAGControl.cs) [AskQuestion] chatHistory #1",
        "(RAGControl.cs) [AskQuestion] chatHistory #2",
        "(CommentHandler.cs) [AIUserMentionTask] UserMentionSystemPrompt",
    ],
    "GenerateUserControl": [
        "(GenerateUserControl.cs) _systemPrompt",
        "GenerateButton.Text",
        "[GenerateButton_Click] TextBoxEmptyException #1",
        "[GenerateButton_Click] TextBoxInvalidFormatException #2",
        "OutputLabel.Text",
        "PreviewButton.Text",
        "$this.Text",
    ],
    "PasswordPrompt": [
        "PasswordTextBox.AccessibleDescription",
        "PasswordTextBox.AccessibleName",
        "PasswordLabel.AccessibleDescription",
        "PasswordLabel.AccessibleName",
        "PasswordLabel.Text",
        "OkButton.AccessibleDescription",
        "OkButton.AccessibleName",
        "OkButton.Text",
        "CancelButton.Text",
        "$this.AccessibleDescription",
        "$this.Text",
    ],
    "RAGControl": [
        "FileListBox.AccessibleDescription",
        "FileListBox.AccessibleName",
        "AddButton.AccessibleDescription",
        "AddButton.AccessibleName",
        "AddButton.Text",
        "RemoveButton.AccessibleDescription",
        "RemoveButton.AccessibleName",
        "RemoveButton.Text",
        "[AddButton_Click] OpenFileDialog #1 Title",
        "[RemoveButton_Click] FileNotSelectedException #3",
        "[ReadPdfFileAsync] InvalidDataException #1",
        "SearchBar.PlaceholderText",
        "progressBar1.AccessibleDescription",
        "progressBar1.AccessibleName",
        "$this.AccessibleDescription",
        "$this.AccessibleName",
        "$this.Text",
    ],
}

# Base XML template
RESX_TEMPLATE = """<?xml version="1.0" encoding="utf-8"?>
<root>
  <xsd:schema id="root" xmlns="" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:msdata="urn:schemas-microsoft-com:xml-msdata">
    <xsd:import namespace="http://www.w3.org/XML/1998/namespace" />
    <xsd:element name="root" msdata:IsDataSet="true">
      <xsd:complexType>
        <xsd:choice maxOccurs="unbounded">
          <xsd:element name="metadata">
            <xsd:complexType>
              <xsd:sequence>
                <xsd:element name="value" type="xsd:string" minOccurs="0" />
              </xsd:sequence>
              <xsd:attribute name="name" use="required" type="xsd:string" />
              <xsd:attribute name="type" type="xsd:string" />
              <xsd:attribute name="mimetype" type="xsd:string" />
              <xsd:attribute ref="xml:space" />
            </xsd:complexType>
          </xsd:element>
          <xsd:element name="assembly">
            <xsd:complexType>
              <xsd:attribute name="alias" type="xsd:string" />
              <xsd:attribute name="name" type="xsd:string" />
            </xsd:complexType>
          </xsd:element>
          <xsd:element name="data">
            <xsd:complexType>
              <xsd:sequence>
                <xsd:element name="value" type="xsd:string" minOccurs="0" msdata:Ordinal="1" />
                <xsd:element name="comment" type="xsd:string" minOccurs="0" msdata:Ordinal="2" />
              </xsd:sequence>
              <xsd:attribute name="name" type="xsd:string" use="required" msdata:Ordinal="1" />
              <xsd:attribute name="type" type="xsd:string" msdata:Ordinal="3" />
              <xsd:attribute name="mimetype" type="xsd:string" msdata:Ordinal="4" />
              <xsd:attribute ref="xml:space" />
            </xsd:complexType>
          </xsd:element>
          <xsd:element name="resheader">
            <xsd:complexType>
              <xsd:sequence>
                <xsd:element name="value" type="xsd:string" minOccurs="0" msdata:Ordinal="1" />
              </xsd:sequence>
              <xsd:attribute name="name" type="xsd:string" use="required" />
            </xsd:complexType>
          </xsd:element>
        </xsd:choice>
      </xsd:complexType>
    </xsd:element>
  </xsd:schema>
  <resheader name="resmimetype">
    <value>text/microsoft-resx</value>
  </resheader>
  <resheader name="version">
    <value>2.0</value>
  </resheader>
  <resheader name="reader">
    <value>System.Resources.ResXResourceReader, System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089</value>
  </resheader>
  <resheader name="writer">
    <value>System.Resources.ResXResourceWriter, System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089</value>
  </resheader>
</root>
"""

# Translate multiple texts
def translate_batch(texts, target_language):
    # Add numbering to the texts
    prompt = (
        "Translate the following text into the target language. Each translated line must match the numbering and order of the input:\n\n"
        + "\n".join([f"{i + 1}. {text}" for i, text in enumerate(texts)])
        + f"\n\nTarget language: '{target_language}'"
    )

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                        "role": "system",
                        "content": (
                            "You are an advanced language translation model specializing in accurate and regionally appropriate translations. "
                            "Your task is to translate each input text into the specified target language, ensuring fidelity to the original meaning. "
                            "If the target language includes a localization (e.g., 'Spanish (Mexico)' or 'French (Canada)'), ensure the translation aligns with that regional variant. "
                            "The input will be provided as a numbered list, and your output must preserve this numbering and the exact order of the input texts. "
                            "Each translated line should correspond to the same numbered input line and should not include any additional commentary, explanation, or meta-information. "
                            "The output must be strictly formatted with each translation appearing on a new line, separated by '\\n', and matching the input order exactly. "
                            "For example:\n\n"
                            "Input:\n"
                            "1. Hello, how are you?\n"
                            "2. Thank you for your help.\n"
                            "Target language: 'Spanish (Mexico)'\n\n"
                            "Output:\n"
                            "1. Hola, ¿cómo estás?\n"
                            "2. Gracias por tu ayuda."
                        )
                },
                {"role": "user", "content": prompt},
            ],
        )

        # Process response to verify order
        translated_texts = response.choices[0].message.content.strip().split("\n")
        if len(translated_texts) != len(texts):
            raise ValueError("Mismatched translation count.")

        # Remove numbering and extra spaces from the output
        cleaned_translations = [
            " ".join(line.split(". ", 1)[1].split()) if ". " in line else " ".join(line.split())
            for line in translated_texts
        ]
        return cleaned_translations

    except Exception as e:
        print(f"Error during {target_language} translation: {e}")
        return []


# Main function
def generate_resx_files():
    base_dir = "../"

    # Parse RESX_TEMPLATE once
    resx_template_tree = ET.ElementTree(ET.fromstring(RESX_TEMPLATE))
    resx_template_root = resx_template_tree.getroot()

    # Group all translation tasks by language
    translations_by_language = {language: [] for language in language_mapping}
    key_file_map = {}  # Map to track (language, resource_file, value) to resource key association

    # Collect all text for translation
    for resource_file, keys in resource_files.items():
        base_file_path = os.path.join(base_dir, f"{resource_file}.resx")
        tree = ET.parse(base_file_path)
        root = tree.getroot()

        for data in root.findall("data"):
            name = data.attrib.get("name")
            if name in keys:
                value = data.find("value").text
                if value:  # Ensure there is a value to translate
                    for language, code in language_mapping.items():
                        translations_by_language[language].append((resource_file, value))
                        key_file_map[(language, resource_file, value)] = name

    # Perform batch translation for each language
    translated_results = {}
    for language, texts_to_translate in translations_by_language.items():
        if not texts_to_translate:
            print(f"No text to translate for {language}.")
            continue
        try:
            # Extract only the text values for translation
            texts = [value for _, value in texts_to_translate]
            translated_results[language] = translate_batch(texts, language)
        except Exception as e:
            print(f"Error during translation for {language}: {e}")
            translated_results[language] = []

    # Write translated RESX files
    for (language, resource_file, original_text), key in key_file_map.items():
        translated_texts = translated_results.get(language, [])
        if not translated_texts:
            print(f"Skipping {resource_file} for {language} due to missing translations.")
            continue

        code = language_mapping[language]
        translated_file_path = os.path.join(base_dir, f"{resource_file}.{code}.resx")

        # Read or create new RESX content
        if os.path.exists(translated_file_path):
            tree = ET.parse(translated_file_path)
            root = tree.getroot()
        else:
            # Clone the template root for a new file
            root = ET.Element("root")
            for child in resx_template_root:
                root.append(child)

        # Add translations
        texts_to_translate = [value for _, value in translations_by_language[language]]
        try:
            translated_index = texts_to_translate.index(original_text)
            translated_value = translated_results[language][translated_index]
        except ValueError:
            print(f"Original text '{original_text}' not found in translations for {language}.")
            continue

        # Ensure the key is unique in the file
        existing_data = root.find(f"./data[@name='{key}']")
        if existing_data is None:
            new_data = ET.SubElement(root, "data", {"name": key})
            ET.SubElement(new_data, "value").text = translated_value

        # Save the RESX file
        tree = ET.ElementTree(root)
        try:
            tree.write(translated_file_path, encoding="utf-8", xml_declaration=True)
            print(f"Created/Updated {resource_file}.{code}.resx")
        except Exception as save_error:
            print(f"Error saving {resource_file}.{code}.resx: {save_error}")


def add_resx_to_csproj():
    # Path configurations can still be defined locally or set globally
    csproj_path = os.path.join('../', 'TextCraft.csproj')
    resx_directory = '../'

    # Load the .csproj file as an XML tree
    tree = ET.parse(csproj_path)
    root = tree.getroot()
    
    # Define the XML namespace and register it
    namespace = {'msbuild': 'http://schemas.microsoft.com/developer/msbuild/2003'}
    ET.register_namespace('', 'http://schemas.microsoft.com/developer/msbuild/2003')

    # Find all existing EmbeddedResource elements
    embedded_resources = root.findall('.//msbuild:EmbeddedResource', namespace)
    existing_resx_files = {elem.attrib['Include'] for elem in embedded_resources}

    # Find all resx files in the specified directory
    resx_files = [f for f in os.listdir(resx_directory) if f.endswith('.resx')]

    # Filter the resx files to match the pattern "{resource_file}.{code}.resx"
    valid_resx_files = set()

    for resource_file in resource_files:
        for language_code in language_mapping.values():
            valid_resx_files.add(f"{resource_file}.{language_code}.resx")

    # For each resx file, check if it's already referenced in the .csproj file and matches the valid pattern
    for resx_file in resx_files:
        if resx_file not in existing_resx_files and resx_file in valid_resx_files:
            # Create a new EmbeddedResource element
            new_resource = ET.Element('EmbeddedResource', {'Include': resx_file})
            # Optionally add DependentUpon if applicable (assumes naming consistency)
            file_base_name = os.path.splitext(resx_file)[0].split('.')[0]  # Get the base name without language code
            ET.SubElement(new_resource, 'DependentUpon').text = f"{file_base_name}.cs"
            
            # Append the new element to the node (usually inside an ItemGroup)
            # Assuming the first ItemGroup is where resources are added
            item_group = root.find('msbuild:ItemGroup', namespace)
            if item_group is None:
                item_group = ET.SubElement(root, 'ItemGroup')
            item_group.append(new_resource)

    # Write the modified .csproj file back to disk
    tree.write(csproj_path, encoding='utf-8', xml_declaration=True)
    print(f"Updated {csproj_path} with new .resx files.")


if __name__ == "__main__":
    generate_resx_files()
    add_resx_to_csproj()
