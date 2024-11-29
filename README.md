# TextCraft
TextCraft is your ultimate AI-powered companion for Microsoft® Word®, seamlessly blending tools like text generation and proofreading directly into your workspace. No internet? No problem. Built for offline use, TextCraft puts cutting-edge AI at your fingertips while keeping your privacy front and center—a smarter, sleeker alternative to Microsoft® Copilot™️.


https://github.com/user-attachments/assets/0e37b253-3f95-4ff2-ab3f-eba480df4c61


# Prerequisities
To install this application, ensure your system meets the following requirements:
1. Windows 10 version 20H2 or later (32/64 bit)
2. Microsoft®️ Office®️ 2010 or later


# Install
Here's how to install TextCraft, the Office® Add-In with powerful AI tools:
1. **Install** [Ollama™️](https://ollama.com/download).
2. **Pull** the language model of your choice, for example:
   - `ollama pull qwen2.5:1.5b`
4. **Pull** an embedding model of your choice, for example:
   - `ollama pull all-minilm`
6. **Download the appropriate setup file:**
    - For a 32-bit system, download [`TextCraft_x32.zip`](https://github.com/suncloudsmoon/TextCraft/releases/download/v1.0.6/TextCraft_x32.zip).
    - For a 64-bit system, download [`TextCraft_x64.zip`](https://github.com/suncloudsmoon/TextCraft/releases/download/v1.0.6/TextCraft_x64.zip).
7. **Extract the contents** of the downloaded zip file to a folder of your choice.
8. **Run** `setup.exe`: This will install any required dependencies for TextCraft, including .NET Framework® 4.8.1 and Visual Studio® 2010 Tools for Office Runtime.
9. **Run** `OfficeAddInSetup.msi` to install TextCraft.
10. **Open Microsoft Word**® to confirm that TextCraft has been successfully installed with its integrated AI tools.


# Support
This add-in is designed to work effortlessly with OpenAI-compatible API endpoints, supporting providers like OpenAI, Ollama, and others such as vLLM. If your provider follows the OpenAI API standard, it should integrate seamlessly. See the table below for more details:

| Provider   | Support                                      |
|------------|---------------------------------------------|
| OpenAI     | Fully supported                             |
| Ollama     | Fully supported                             |
| Others     | Generally supported; check provider details |

The add-in is fully compatible with all display languages supported by Office (except Quechua (Peru) due to compatibility reasons), ensuring a seamless experience for users across different language preferences. For a complete list of supported languages, visit [Microsoft's support page](https://support.microsoft.com/en-us/office/what-languages-is-office-available-in-26d30382-9fba-45dd-bf55-02ab03e2a7ec).


# Features
Here’s a list of user environment variables you can adjust to tailor the settings to your needs.

| Environment Key              | Example Value                    | Description                                                                                 |
|------------------------------|-----------------------------------|---------------------------------------------------------------------------------------------|
| `TEXTCRAFT_OPENAI_ENDPOINT`  | `https://api.openai.com/v1`      | Specifies the API endpoint for TextCraft. Defaults to Ollama (`http://localhost:11434/v1`) if not specified. |
| `TEXTCRAFT_API_KEY`          | `dummy_key`                      | Sets the API key to authenticate with the API endpoint.                                     |
| `TEXTCRAFT_EMBED_MODEL`      | `all-minilm:latest`              | Specifies the embedding model to use in TextCraft.                                          |
| `TEXTCRAFT_GENERATE_PROMPT`  | `You are an AI assistant...`             | Allows you to customize the system prompt for the Generate feature, tailoring it to your needs. |
| `TEXTCRAFT_REVIEW_PROMPT`    | `You are an expert writing assistant...`             | Lets you override the system prompt for the Review feature, ensuring it aligns with your goals. |
| `TEXTCRAFT_PROOFREAD_PROMPT` | `You are a proofreading assistant...`             | Provides the flexibility to define your own system prompt for the Proofread feature.         |
| `TEXTCRAFT_REWRITE_PROMPT`   | `You are an advanced language model...`             | Enables you to set a personalized system prompt for the Rewrite feature.                    |
| `TEXTCRAFT_COMMENT_PROMPT`   | `You are a AI assistant...`             | Controls the system prompt for AI mentions in comments. |

## Generate
Click the "Generate" button to open a task pane with a text box and a "Generate" button. Simply type your prompt in the text box (plaintext only) and press CTRL+Enter or click the button to generate a response. The AI will use any available RAG context and the contents of your Word document to craft its reply. Once the response is ready, it will be seamlessly converted from markdown into Word formatting.

## Writing Tools
Expand the "Writing Tools" menu to choose from three options: Review, Proofread, and Rewrite.
- Review: Select text in your document and click "Review" for AI feedback. It will provide suggestions to improve your writing, acting like a supportive mentor.
- Proofread and Rewrite: Simply select text and let the AI refine or rework it for you.

## RAG Control
Use the RAG Control to enhance the AI’s understanding by adding PDFs with relevant context. This additional context is applied when using the Generate and Review features.

## Comments
Ask questions directly within your document by tagging the AI in a Word comment using "@". It’s intuitive and integrates seamlessly with [Word’s modern comments feature](https://support.microsoft.com/en-us/office/using-modern-comments-in-word-edc6ae71-0a2d-49fe-8faa-986f1e48136a), making collaboration and guidance easier than ever. For example, if you’ve selected the Qwen2.5:1.5b model in the dropdown, simply type "@qwen2.5:1.5b" in a comment, and the AI will reply right there, streamlining your workflow.

# Development
To customize and develop this Software, follow these steps:
1. Download and run the [Visual Studio® 2022 installer](https://visualstudio.microsoft.com/vs/).
2. When the installer opens, select both ".NET desktop development" and "Office/SharePoint development." When ".NET desktop development" is selected, under "Installation details" on the right, scroll to the "Optional" section and ensure the ".NET Framework 4.8.1 development tools" checkbox is selected.
3. Click "Install."
4. Clone this repository using Git.
5. Open the solution by double-clicking on the "TextForge.sln" file.
6. First, build the "TextCraft" project (in Release mode), then right-click on the "OfficeAddInSetup" project and select "Build."
7. Navigate to the project directory and find the "OfficeAddInSetup" folder, which contains the add-in setup files.
8. Be sure to clean the "TextCraft" project in both "Debug" and "Release" modes by right-clicking on the project in Visual Studio and selecting "Clean." This is necessary because building the project in Visual Studio, even without running the add-in installer, will automatically install the add-in in Microsoft Word. Cleaning is the only way to remove it.

> **Note**: For more information on VSTO (Visual Studio® Tools for Office®) development, you can refer to the [official Microsoft® documentation](https://learn.microsoft.com/en-us/visualstudio/vsto/walkthrough-creating-your-first-vsto-add-in-for-word?view=vs-2022&tabs=csharp).


# FAQ
1. **How do I uninstall the TextCraft add-in?**  
   - To uninstall TextCraft, go to **Settings > Apps > Apps & Features**, search for **TextCraft**, click **Uninstall**, and follow the prompts; if you manually installed it via Visual Studio®, clean the "TextCraft" project in both Debug and Release modes within Visual Studio®.
2. **How do I update the add-in**?
    - Simply uninstall the current version, download the latest one, and follow the installation steps.
2. **What should I do if I run into any issues?**
    - Don't worry! If you hit a snag, just [open an issue in this repository](https://github.com/suncloudsmoon/TextCraft/issues/new), and we'll be happy to help.


# Credits
1. [Ollama™️](https://github.com/ollama/ollama)
2. [OpenAI®️ .NET API library](https://github.com/openai/openai-dotnet)
3. [HyperVectorDB](https://github.com/suncloudsmoon/HyperVectorDB)
4. [PdfPig](https://github.com/UglyToad/PdfPig)
 
