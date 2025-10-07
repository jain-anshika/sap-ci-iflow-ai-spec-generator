# Instructions to Set Up and Run the Project

## 1. Clone the Repository

Open PowerShell and run:

```powershell
git clone https://github.com/jain-anshika/sap-ci-iflow-ai-spec-generator.git
cd sap-ci-iflow-ai-spec-generator
```

## 2. Create a Python Virtual Environment

Run the following command to create a virtual environment named `venv`:

```powershell
python -m venv venv
```

## 3. Activate the Virtual Environment

Run the following command to activate the environment:

```powershell
.\venv\Scripts\Activate
```

## 4. Install Dependencies

Install required packages using:

```powershell
pip install python-docx matplotlib networkx requests
```


## 5. Configure API Keys and Paths

Before running the script, open `config_file.json` in a text editor and update the following fields:

- `gemini_api_url`: The Gemini API endpoint URL (leave as default unless instructed otherwise).
- `gemini_api_key`: Your Gemini API key. Replace the value with your own key.
- `source_xml_path`: Full path to your source XML file.
- `target_docx_path`: Full path where you want the generated DOCX file to be saved.
- `groovy_scripts_folder`: (Optional) Path to the folder containing Groovy scripts, if used in your iFlow.

Example:
```json
{
	"gemini_api_url": "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent",
	"gemini_api_key": "YOUR_API_KEY_HERE",
	"source_xml_path": "C:/path/to/your/source.iflw",
	"target_docx_path": "C:/path/to/output.docx",
	"groovy_scripts_folder": "C:/path/to/groovy/scripts"
}
```

## 6. Run the Script

Execute the main script:

```powershell
python Dynamic_generate_iflow_spec_using_ai.py
```

---

**Note:**
- Make sure you have Python installed. You can check by running `python --version`.
- If you encounter any issues, ensure you are using PowerShell and have the necessary permissions.
