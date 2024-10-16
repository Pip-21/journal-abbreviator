# Journal Abbreviator

This repository contains VBA code to abbreviate journal citations in Microsoft Word.

## Installation Instructions

### Prerequisites:
- **Microsoft Word** with VBA enabled.

### Steps to Install the Macro:
1. **Download Files**: 
   - Download both the `.bas` file (the macro code) and the `termlist.txt` (list of journal name abbreviations) from this repository.
   
2. **Import the Macro in Word**:
   - Open Microsoft Word.
   - Go to `View > Macros > View Macros`, then click on **Create** or **+**. This will open the **VBA Editor**.
   - In the VBA Editor, click `File > Import File` and locate the `.bas` file you downloaded.
   - The macro will be imported under a module named `Abbreviator`. You should see this module listed on the left pane in the VBA Editor.

3. **Configure the File Path**:
   - Open the `Abbreviator` module in the VBA Editor.
   - Find the line in the code that defines the file path for `termlist.txt`:
     ```vba
     filePath = "/path/to/your/termlist.txt"
     ```
   - Replace `"/path/to/your/termlist.txt"` with the actual file path where you saved `termlist.txt` on your system.

4. **Save and Close**:
   - Save your changes and close the VBA Editor.

## Usage Instructions

1. **Prepare Your Document**:
   - Open your Word document and select the bibliography or references section you want to abbreviate.

2. **Run the Macro**:
   - Go to `View > Macros > View Macros`, select `journal_abbreviator`, and click **Run**.

3. **Follow the Prompts**:
   - The macro will prompt you with confirmation messages. Click **OK** to proceed through the steps.
   - Once completed, your citations will be abbreviated based on the terms defined in `termlist.txt`.

## Notes
- Ensure that `termlist.txt` follows the correct format: each line should have the full journal name followed by its abbreviation, separated by a tab.
- The macro is optimized to replace journal names in a descending order of length, ensuring that longer names are replaced first to avoid partial replacements.
