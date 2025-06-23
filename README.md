# Advanced Replacer for Google Docs

An advanced, AI-powered batch find-and-replace add-on for Google Docs. This tool allows users to perform complex text replacements using a simple JSON format, with support for exact, fuzzy (similarity-based), and AI-driven matching.

![Screenshot of the Advanced Replacer sidebar in action](https://i.imgur.com/your-screenshot.png) 
*(Note: Replace with an actual screenshot URL after uploading one.)*

## ✨ Features

- **Multi-Level Matching**:
  - **Exact**: Standard, case-sensitive text replacement.
  - **Fuzzy**: Finds text with similar spelling or structure based on the Levenshtein distance.
  - **AI-Powered**: Uses `gpt-4o-mini` to find the best semantic match when exact or fuzzy searches fail.
- **Full Document Support**: Processes paragraphs, headings, lists, and tables.
- **Interactive Sidebar**: A modern, intuitive UI for managing replacements.
- **One-Click Undo**: Instantly revert the last batch of applied changes.
- **Real-Time Progress Bar**: Visual feedback for large operations, so you're never left guessing.
- **Safe & Secure**: All processing happens within your Google Account. The OpenAI API key is stored securely in your Script Properties.

## 🚀 Installation

1.  **Open Google Docs**: Go to the Google Doc you want to work with.
2.  **Open Apps Script**: Click on `Extensions` > `Apps Script`.
3.  **Copy the Code**:
    -   Delete any content in the default `Code.gs` file. Copy the entire content of `Code.gs` from this repository and paste it in.
    -   Click the `+` icon in the "Files" list and select `HTML`. Name the new file `Sidebar.html`.
    -   Delete the default content of `Sidebar.html`. Copy the entire content of `Sidebar.html` from this repository and paste it in.
4.  **Save the Project**: Click the "Save project" icon (💾).

## 🔧 Configuration (for AI Features)

To enable the AI matching feature, you need to add your OpenAI API key:

1.  **Open Project Settings**: In the Apps Script editor, click on the "Project Settings" icon (⚙️) on the left sidebar.
2.  **Add Script Property**: Scroll down to the "Script Properties" section and click "Add script property".
    -   **Property**: `OPENAI_API_KEY`
    -   **Value**: `sk-YourSecretApiKeyHere`
3.  **Save**: Click "Save script properties".

## 📖 How to Use

1.  **Open the Sidebar**: After installing, refresh your Google Doc. A new menu item `🚀 Advanced Replacer` will appear. Click it and select `Open Sidebar`.
2.  **Enter Directives**: In the sidebar's text area, paste a JSON array of "directives". Each directive is an object that specifies what to find and what to replace it with.

    **JSON Format:**
    ```json
    [
      {
        "fragment": "The old text to find.",
        "replaceWith": "The new text to insert."
      },
      {
        "fragment": "Another phrase to search for",
        "replaceWith": "Its replacement"
      }
    ]
    ```

3.  **Find Replacements**: Click the `🔍 Find Replacements` button. The script will scan your document and display a card for each potential change.
4.  **Review Suggestions**: Each card shows:
    -   The type of match (EXACT, FUZZY, or AI).
    -   A "diff" view of the proposed change.
    -   The element type (e.g., Paragraph, Heading).
5.  **Apply Changes**: Uncheck any suggestions you don't want to apply. Then, click `✅ Apply Changes`.
6.  **Undo (If Needed)**: If you're not happy with the result, click `↩️ Undo Last Run` to revert all the changes from that batch.

## 📜 License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details. 