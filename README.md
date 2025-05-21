# Autodesk Inventor iLogic Script: Заявки (Applications/Requests) Management

This repository contains an iLogic script (`Заявки.iLogic.vb`) designed for use within Autodesk Inventor. It provides a user interface to help automate the creation and management of a structured folder system for daily project-related files, referred to as "Заявки" (which translates from Russian to "Applications" or "Requests").

## Purpose

The primary goal of this script is to streamline the organization of project deliverables by:

*   Automating the creation of a consistent, date-based folder structure for daily "applications" or "requests".
*   Providing tools for managing associated files, particularly PDFs.
*   Helping users quickly navigate to relevant project folders.

## Features

The script presents a user interface with two main tabs.

### 1. Основные операции (Main Operations) Tab

This tab is focused on setting up and managing the primary folder structure:

*   **Target Directory Management:**
    *   **Current Path Display:** Shows the base path where the application folders will be created. Defaults to the directory of the current Inventor file. The crucial subfolder created is named `- Заявки`.
    *   **Выбрать... (Browse...):** Allows selection of a different base directory.
    *   **Авто (Auto):** Resets the base directory to the active Inventor document's location.
    *   **Status Label:** Indicates if the `- Заявки` subfolder exists in the current path and whether it will be created.

*   **Folder Creation Options:**
    *   **Predefined Folders:** Checkboxes to select standard subfolders to be created within the daily application folder:
        *   `[00] Чертежи` (Drawings)
        *   `[02] Лазерный стол` (Laser Table)
        *   `[06] Лазерный труборез` (Laser Pipe Cutter)
        *   `[08] Нестинг` (Nesting)
        *   `[08] Присадка` (Fitting/Doweling)
    *   **Select/Deselect All:** Buttons to quickly check or uncheck all predefined folder options.
    *   **Пользовательские папки (Custom Folders):** Up to six text fields to define additional custom subfolder names.

*   **Action Buttons:**
    *   **Создать (Create):**
        *   Creates the main `- Заявки` folder in the target directory if it doesn't already exist.
        *   Inside `- Заявки`, it creates a new subfolder named with the current date (format: `dd.MM.yyyy`).
        *   Populates this date-stamped folder with subfolders selected from the predefined options and any custom folders specified.
        *   **Conflict Handling:** If a folder for the current date already exists, a dialog prompts the user to:
            *   `Удалить всё и создать заново` (Delete all and create anew).
            *   `Дозаписать недостающие директории` (Append missing directories).
    *   **Закрыть (Close):** Exits the script's interface.
    *   **Открыть папку (Open Folder):** Opens the `- Заявки` directory in Windows Explorer. Enabled only if the folder exists.
    *   **Стереть всё (Delete All):** Deletes the entire `- Заявки` folder and all its contents after user confirmation. Enabled only if the folder exists.

*   **Directory Information Lists:**
    *   **Существующие папки (Existing Folders):** Displays a list of date-named (`dd.MM.yyyy`) subdirectories found within `- Заявки`. Highlights if an entry for the current date already exists.
    *   **Несвязанные папки (Unrelated Folders):** Lists any subdirectories inside `- Заявки` that do not conform to the `dd.MM.yyyy` naming pattern.

### 2. Доп. операции (Additional Operations) Tab

This tab provides tools for file management and folder verification:

*   **Операции с PDF (PDF Operations):**
    *   **Сканировать PDF (Scan PDF):**
        *   Scans the selected `baseDirectory` (and its subdirectories, excluding the `- Заявки` path) for PDF files.
        *   Displays statistics: total number of PDF files found and the count of files with duplicate names.
    *   **Переместить PDF (Move PDF):**
        *   Moves all PDF files found in the `baseDirectory` (excluding those already in `- Заявки`) to the `[00] Чертежи` (Drawings) subfolder within the current day's application folder (i.e., `baseDirectory/- Заявки/dd.MM.yyyy/[00] Чертежи`).
        *   If the target `[00] Чертежи` folder doesn't exist, it is created.
        *   **Duplicate Handling:** If a file with the same name already exists in the destination, the moved file is renamed with a numerical suffix (e.g., `filename_(01).pdf`, `filename_(02).pdf`).
        *   Displays statistics: total number of files moved and the count of duplicates encountered during the move.
        *   Prompts the user if the main application folder for the current date does not exist, suggesting its creation via the "Main Operations" tab.

*   **Проверка наполнения папок (Folder Content Check):**
    *   **Проверить папки (Check Folders):**
        *   Verifies if the application folder for the current date (e.g., `baseDirectory/- Заявки/dd.MM.yyyy/`) exists.
        *   If it exists, the script checks for any empty subdirectories within it.
        *   If empty subdirectories are found, a new window appears listing these empty folders.
        *   If all subdirectories contain files, or if the current date folder doesn't exist, an appropriate status message is displayed.

## How to Use

1.  **Open Autodesk Inventor.**
2.  **Access iLogic:** Go to the "Tools" tab and click on "iLogic Browser" or "Add Rule".
3.  **Add the Script as an External Rule (Recommended):**
    *   In the iLogic Browser, right-click on the "External Rules" tab and select "Add External Rule".
    *   Browse to and select the `Заявки.iLogic.vb` file.
    *   You can now run the script by right-clicking it in the External Rules list and selecting "Run Rule".
    *   Alternatively, you can create a button in the Inventor ribbon to trigger this external rule for easier access.
4.  **Or, Add as a Document Rule:**
    *   Open the specific Inventor document (part, assembly, or drawing) you want to associate this rule with.
    *   In the iLogic Browser (on the "Rules" or "Forms" tab for that document), right-click and select "Add Rule".
    *   Copy and paste the content of `Заявки.iLogic.vb` into the rule editor.
    *   Save the rule. You can run it from the iLogic Browser.
5.  **Using the Interface:**
    *   Once the script is run, the "Выбор опций" (Option Selection) window will appear.
    *   **Main Operations Tab:**
        *   Verify or select the desired `baseDirectory` where the `- Заявки` folder will be managed.
        *   Choose the standard subfolders (Чертежи, Лазерный стол, etc.) you need.
        *   Add any custom folder names.
        *   Click "Создать" (Create) to generate the folder structure for the current date.
        *   Use "Открыть папку" (Open Folder) or "Стереть всё" (Delete All) as needed.
    *   **Additional Operations Tab:**
        *   Use "Сканировать PDF" (Scan PDF) to check for PDFs in your `baseDirectory`.
        *   Use "Переместить PDF" (Move PDF) to organize found PDFs into the current day's "[00] Чертежи" folder.
        *   Use "Проверить папки" (Check Folders) to ensure no required daily subfolders are empty.

## Folder Structure

The script creates and manages a folder structure based on the following pattern:

```
[Your Selected Base Directory]/
└── - Заявки/
    └── [Current Date, e.g., dd.MM.yyyy]/
        ├── [00] Чертежи/
        ├── [02] Лазерный стол/
        ├── [06] Лазерный труборез/
        ├── [08] Нестинг/
        ├── [08] Присадка/
        ├── [Custom Folder 1]/
        └── [Custom Folder 2]/
            ... (and so on for selected and custom folders)
```

*   **`[Your Selected Base Directory]`**: The path you choose or the default path of the active Inventor document.
*   **`- Заявки`**: The main folder for all applications/requests.
*   **`[Current Date, e.g., dd.MM.yyyy]`**: A subfolder named with the date when the "Создать" (Create) button was used.
*   **`[Selected/Custom Folders]`**: The subfolders created based on your selections in the "Основные операции" tab. PDF files moved using the "Доп. операции" tab are typically placed into the `[00] Чертежи` folder for the respective day.

## Potential Enhancements

*   **Configuration File:** Store predefined folder names, UI text, or default paths in an external configuration file (e.g., XML, JSON) for easier customization without editing the script directly.
*   **Localization:** Add support for multiple languages in the UI.
*   **Advanced PDF Naming:** Implement more sophisticated renaming rules for moved PDFs (e.g., based on drawing properties if run within an Inventor drawing's context).
*   **Logging:** Add a log file to record actions performed by the script, such as folder creation, file movements, and errors.
*   **Batch Processing:** Allow selection of multiple dates for folder creation or PDF organization.
