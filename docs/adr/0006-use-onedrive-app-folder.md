# Use OneDrive App Folder for working copies

The demo will store **OneDrive作業コピー** files in the editor's OneDrive App Folder and request `Files.ReadWrite.AppFolder` rather than broad file access. This keeps OAuth consent narrow across personal and work or school OneDrive accounts, at the cost of not letting editors choose arbitrary OneDrive locations for the temporary Word files.
