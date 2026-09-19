# Gemini Development Rules for ETABS Column Rebar Auditor

1. **Do Not Auto-Build the Executable:** Do not trigger PyInstaller compiles automatically after modifying the codebase.
2. **Explicit Compilation Permission:** Always ask the engineer/user if they want to build/repackage the application into a standalone `.exe` after making changes.
3. **Default Saving Path:** The default PDF generation output folder should always be configured to match the directory path of the active loaded/selected `.EDB` ETABS model.
