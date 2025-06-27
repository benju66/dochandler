# Name of the installer
Outfile "DocHandlerInstaller.exe"

# Set the default installation directory
InstallDir "$PROGRAMFILES\DocHandler"

# Request administrator privileges
RequestExecutionLevel admin

# Default section: Installing the application
Section "Install"
    SetOutPath "$INSTDIR"

    # ✅ Copy the application executable
    File /r "C:\document_handler\dist\DocHandler.exe"

    # ✅ Copy all required icons
    SetOutPath "$INSTDIR\resources\icons"
    File /r "C:\document_handler\resources\icons\main_application_icon.ico"
    File /r "C:\document_handler\resources\icons\shortcut_icon.ico"
    File /r "C:\document_handler\resources\icons\taskbar_icon.ico"

    # ✅ Create Start Menu shortcut with the correct icon
    CreateDirectory "$SMPROGRAMS\DocHandler"
    CreateShortCut "$SMPROGRAMS\DocHandler\DocHandler.lnk" "$INSTDIR\DocHandler.exe" "$INSTDIR\resources\icons\shortcut_icon.ico"

    # ✅ Create Desktop shortcut with the correct icon
    CreateShortCut "$DESKTOP\DocHandler.lnk" "$INSTDIR\DocHandler.exe" "$INSTDIR\resources\icons\shortcut_icon.ico"

    # ✅ Write the uninstaller
    WriteUninstaller "$INSTDIR\Uninstall.exe"
SectionEnd

# Uninstall Section: Removing the application
Section "Uninstall"
    # ✅ Remove the application executable
    Delete "$INSTDIR\DocHandler.exe"

    # ✅ Remove Start Menu shortcut
    Delete "$SMPROGRAMS\DocHandler\DocHandler.lnk"

    # ✅ Remove Desktop shortcut
    Delete "$DESKTOP\DocHandler.lnk"

    # ✅ Remove the entire "resources/icons/" folder
    RMDir /r "$INSTDIR\resources\icons"

    # ✅ Remove the installation directory
    RMDir "$INSTDIR"
SectionEnd
