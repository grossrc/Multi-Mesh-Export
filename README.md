# Multi Mesh Export – Fusion 360 Add-in
Batch-export multiple STL files from a Fusion 360 design in one operation.

![alt text](/resources/ScreenshotC.png)
![Multi Mesh Export Demo](/resources/demo.gif)

## Features
- **Batch Export:** Select multiple bodies or use "Select All" to export them at once.
- **Custom Names:** Edit the save name for each selected body directly in the dialog before exporting.
- **Mesh Quality:** Choose High, Medium, or Low refinement.
- **Persistent Save Location:** Remembers your last used export folder (defaults to Downloads).
- **Safe Overwrites:** Prompts to Overwrite, Skip, or Cancel if a file already exists.
- **Progress Tracking:** Includes a progress bar with cancel support.

## Installation (option A)
1. Copy the `MultiMeshExport` folder to your Fusion 360 Add-Ins directory:
   - **Windows:** `%APPDATA%\Autodesk\Autodesk Fusion 360\API\AddIns\`
   - **macOS:** `~/Library/Application Support/Autodesk/Autodesk Fusion 360/API/AddIns/`
2. In Fusion 360, go to **UTILITIES → Add-Ins**, select the **Add-Ins** tab, and click **Run** next to *Multi Mesh Export*. Check **Run on Startup** to load it automatically.

## Installation (option B)
1. Clone the repo and extract the folder.
2. In Fusion go to `Utilities > Add-Ins` and click the icon. It should look like this:
![alt text](/resources/ScreenshotB.png)
3. In the Add-in window, click the plus symbol. Then click `Script or add-in from device` and select the program folder.
4. Click the dropdown on utilities and run the program in your file that you want to export.
![alt text](/resources/ScreenshotA.png)