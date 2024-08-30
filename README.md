# Convert PPT to PDF

Batch convert PowerPoint files to PDF using Automator on macOS

## News
- **August 30, 2024**: Pull Request Merged
	- A [PR#6](https://github.com/jeongwhanchoi/convert-ppt-to-pdf/pull/6) updating the convert-ppt-to-pdf.applescript has been merged. This update includes:
		- Improved compatibility with the new Mac OS.
		- Use of POSIX file paths to get appropriate new file paths.
		- Additional logging functionality for easier troubleshooting (commented out by default).

- **August 30, 2024**: README and Script Updated
	- We have updated both the README and the script in response to the macOS Sonoma compatibility issue (see [Issue#5](https://github.com/jeongwhanchoi/convert-ppt-to-pdf/issues/5)). The changes include:
		- A new script that addresses the alias handling problem in macOS Sonoma.

## How to make the Convert to PDF

1. Launch the Automator
2. Choose a type as *Quick Action* for your documents
3. Find Actions, 'Run AppleScript'
4. Type the AppleScript
   - Set the input setting
   - Type the script
5. Save the *Quick Action*

---

## Creating the Quick Action

### 1. Launch Automator
Open the Automator application on your Mac.

<img src="img/img-00.png" width="300">

### 2. Choose Quick Action
When creating a new document, select 'Quick Action' as the type for your workflow.

<img src="img/img-01.png" width="300">

### 3. Add Run AppleScript Action
In the library, search for and add the 'Run AppleScript' action to your workflow.

<img src="img/img-02.png" width="300">

### 4. Configure Input Settings
Set the workflow to receive input from 'documents' in 'Finder'.

<img src="img/img-03.png" width="300">

### 5. Enter AppleScript Code
Copy and paste the following AppleScript code ([link](https://github.com/jeongwhanchoi/convert-ppt-to-pdf/blob/master/convert-ppt-to-pdf.applescript)) into the 'Run AppleScript' action:


```app
on run {input, parameters}
    set theOutput to {}
    -- set logFile to (POSIX path of (path to desktop folder)) & "conversion_log.txt" -- for logging
    -- do shell script "echo 'Starting conversion...' > " & logFile -- for logging

    tell application "Microsoft PowerPoint" -- work on version 15.15 or newer
        launch
        repeat with i in input
            set t to i as string
            -- do shell script "echo 'Processing: " & t & "' >> " & logFile -- for logging
            if t ends with ".ppt" or t ends with ".pptx" then
                set pdfPath to my makeNewPath(i)
                -- do shell script "echo 'Saving to: " & pdfPath & "' >> " & logFile -- for logging
                try
                    open i
                    set activePres to active presentation
                    -- Export the active presentation to PDF
                    save activePres in (POSIX file pdfPath) as save as PDF
                    close active presentation saving no
                    set end of theOutput to pdfPath
                    -- Log the path to the console
                    -- do shell script "echo 'Saved PDF to: " & pdfPath & "' >> " & logFile -- for logging
                on error errMsg
                    -- do shell script "echo 'Error: " & errMsg & "' >> " & logFile -- for logging
                end try
            end if
        end repeat
    end tell
    tell application "Microsoft PowerPoint" -- work on version 15.15 or newer
        quit
    end tell
    -- do shell script "echo 'Conversion completed.' >> " & logFile -- for logging
    return theOutput
end run

on makeNewPath(f)
    set t to f as string
    if t ends with ".pptx" then
        return (POSIX path of (text 1 thru -6 of t)) & ".pdf"
    else
        return (POSIX path of (text 1 thru -5 of t)) & ".pdf"
    end if
end makeNewPath
```

---

### 6. Save the Quick Action
Give your workflow a name (e.g., "Convert PPT to PDF") and save it.

## How to Use
1. In Finder, select the PowerPoint files you want to convert.
2. Right-click to open the context menu.
3. Navigate to Quick Actions > "Convert PPT to PDF".
4. Wait for the conversion process to complete.

<img src="img/img-05.png" width="300">

## Star History

[![Star History Chart](https://api.star-history.com/svg?repos=jeongwhanchoi/convert-ppt-to-pdf&type=Date)](https://star-history.com/#jeongwhanchoi/convert-ppt-to-pdf&Date)