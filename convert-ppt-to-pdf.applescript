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
