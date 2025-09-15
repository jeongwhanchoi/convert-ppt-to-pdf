on run {input, parameters}
    set theOutput to {}

    tell application "Keynote"
        launch
        repeat with i in input
            set t to i as string
            if t ends with ".ppt" or t ends with ".pptx" then
                set pdfPath to my makePDFPath(i)
                try
                    -- Open the PPT/PPTX in Keynote
                    open i
                    my waitForOpen()

                    -- Export to PDF (Keynote's native export)
                    export document 1 to (POSIX file pdfPath) as PDF

                    -- Close without saving changes to .key
                    close document 1 saving no

                    set end of theOutput to pdfPath
                on error errMsg number errNum
                    try
                        if (count of documents) > 0 then close document 1 saving no
                    end try
                    -- optional: display dialog ("Export error: " & errMsg)
                end try
            end if
        end repeat
        quit
    end tell

    return theOutput
end run

on makePDFPath(f)
    set t to f as string
    if t ends with ".pptx" then
        return (POSIX path of (text 1 thru -6 of t)) & ".pdf"
    else
        return (POSIX path of (text 1 thru -5 of t)) & ".pdf"
    end if
end makePDFPath

on waitForOpen()
    tell application "Keynote"
        repeat 200 times
            if (count of documents) > 0 then exit repeat
            delay 0.1
        end repeat
    end tell
end waitForOpen
