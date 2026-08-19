function Populate-WordTable {
    param (
        [object]$doc,
        [string]$placeholder,
        [array]$inputData,
        [string[]]$headers
    )

    
    # Determine the cell reference based on the placeholder
    switch ($placeholder) {
        "AggregationTable" { $cellRow = 3; $cellCol = 1 }
        default            { $cellRow = 2; $cellCol = 1 }
    }

    # Find the table with the placeholder
    $table = $null
    foreach ($tbl in $doc.Tables) {
        try {
            # Check if the table has enough rows and columns
            if ($tbl.Rows.Count -ge $cellRow -and $tbl.Columns.Count -ge $cellCol) {
                $cellText = $tbl.Cell($cellRow, $cellCol).Range.Text
                if ($cellText -like "*{{$placeholder}}*") {
                    $table = $tbl
                    break
                }
            }
        } catch { }
    }

    if (-not $table) {
        Write-Warning "Table with placeholder '$placeholder' not found."
        return
    }

    # Populate the table with input data
    foreach ($entry in $inputData) {
        $newRow = $table.Rows.Add()
        for ($i = 0; $i -lt $headers.Count; $i++) {
            $value = $entry.$($headers[$i])
            $cell = $newRow.Cells.Item($i + 1)
            $cell.Range.Text = "$value"

            # Apply color if this is the last column and a Colour property exists
            if ($i -eq ($headers.Count - 1) -and $entry.PSObject.Properties["Colour"]) {
                if ($null -ne $entry.Colour) {
                    # Apply colour
                    $cell.Shading.BackgroundPatternColor = [string]$entry.Colour
                    
                } else {
                    # REQUIRED: reset index so Word accepts RGB
                    # Resetting so null colour check will reset colour to nothing
                    $cell.Shading.BackgroundPatternColorIndex = 0 

                }
            }
        }
    }


    # Remove the placeholder row
    $table.Rows.Item($cellRow).Delete()
}
function Populate-DuplicateSourceTables {
    param (
        [object]$doc,
        [array]$duplicateGroups,
        [string[]]$headers
    )

    # ---- Validate anchor ----
    if (-not $doc.Bookmarks.Exists("OtherSourcesDynamicAnchor")) {
        Write-Warning "Bookmark OtherSourcesDynamicAnchor not found"
        return
    }

    # ---- Initial insertion range ----
    $insertRange = $doc.Bookmarks.Item("OtherSourcesDynamicAnchor").Range
    $insertRange.Collapse(0)
    $insertRange.ListFormat.RemoveNumbers()

    # ---- Remove page/section breaks at anchor (once) ----
    while ($insertRange.Text -match "`f" -or
           $insertRange.Text -match "\x0b" -or
           $insertRange.Text -match "\x0c") {
        $insertRange.Delete()
    }

    # ---- Section numbering control ----
    # Assumes Sources section is "4" and Delimited File was 4.5
    $sectionBase  = 4
    $sectionIndex = 6   # First dynamic section becomes 4.6

    foreach ($group in $duplicateGroups) {

        # -----------------------------
        # Dynamic Heading (Heading 2, manual number)
        # -----------------------------
        $insertRange.InsertParagraphAfter()
        $insertRange.Collapse(0)

        $insertRange.Text = "$sectionBase.$sectionIndex $($group.Name)"
        $para = $insertRange.Paragraphs.Last
        $para.Style = "Heading 2"

        # Critical: do NOT let Word auto-number
        $para.OutlineLevel = 10                   # BodyText behaviour
        $para.Range.ListFormat.RemoveNumbers()
        $para.Range.ParagraphFormat.KeepWithNext = $true
        $para.Range.ParagraphFormat.KeepTogether = $true

        $insertRange.Collapse(0)

        # -----------------------------
        # Instructional note (Normal text)
        # -----------------------------
        $insertRange.InsertParagraphAfter()
        $insertRange.Collapse(0)

        $insertRange.Text = "Please update table & title style and refresh table of contents"

        $notePara = $insertRange.Paragraphs.Last
        $notePara.Style = "Normal"
        $notePara.Range.Font.Italic = $true
        $notePara.Range.Font.Color = 8421504   # Grey
        $notePara.OutlineLevel = 10                 # Body text (never numbered)
        $notePara.Range.ListFormat.RemoveNumbers()

        # Keep note with the table
        $notePara.Range.ParagraphFormat.KeepWithNext = $true
        $notePara.Range.ParagraphFormat.KeepTogether = $true

        $insertRange.Collapse(0)

        # -----------------------------
        # Table
        # -----------------------------
        $insertRange.InsertParagraphAfter()
        $insertRange.Collapse(0)

        $table = $doc.Tables.Add($insertRange, 1, $headers.Count)

        # Header row
        for ($i = 0; $i -lt $headers.Count; $i++) {
            $table.Cell(1, $i + 1).Range.Text = $headers[$i]
        }

        # Data rows
        foreach ($entry in $group.Group) {
            $row = $table.Rows.Add()
            for ($i = 0; $i -lt $headers.Count; $i++) {
                $cell = $row.Cells.Item($i + 1)
                $cell.Range.Text = $entry.$($headers[$i])

                if ($i -eq ($headers.Count - 1) -and
                    $entry.PSObject.Properties["Colour"]) {

                    if ($null -ne $entry.Colour) {
                        $cell.Shading.BackgroundPatternColor = [string]$entry.Colour
                    }
                    else {
                        $cell.Shading.BackgroundPatternColorIndex = 0
                    }
                }
            }
        }

        # -----------------------------
        # Reset table formatting (NO outline numbering)
        # -----------------------------
        $table.Rows.HeadingFormat = $false
        $table.Borders.Enable = 1
        $table.Range.ParagraphFormat.KeepTogether = $true

        foreach ($cell in $table.Range.Cells) {
            foreach ($p in $cell.Range.Paragraphs) {
                $p.Range.ListFormat.RemoveNumbers()
                $p.OutlineLevel = 10   # BodyText
                $p.Style = "Normal"
            }
        }

        # -----------------------------
        # Advance insertion point
        # -----------------------------
        $insertRange = $table.Range
        $insertRange.Collapse(0)
        $insertRange.InsertParagraphAfter()
        $insertRange.Collapse(0)

        $doc.Bookmarks.Add("OtherSourcesDynamicAnchor", $insertRange)

        $sectionIndex++
    }

    # -----------------------------
    # Renumber static "Other Sources"
    # -----------------------------
    foreach ($para in $doc.Paragraphs) {
        if ($para.Range.Text -match "Other Sources") {
            $para.Range.Text = "$sectionBase.$sectionIndex Other Sources"
            break
        }
    }

    # -----------------------------
    # Final cleanup: remove breaks before static section
    # -----------------------------
    while ($insertRange.Text -match "`f" -or
           $insertRange.Text -match "\x0b" -or
           $insertRange.Text -match "\x0c") {
        $insertRange.Delete()
    }

    # -----------------------------
    # Update TOC
    # -----------------------------
    foreach ($toc in $doc.TablesOfContents) {
        $toc.Update()
    }
}

function Create-Password-Policy-WordTables {
    param (
        [object]$doc,
        [array]$policyData
    )

    $placeholder = "{{PasswordPolicyTables}}"
    $range = $doc.Content
    $found = $range.Find.Execute($placeholder)

    if ($found) {
        $insertRange = $range.Duplicate
        $insertRange.Text = ""  # Clear the placeholder
    }

    foreach ($policy in $policyData) {
        $properties = $policy.PSObject.Properties

        # Create a new table with 17 rows and 2 columns
        $table = $doc.Tables.Add($insertRange, 17, 2)

        $row = 1
        foreach ($prop in $properties) {
            $table.Cell($row, 1).Range.Text = $prop.Name
            $table.Cell($row, 2).Range.Text = "$($prop.Value)"  # Safe string conversion
            $table.Cell($row, 1).Range.ParagraphFormat.SpaceAfter = 0
            $table.Cell($row, 2).Range.ParagraphFormat.SpaceAfter = 0
            $row++
        }

        # Format the table
        $headerRow = $table.Rows.Item(1)
        $headerRow.Shading.BackgroundPatternColor = 15987699
        # $headerRow.Shading.BackgroundPatternColor = [Microsoft.Office.Interop.Word.WdColor]::wdColorGray05
        $headerRow.Range.Font.Bold = $true
        $table.Borders.Enable = $true

        # Move the range after the inserted table
        $insertRange = $table.Range
        $insertRange.Collapse(0)  # 0 = wdCollapseEnd

        # Insert a paragraph and set its style to Normal
        $para = $insertRange.Paragraphs.Add()
        $para.Range.Style = "Normal"

        # Move the range to the end of the inserted paragraph
        $insertRange = $para.Range
        $insertRange.Collapse(0)

    }
}

function Replace-Placeholder-Text {
    param (
        [object]$range,
        [string]$placeholder,
        [string]$newText
    )

    try {
        $find = $range.Find
    
        # check if invaid objects exist
        if ($null -eq $find) {return}
    
        $find.ClearFormatting()
        $find.Text = $placeholder
        $find.Replacement.ClearFormatting()
        $find.Replacement.Text = $newText

        $find.Execute(
            [ref]$find.Text,
            $false,  # MatchCase
            $false,  # MatchWholeWord
            $false,  # MatchWildcards
            $false,  # MatchSoundsLike
            $false,  # MatchAllWordForms
            1,       # Forward
            1,       # Wrap (wdFindContinue)
            $false,  # Format
            [ref]$newText,
            2        # ReplaceAll (wdReplaceAll)
        )
    }
    catch {
        Write-Warning "Failed to replace text"
    }

}