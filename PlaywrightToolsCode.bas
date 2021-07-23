' =====================================================================
' Playwright Tools
' [2021-07-22]
'
' GitHub: https://github.com/pmbrennan/playwright-tools
'
' by Patrick M Brennan (pbrennan@gmail.com)
' Copyright (C) 2011-2021 Patrick M Brennan
' 
' This is a package for LibreOffice, consisting of a document template 
' and a set of macros intended to simplify the task of formatting
' stage plays, and to provide the playwright with an environment
' which promotes rapid writing. The package also provides an easy
' means to highlight individual character lines and stage directions,
' intended for use when printed scripts are prepared for readings or
' productions.
'
' NOTE: This package can also be used, with minor modifications,
' with OpenOffice (see the section called "USING THIS PACKAGE WITH
' OPENOFFICE").
'
' ********************************************************************
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License 
' along with this program (in the file LICENSE.txt).  If not, see
' <http://www.gnu.org/licenses/>.
' *********************************************************************
'
' THE PROBLEM
'
' Formatting a stage play is a nontrivial task. Time spent fussing
' with character slugs and stage directions is time which could be
' more profitably spent thinking about the characters and writing
' good dialogue.
' 
' From talking to playwrights about their work process, I have come
' to believe that many playwrights needlessly spend time doing
' rather pointless mechanical work: writing each character's entire
' name, every time the character has a line, and then manually
' centering and capitalizing it, before actually typing the
' line. This is not only an enormous waste of time, it also breaks
' the flow of the dialogue in the playwright's head, and may
' therefore actually impact the quality of the work.
'
' Good writing is about emotions, and it's hard to get that right
' if you're concerned with the minutiae of stage formatting.
' Fortunately, computers are really good at handling that kind of
' thing, leaving you free to write!
'
' Although it's true that there are specialized software packages
' which, at least to some extent, automate this drudgery, a
' surprisingly large proportion of playwrights do not use these
' packages, preferring instead to stick to less expensive, or less
' specialized options.
'
' This package allows a playwright to quickly write and annotate plain
' text with a few easy to type tags, which are subsequently transformed
' into proper stageplay format. 
'
' THE BENEFITS OF USING THIS PACKAGE
'
' 1. You can write faster, since you are typing fewer letters.
'
' 2. You can write with better "flow", since you are not interrupting
'    your writing to fiddle with formatting.
'
' 3. You can easily import text from multiple sources and quickly
'    format it into your script.
'
' 4. You can easily export text from stageplay format to plain text.
' 
' HOW TO USE THIS PACKAGE
' 
' You're a playwright. You want to write a script.
'
' You want your script to be formatted so that your lines look like 
' this:
'
'                     JOSEPH
'     This is my line. I am saying my line now.
'
' But you don't want to spend all your writing time typing names
' and formatting them to center on the line.  So just write your 
' script! Don't worry about formatting. Simply choose an easy-to-
' remember "tag" for each character. One or two or three letters 
' is plenty. Enclose your tag with slashes ("/"), and write your
' line thus:
'
' /j/ This is my line. I am saying my line now.
'
' Now press the "ApplyFormatting" button when you want to transform
' your text. The code will notice new tags (which you haven't
' previously associated with slugs) and ask you to supply the proper
' slug (JOSEPH in this example.) It will then save the tag/slug pair
' in a file called "tags.txt" in the same folder as your document,
' and then apply the formatting. Presto! All your lines are 
' formatted into proper stageplay style.
'
' You can continue writing with tags as long as you wish and press 
' "ApplyFormatting" any time you wish.
'
' You may also press "StripFormatting" to revert the script to tagged
' format at any time.
'
' There are 5 types of styles which are handled by this script:
'
' 1. CHARACTER LINES, which consist of a tag identifying the
'    speaker, followed by the words to be spoken. 
'    Example: The line
'
'        /a/ Hello, world!
'
'    will be formatted as (e.g.):
'
'                    AMBROSE
'    Hello, world!
'
'    Character lines are indicated by surrounding a custom character
'    tag with a slash at the beginning and end.
'
' 2. OVERLINE STAGE DIRECTIONS, which consist of a hint to the
'    actor about your overall intent for the line. These are sometimes
'    called "wrylies" (and are often abused).
'    Example: the line
'
'         /a/ [[Slyly]] Hello, world!
'
'    will be formatted as
'
'                    AMBROSE
'                    (Slyly)
'    Hello, world!
'
'    Overline stage directions are indicated by surrounding them with
'    the tags [[ and ]]. IMPORTANT: These will only work if they are at
'    the START of the line, right after the character tag.
'
' 3. INLINE STAGE DIRECTIONS, which consist of a hint to the actor
'    about your intent at a certain moment.
'    Example: the line
'
'    /a/ Hey there! (Winks) Come on over!
'
'    will be formatted as
'
'                    AMBROSE
'    Hey there! (Winks) Come on over!
'
'    The inline stage directions will be italicized.
'
' 4. STAGE DIRECTIONS BLOCK, which consist of a standalone
'    block of stage directions.
'    Example: the line
'
'    ## AT RISE: The stage is dark. The sounds of thunder and 
'    lightning may be heard.
'
'    will be formatted as
'
'                            AT RISE: The stage is dark. The sounds of
'                            thunder and lightning may be heard.
'
'    Stage directions blocks are indicated by beginning the paragraph 
'    with ##.
'
' 5. SIMPLE TEXT DECORATION, which includes bold, italic and underline
'    attributes of text.
'    Text which is surrounded by * will be bolded.
'    Text which is surrounded by _ will be underlined.
'    Text which is surrounded by \ will be italicized.
'
'    Example: the line
'
'    /a/ Hey there! *Be bold!* _Draw a line!_ \Italics?\
'
'    will be formatted as 
'
'                    AMBROSE
'    Hey there! Be bold! Draw a line! Italics?
'
'    with the appropriate sections decorated correctly.
'
' 6. CENTERED LINES, indicated by beginning the paragraph with @@.
'
' ************************************************************************
'
' HOW TO ENABLE THIS PACKAGE
'
' Before you can use this package for one of your scripts, you must
' enable LibreOffice to run these macros. In order to do this, follow the
' following steps:
'
' - Open LibreOffice.
' - Open your LibreOffice Preferences
'   - on Mac, it's LibreOffice > Preferences...
' - Select LibreOffice > Security
' - Press the "Macro Security..." button.
' - Select a Security Level for running Macros. My recommendation is
'   to set it at "Medium". This will make LibreOffice ask you if it's
'   OK to run macros every time you open your script, which I don't mind
'   doing.
'
' HOW TO USE THIS PACKAGE TO WRITE YOUR PLAYS
'
' Put a copy of this ODT file into its own directory. I recommned
' you use a different directory for every new play, as I do.
'
' Let's say I want to start a new play called "When Mary Met Sally". I'll
' start by creating a new folder called "When_Mary_Met_Sally". Then
' I'll copy this file into the folder and rename it as "MarySally.odt".
' Now when I open it up in LibreOffice, LibreOffice will ask me if it's
' OK to run macros and I'll click "Enable Macros". Now it's time to write.
'
' I'll write a few lines:
'
'   /m/ Hi, I'm Mary.
'
'   /s/ Hey Mary, I'm Sally.
'
'   /m/ Nice to meet you!
'
' And then I'll hit "ApplyFormatting". A dialog box will appear, saying,
' "I found an unknown tag '/m/'. Would you like to replace it?". I'll
' hit "Yes" and in the next dialog, I'll enter "Mary".
'
' Another dialog box will now appear saying, "I found an unknown tag '/s/'.
' Would you like to replace it?" I'll hit "Yes" and in the next dialog, I'll
' enter "Sally".
'
' Now my script will look like this:
'
'                                       MARY
'   Hi, I'm Mary.
'
'                                      SALLY
' 
'   Hey Mary, I'm Sally.
'
'
'                                       MARY
'   Nice to meet you!
'
' I can now continue to write, using the /m/ and /s/ tags as I wish,
' and every once in a while, when I want to see everything nicely
' formatted, I can press the "ApplyFormatting" button. Since LibreOffice
' already knows about /m/ and /s/ at this point, you won't be asked to
' supply the character names any more. And you'll notice there's a new file
' in your folder, called tags.txt, which looks like this (you'll need to
' open it up in a text editor):
'
' MARY,m
' SALLY,s
'
' This is so that the next time you open your file, you won't need
' to supply these names again.
'
' HOW TO USE THIS PACKAGE TO HIGHLIGHT YOUR SCRIPTS
'
' Another common task this package addresses is automatic highlighting.
' I find I spend a large amount of time preparing hardcopy scripts for
' script readings by manually marking each character's lines with a
' highlighter pen. In order to speed this up, I wrote a function to
' perform automatic highlighting. Here's how you use it.
'
' 1. Once you've applied your formatting, hit the "HighlightOptions"
'    button.
'
' 2. You'll be presented with a dialog which has three options:
'    - Remove All Highlights
'    - Highlight Stage Directions
'    - Highlight One Character's Lines
'
' 3. If you choose to highlight one character's lines, you will then
'    need to select the characters whose lines you wish to highlight.
'
' Therefore, to prepare for a reading, I will open my script, highlight
' the stage directions, and print. Then I will highlight one character's
' lines and print. I will repeat that step, highlighting each character's
' lines in turn, printing the script, until I have one copy for each
' character and a copy for stage directions. The highlight color is
' hardcoded as yellow, but if you look in the source code for
' "RGB(255,255,0)" (Yellow), you can change it to suit your own needs.
'
' ************************************************************************
'
' USING THIS PACKAGE WITH OPENOFFICE
'
' This package was originally designed to work with OpenOffice, but
' has since been modified to work only with LibreOffice, since
' LibreOffice is now the standard FOSS office package going forward.
' If, for some reason, you need to continue to use OpenOffice but
' you want to use this package, it's reasonably easy to modify so that
' you can do so. You can open this code up in your OpenOffice macro
' editor:
'
' - Menu item Tools > Macros > Organize Macros > LibreOffice BASIC...
' - Select "Macro From:" and click to expand your document.
' - click to expand "Standard" and then "Module1".
' - click on "Edit". This will open an editor window into this
'   package.
' - Now select Edit > Find & Replace...
' - Replace each instance of "Default Paragraph Style" (INCLUDE THE QUOTATION
'   MARKS!) with "Default" (AGAIN, INCLUDE THE QUOTATION MARKS!)
'   - This is because OpenOffice calls its default paragraph style
'     "Default", and LibreOffice calls it "Default Paragraph Style".
' - Save your work.
' - You can test by "round-tripping" a script text. If you start with
'   a file which is in tagged format and apply styling, then remove it,
'   you should end up with the same text. If the macro crashes with
'   an error message, you will know you have missed something.
'
' =====================================================================
' =====================================================================
'
' Still TODO:
'
' 1. Complete documentation of these functions.
'
' 2. Add more documentation from the LibreOffice wiki. Currently
'    there are a lot of references to the OpenOffice wiki, but
'    this support is now officially deprected.
'
' 3. Put the style names into the global tables rather than
'    sprinkling them into the functions.
'
' 4. Add support to create the proper styles all in code. Currently
'    these are:
'    CHARACTERNAME
'    LINE
'    STAGE_DIRECTION_BLOCK
'    STAGE_DIRECTION_OVERLINE
'    STAGE_DIRECTION_INLINE
'
' 5. One button editing of the tags file.
'    e.g. launch a text editor
'
' 6. Ability to pull up a dialog and enter a new slug/tag
'    combination and have it entered into the table.
'
' 7. Ability to delete a tags/slug combination.
'
' 8. Add automatic save/restore of the selection to all the buttons,
'    so that the user's cursor isn't changed by the action of the
'    macros.
'
' 9. Tweak the speech-break algorithm to favor breaking the speech
'    up at blank lines.
'
' =====================================================================
' =====================================================================
'
' COMPLETED TASKS
'
' 1. (DONE) Add support for overline directions, i.e. between the
'    character slug and the character's line.
'
' 2. (DONE) Add support for the user to specify the correspondence 
'    between tags and slugs, rather than forcing the user 
'    to do the mapping in this source file.
'    2a. See http://wiki.services.openoffice.org/wiki/Documentation/
'        BASIC_Guide/Files_and_Directories_(Runtime_Library)
'    this will allow the user to write a text file specifying the
'    mapping.
'
' 3. (DONE) One-button mapping and unnmapping.
'
' 4. (DONE) Replace all [[ and ]] in code with symbolic constants.
'
' 5. (DONE) Add support for stage directions inside the character's line.
'    After formatting, look for LINEs which have ( * ). Apply 
'    STAGE_DIRECTION_INLINE to the block between the ()s. 
'    Also perform the opposite transform.
'
' 6. (DONE) Add ability to write the tags.txt file out.
'
' 7. (DONE) Ability to detect new tags and interactively add these to the 
'    table.
'
' 8. (DONE) Ability to detect new slugs and interactively add these to the 
'    table.
'
' 9. (DONE) Ability to decorate text with *bold*, \italic\, and _underline_
'    markup.
'
' 10. (DONE) Ability to break up long speeches, with the correct (CONT'D)
'     tag supplied on the top of the following page.
'
' 11. (DONE) Ability to highlight any character's lines, or all stage
'     directions, as a convenience function for printing.
'
' 12. (DONE) Migrate this package from OpenOffice to LibreOffice.
'
' =====================================================================
' Global variables

Global Initialized as Integer  ' Have the tables been set up correctly?
Global NTags as Integer        ' How many tag/slug pairs are there?
Global Tags  as Variant        ' Array of the character tags
Global Slugs as Variant        ' Array of the character slugs
Global StageDirsTag as Variant ' Tag for Stage Directions blocks
Global CenteredLineTag as Variant ' Tag for centered paragraph
Global OverlineDirsOpenTag as Variant ' Open tag for overline dirs
Global OverlineDirsCloseTag as Variant ' Close tag for overline dirs
Global BoldOpenTag as Variant ' Open tag for bold
Global BoldCloseTag as Variant ' Close tag for bold
Global UnderlineOpenTag as Variant ' Open tag for underline
Global UnderlineCloseTag as Variant ' Close tag for underline
Global ItalicOpenTag as Variant ' Open tag for italic
Global ItalicCloseTag as Variant ' Close tag for italic
Global StrikethroughOpenTag as Variant ' Open tag for strikethrough
Global StrikethroughCloseTag as Variant ' Close tag for strikethrough
Global NLinesPerPage as Integer


' =====================================================================

Sub Main

End Sub

' =====================================================================
' Split
'
' Split up a comma-separated string into an array of strings
' example:
' Split ( "Fee, fi, fo, fum" ) will return
' [ "Fee", "fi", "fo", "fum" ]
Function Split( InString as String ) as Array

    Dim CommaPos as Integer
    Dim WorkString as String
    Dim Index as Integer
    Dim OutArray ( 2 ) as String
    
    Dim NFields as Integer
    
    ' Determine how many fields will be returned.
    CommaPos =  InStr ( 1 , InString, "," )
    NFields = 1
    Do While (CommaPos > 0) 
    
        CommaPos = InStr ( CommaPos + 1 , InString, "," )
        NFields = NFields + 1
        
    Loop

    If ( NFields > UBound ( OutArray )) Then
        Redim OutArray ( NFields )
    End If      
        
    WorkString = InString
    Index = 1
    
    Do While ( Len (WorkString) > 0 )
    
        CommaPos = InStr (1 , WorkString, "," )
        
        If ( CommaPos = 0 ) Then
            
            OutArray(Index) = Trim ( WorkString )
            WorkString = ""
            
        Else
        
            OutArray(Index) = Trim ( Left ( WorkString, CommaPos - 1))
            WorkString = Mid ( WorkString, CommaPos + 1 )
            
        End If
        
        Index = Index + 1
    
    Loop
    
    Set Split =  OutArray 

End Function

' =====================================================================
' HasDelimiterEnclosedText
'
' Determine if the given string contains text which begins with the
' OpenDelimiter and ends with the CloseDelimiter.
' If Strict is set, then the open delimiter MUST immediately precede
' a non-whitespace character, and the close delimiter MUST immediately 
' follow a non-whitespace character.
Function HasDelimiterEnclosedText ( InString as String, _
                                    OpenDelimiter as String, _
                                    CloseDelimiter as String, _
                                    Strict as Boolean ) _
                                    as Boolean
                                    
    Dim RetVal as Boolean
    Dim FirstTagPos as Integer
    Dim SecondTagPos as Integer
    
    RetVal = False
    
    FirstTagPos = InStr ( 1 , InString, OpenDelimiter)
    
    If FirstTagPos > 0 Then
    
        SecondTagPos = InStr ( FirstTagPos + Len(OpenDelimiter), InString, _
            CloseDelimiter )
        
        If SecondTagPos > 0 Then
        
            If SecondTagPos > FirstTagPos Then
            
                If Strict Then
                
                    ' TODO : Check here
                    RetVal = True
                    
                Else
            
                    RetVal = True
                    
                   End If
                
            End If
        
        End If
    
    End If
    
    HasDelimiterEnclosedText = RetVal                                   
                                    
End Function                                    

' =====================================================================
' HasOverlineStageDirection
'
' Determine if the given string contains an overline stage direction,
' i.e. contains a "[[" followed by zero or more characters, followed
' by a "]]"
Function HasOverlineStageDirection ( InString as String ) as Boolean
    
    HasOverlineStageDirection = HasDelimiterEnclosedText ( _
        InString, OverlineDirsOpenTag, OverlineDirsCloseTag, False )

End Function


' =====================================================================
' HasInlineStageDirection
'
' Determine if the given string contains an inline stage direction,
' i.e. contains a "(" followed by zero or more characters, followed
' by a ")"
Function HasInlineStageDirection ( InString as String ) as Boolean
    
    HasInlineStageDirection = HasDelimiterEnclosedText ( _
        InString, "(", ")", False )

End Function

' =====================================================================
' HasBold
'
' Determine if the given string contains bold markup.
Function HasBold ( InString as String ) as Boolean

    HasBold = HasDelimiterEnclosedText ( _
        InString, BoldOpenTag, BoldCloseTag, True )
        
End Function        

' =====================================================================
' StripParentheses
'
Function StripParentheses (InString as String) as String

    Dim OutString as String
    Dim Index as Integer
    
    OutString = Trim ( InString )
    
    Index = InStr ( 1, OutString, "(" )
    
    ' Must be the first character
    If (Index = 1) Then
    
        Index = InStr (1, OutString, ")")
        
        'Must be the last character
        If ( Index = Len(OutString) ) Then
        
            OutString = Trim ( Mid ( OutString, 2, Len(OutString) - 2 ) )
        
        End If
    
    End If
    
    StripParentheses = OutString 

End Function

' =====================================================================
' ReadTagsAndSlugs
' 
' Read the character tag/slug table.
' The file is in the same directory as the document, and it is called
' "tags.txt".
'
' The file is formatted as so:
'
' CHARACTER,tag
' e.g.
' KATRINA,k
' PAUL,p
'
' The function will return True if the file was found and successfully
' read, else it will return False.
'
Function ReadTagsAndSlugs as Boolean

    GlobalScope.BasicLibraries.loadLibrary("Tools")
    
    Dim FileNo As Integer
    
    Dim CurrentLine As String
    Dim File As String
    Dim Msg as String
    Dim Index as Integer
    
    Dim URLStr as String
    Dim Folder as String
    
    Dim InputList as Variant
    
    Dim Tag as String
    Dim Slug as String  
    
    Dim Success as Boolean
    
    Success = False
    
    URLStr = ThisComponent.getURL() 
    Folder = DirectoryNameoutofPath(URLStr, "/") 
     
    ' Define filename 
    ' TODO: file name better defined, customizable even
    Filename = Folder + "/tags.txt"
     
    ' Establish free file handle
    FileNo = Freefile
     
    If (not FileExists(Filename)) Then
    
        'msgBox "File " + Filename + " Cannot be found, sorry."
        Success = False
    
    Else
         
        ' Open file (reading mode)
        Open Filename For Input As FileNo
        
        Index = 1
         
        ' Check whether file end has been reached
        Do While not eof(FileNo)
          ' Read line 
          Line Input #FileNo, CurrentLine   
          If CurrentLine <> "" then
            InputList = Split ( CurrentLine )
            Slug = UCase ( InputList(1) )
            Tag = "/" + InputList(2) + "/ "
            
            Tags ( Index ) = Tag
            Slugs ( Index ) = Slug
            NTags = Index
            Index = Index + 1
            
            'Msg = "Tag = '" + Tag + "'" + vbNewLine + _
                '"Slug = '" + Slug + "'"
            'MsgBox Msg
          end if
        Loop
         
        ' Close file         
        Close #FileNo
        
        Success = True
    
    End If  
    
    ReadTagsAndSlugs = Success

End Function

' =====================================================================
' WriteTagsAndSlugs
' 
' Write the character tag/slug table to an external file.
'
' The file is in the same directory as the document, and it is called
' "tags.txt".
'
' The file is formatted as so:
'
' CHARACTER,tag
' e.g.
' KATRINA,k
' PAUL,p
'
Function WriteTagsAndSlugs as Boolean

    GlobalScope.BasicLibraries.loadLibrary("Tools")
    
    Dim FileNo As Integer
    
    Dim CurrentLine As String
    Dim File As String
    Dim Msg as String
    Dim Index as Integer
    
    Dim URLStr as String
    Dim Folder as String
    
    Dim Tag as String
    Dim Slug as String  
    
    Dim Success as Boolean
    
    Success = False
    
    URLStr = ThisComponent.getURL() 
    Folder = DirectoryNameoutofPath(URLStr, "/") 
     
    ' Define filename 
    ' TODO: file name better defined, customizable even
    Filename = Folder + "/tags.txt"
     
    ' Establish free file handle
    FileNo = Freefile
         
    ' Open file (writing mode)
    Open Filename For Output As FileNo
        
    For Index = 1 to NTags
    
        Tag = Mid( Tags ( Index ), 2, Len ( Tags ( Index ) ) - 3 )
        Slug = Left ( Slugs ( Index ), Len ( Slugs ( Index ) ) )
        
        'MsgBox ( "Tag = '" + Tag + "' Slug = '" + Slug + "'" )
        
        Print #FileNo, Slug + "," + Tag
            
    Next Index
         
    ' Close file         
    Close #FileNo
    
    Success = True
    
    WriteTagsAndSlugs = Success

End Function

' =====================================================================
' TagExists
'
' Determine whether a tag already exists in the table.
'
' The tag must be in the format "/xxx/ ".
'
Function TagExists ( Tag As String ) As Boolean

    Dim Index as Integer
    Dim Found as Boolean
    
    Found = False
    For Index = 1 to NTags
    
        If Tags ( Index ) = Tag Then
        
            Found = True
            Exit For
        
        End If
    
    Next Index
    
    TagExists = Found

End Function

' =====================================================================
' SlugExists
'
' Determine whether a slug already exists in the table.
'
' The slug must be in the format "XXXXX"
Function SlugExists ( Slug As String ) As Boolean

    Dim Index as Integer
    Dim Found as Boolean
    
    Found = False
    For Index = 1 to NTags
    
        If Slugs ( Index ) = Slug Then
        
            Found = True
            Exit For
        
        End If
    
    Next Index
    
    SlugExists = Found

End Function

' =====================================================================
' FindUnknownTags
'
' Find and report on tags which aren't in the tags table.
'
' NOTE: Currently the length of the tag is limited to 8 characters.
' This is an arbitrary limit and can easily be increased, but it 
' seems to me unnecessary.
'
' NOTE: Currently the tag can only include the following characters:
' letters, numbers, _ and -. I see no reason why new characters 
' can't be added, but I see no reason to add them, either.
'
' Returns True if any new tag/slug combinations have been defined; 
' False otherwise.
'
Function FindUnknownTags

    Dim Doc As Object
    Dim Cursor As Object
    Dim Choice As Integer
    Dim NewTag As String
    Dim NewSlug As String
    Dim Changed As Boolean
    
    Changed = False
    InitTables
    
    Doc = ThisComponent
    Cursor = Doc.Text.createTextCursor
    
    Dim SearchDesc As Object
    SearchDesc = Doc.createSearchDescriptor
    SearchDesc.searchString = "^\/([a-zA-Z0-9_\-]{1,8})\/ "
    SearchDesc.SearchBackwards = false
    SearchDesc.SearchRegularExpression = true
    SearchDesc.SearchCaseSensitive = true
    SearchDesc.SearchSimilarity = false
    
    Cursor = Doc.FindFirst(SearchDesc)
    
    Do Until IsNull(Cursor)
    
        If ( TagExists ( Cursor.String ) <> True ) Then
        
            Choice = MsgBox ( "I found an unknown tag '" + _
                              Cursor.String + "' " + Chr(13) + Chr(10) + _
                              "Would you like to replace it? " + _
                              Chr(13) + Chr(10) + _
                              "(Hit Cancel to exit this loop)", _
                              MB_YESNOCANCEL + MB_DEFBUTTON1 + _
                                  MB_ICONQUESTION, _
                              "Unknown Tag Found!" )
                            
            If Choice = IDCANCEL Then   
                Exit Do
            End If
            
            If Choice = IDYES Then
            
                NewSlug = InputBox ( "Enter the new slug for the tag '" + _
                                     Cursor.String + "' " + _
                                     "(or empty string to skip):" + _
                                     Chr(13) + Chr(10) + Chr(13) + Chr(10) + _
                                     "* Note: You do not need to " + _
                                     "CAPITALIZE the slug.", _
                                     "Enter new slug for tag " + _
                                     Cursor.String, "" )
                                     
                If (NewSlug <> "") Then
                
                    NewSlug = UCase ( NewSlug )
                    NewTag = Cursor.String
                    
                    If ( SlugExists ( NewSlug ) ) Then
                    
                        MsgBox ( "The Slug '" + NewSlug + _
                                 "' already exists. Skipping..." )
                        
                    Else
                    
                        NTags = NTags + 1
                        Tags ( NTags ) = NewTag
                        Slugs ( NTags ) = NewSlug
                        Changed = True
                        
                    End If                  
                
                End If              
            
            End If    

        End If
        
        Cursor = Doc.FindNext(Cursor.End, SearchDesc)
                    
    Loop
    
    If Changed Then
        ' DisplayTagsTable
        WriteTagsAndSlugs
    End If

    FindUnknownTags = Changed

End Function

' =====================================================================
' FindUnknownSlugs
'
' Find and report on slugs which aren't in the slugs table.
'
' NOTE: Currently the length of the slug is limited to 50 characters.
' This is an arbitrary limit and can easily be increased, but it 
' seems to me unnecessary.
'
' NOTE: Currently the slug can only include the following characters:
' letters, numbers, spaces, _ and -. I see no reason why new 
' characters can't be added, but I see no reason to add them, either.
'
' Returns True if any new tag/slug combinations have been defined; 
' False otherwise.
'
Function FindUnknownSlugs

    Dim Doc As Object
    Dim Cursor As Object
    Dim Choice As Integer
    Dim NewTag As String
    Dim NewSlug As String
    Dim Changed As Boolean
    
    Changed = False
    InitTables
    
    Doc = ThisComponent
    Cursor = Doc.Text.createTextCursor
    
    Dim SearchDesc As Object
    SearchDesc = Doc.createSearchDescriptor
    SearchDesc.searchString = "^[a-zA-Z0-9_\- \#]{1,50}$"
    SearchDesc.SearchBackwards = false
    SearchDesc.SearchRegularExpression = true
    SearchDesc.SearchCaseSensitive = true
    SearchDesc.SearchSimilarity = false
    
    Cursor = Doc.FindFirst(SearchDesc)
    
    Do Until IsNull(Cursor)
    
        NewSlug = UCase ( Cursor.String )
    
        If ( SlugExists ( NewSlug ) <> True ) Then
        
            Choice = MsgBox ( "I found an unknown slug '" + _
                              NewSlug + "' " + Chr(13) + Chr(10) + _
                              "Would you like to replace it? " + _
                              Chr(13) + Chr(10) + _
                              "(Hit Cancel to exit this loop)", _
                              MB_YESNOCANCEL + MB_DEFBUTTON1 + _
                                  MB_ICONQUESTION, _
                              "Unknown Slug Found!" )
                            
            If Choice = IDCANCEL Then   
                Exit Do
            End If
            
            If Choice = IDYES Then
            
                NewTag = InputBox ( "Enter the new tag for the slug '" + _
                                     Cursor.String + "' " + _
                                     "(or empty string to skip):" + _
                                     Chr(13) + Chr(10) + Chr(13) + Chr(10) + _
                                     "* Note: You do not need to " + _
                                     "add the leading and trailing /", _
                                     "Enter new tag for slug " + _
                                     NewSlug, "" )
                                     
                If (NewTag <> "") Then
                
                    NewTag = "/" + NewTag + "/ "
                    
                    If ( TagExists ( NewTag ) ) Then
                    
                        MsgBox ( "The Tag '" + NewTag + _
                                 "' already exists. Skipping..." )
                        
                    Else
                    
                        NTags = NTags + 1
                        Tags ( NTags ) = NewTag
                        Slugs ( NTags ) = NewSlug
                        Changed = True
                        
                    End If                  
                
                End If              
            
            End If    

        End If
        
        Cursor = Doc.FindNext(Cursor.End, SearchDesc)
                    
    Loop
    
    If Changed Then
        ' DisplayTagsTable
        WriteTagsAndSlugs
    End If

    FindUnknownSlugs = Changed

End Function

' =====================================================================
' DisplayTagsTable
'
Sub DisplayTagsTable

    Dim Index as Integer
    Dim Msg as String
    
    Msg = "NTags = " + NTags + Chr(13) + Chr(10)
    
    For Index = 1 to NTags
    
        Msg = Msg + "Tag : " + Tags(Index) + _
                    " Slug : " + Slugs ( Index ) + _
                    Chr(13) + Chr(10)
    
    Next Index
    
    MsgBox ( Msg, 0, "Tags Table" )

End Sub

' =====================================================================
' InitTables
'
' Properly set up the character tags and their
' corresponding slugs.
'
Sub InitTables

    If (Initialized = 0) Then
        ' msgBox ( "Variables not initialized..." )
        
        NTags = 0
        
        ' TODO: Figure out a better way to do this 
        ' than to dimension a huge-ass array here.
        ReDim Tags (100)
        ReDim Slugs (100)
        
        If (ReadTagsAndSlugs) Then
        
            'DisplayTagsTable
            
        End If
        
        StageDirsTag = "## " ' Block stage directions
        OverlineDirsOpenTag = "[[" ' Overline stage directions (between character slug and line)
        OverlineDirsCloseTag = "]]"
        BoldOpenTag = "*" ' Bold
        BoldCloseTag = "*"
        UnderlineOpenTag = "_" ' Underline
        UnderlineCloseTag = "_"
        ItalicOpenTag = "\" ' Italics
        ItalicCloseTag = "\"
        StrikethroughOpenTag = "-" ' Strikethrough
        StrikethroughCloseTag = "-"
        CenteredLineTag = "@@ " ' Center this line
        
        NLinesPerPage = 46
            
        Initialized = 1
            
    End If
    
End Sub 

' =====================================================================
' Reload
'
' Reset the tags/slugs table and reload.
Sub Reload
    Initialized = 0
    InitTables
End Sub

' =====================================================================
' ReplaceSlugWithTag
'
' Find each instance of a TargetString; 
' the target string must be all by itself on a paragraph.
' replace with the ReplaceString;
' Also collapse the hit paragraph and the following paragraph into
' one paragraph and apply the Default Paragraph Style to it.
' 
' The primary use of this method is to turn character-slug/line pairs 
' into single paragraphs wherein the line follows a character
' tag.
Sub ReplaceSlugWithTag ( TargetString as string, _
                         ReplaceString as string )
    Dim Doc As Object
    Dim Cursor As Object
    
    Doc = ThisComponent
    Cursor = Doc.Text.createTextCursor
    
    Dim SearchDesc As Object
    SearchDesc = Doc.createSearchDescriptor
    SearchDesc.searchString = "^" + TargetString + "$"
    SearchDesc.SearchBackwards = false
    SearchDesc.SearchRegularExpression = true
    SearchDesc.SearchCaseSensitive = false
    SearchDesc.SearchSimilarity = false
    
    Cursor = Doc.FindFirst(SearchDesc)
    
    Do Until IsNull(Cursor)
        
        Cursor.setString(ReplaceString)
        
        ' Now apply a new paragraph style
        Cursor.collapseToStart()
        Cursor.goRight(len(ReplaceString), false)
        Cursor.goRight(1, true)
        Cursor.setString("")
        Cursor.ParaStyleName = "Default Paragraph Style"
        
        Cursor = Doc.FindNext(Cursor.End, SearchDesc)           
    Loop

End Sub


' =====================================================================
' CollapseContdSlugs
'
' Find each instance of a slug ending in the string " (CONT'D)" 
' Collapse the hit paragraph into the previous paragraph.
' 
' This method is used to remove slugs which are generated by
' BreakUpLongSpeeches().
'
Sub CollapseContdSlugs (  )
    Dim Doc As Object
    Dim Cursor As Object
    Dim Hit As String
    Dim HitChr As Integer
    
    Doc = ThisComponent
    Cursor = Doc.Text.createTextCursor
    
    Dim SearchDesc As Object
    SearchDesc = Doc.createSearchDescriptor
    'SearchDesc.searchString = "^.* \(CONT\'D\):$"
    SearchDesc.searchString = "^.*\(CONT'D\)$"
    SearchDesc.SearchBackwards = false
    SearchDesc.SearchRegularExpression = true
    SearchDesc.SearchCaseSensitive = false
    SearchDesc.SearchSimilarity = false
    
    Cursor = Doc.FindFirst(SearchDesc)
    
    Do Until IsNull(Cursor)
        
        Hit = Cursor.GetString
        Cursor.setString("")
        
        Cursor.collapseToStart()
        Cursor.goRight(1, true)
        Cursor.setString("")
        
        Cursor.collapseToStart()
        Cursor.goLeft(1, true)        
        Hit = Cursor.GetString
        HitChr = ASC(Hit)
        If (HitChr = 10) Then
        	Cursor.setString("")
        End If
        
        Cursor.collapseToStart()
        Cursor.goLeft(1, true)        
        Hit = Cursor.GetString
        HitChr = ASC(Hit)
        If (HitChr = 10) Then
        	Cursor.setString(" ")
        End If
        
        'Cursor.ParaStyleName = "LINE"
        
        Cursor = Doc.FindNext(Cursor.End, SearchDesc)           
    Loop

End Sub


' =====================================================================
' ReplaceTagWithTextAndStyles
'
' Find each instance of the TargetString; replace with the ReplaceString; 
' also apply TargetStyle where the TargetString was found, insert a new 
' paragraph after that and apply NextParaStyle to the following
' paragraph.
' 
' The primary use of this method is to replace character name tags with 
' properly-formatted character slugs, followed by properly-formatted 
' lines.
'
Sub ReplaceTagWithTextAndStyles ( TargetString as string, _
                                  ReplaceString as string, _
                                  TargetStyle as string, _
                                  NextParaStyle as string)

    Dim Doc As Object
    Dim Cursor As Object
    
    Doc = ThisComponent
    Cursor = Doc.Text.createTextCursor
    
    Dim SearchDesc As Object
    SearchDesc = Doc.createSearchDescriptor
    SearchDesc.searchString = TargetString
    SearchDesc.SearchBackwards = false
    SearchDesc.SearchRegularExpression = false
    SearchDesc.SearchCaseSensitive = true
    SearchDesc.SearchSimilarity = false
    
    Cursor = Doc.FindFirst(SearchDesc)
    
    Do Until IsNull(Cursor)
        Cursor.setString(ReplaceString)
        Cursor.CharUnderline = com.sun.star.awt.FontUnderline.NONE
        Cursor.CharWeight = com.sun.star.awt.FontWeight.NORMAL
        Cursor.CharPosture = com.sun.star.awt.FontSlant.NONE
        
        ' Experimental
        'Cursor.CharBackColor = 16776960
        
        ' Now the hard part: apply a new paragraph style
        Cursor.collapseToStart()
        Cursor.goRight(len(ReplaceString), false)
        
        Doc.Text.insertControlCharacter(Cursor, _
            com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)  
        
        ' TODO: Make this style a parameter of the function
        If (Cursor.ParaStyleName <> "STAGE_DIRECTION_OVERLINE") then        
            Cursor.ParaStyleName = NextParaStyle
        End If
            
        Cursor.goLeft(len(ReplaceString), false)
        Cursor.ParaStyleName = TargetStyle
        
        Cursor = Doc.FindNext(Cursor.End, SearchDesc)           
    Loop

End Sub

' =====================================================================
' CharacterSlugs
'
' Replace all tags with proper character slugs and format the
' subsequent lines correctly.
Sub CharacterSlugs

    InitTables()
    
    For Index = 1 to NTags
        
        ' TODO: Make these styles into parameters
        ReplaceTagWithTextAndStyles(Tags(Index), Slugs(Index), _
            "CHARACTERNAME", "LINE")
        
    Next Index

End Sub

' =====================================================================
' CharacterTags
'
' Replace all character slug lines with the corresponding tags and 
' collapse them into a one-line representation of the character's 
' line.
Sub CharacterTags

    InitTables()

    For Index = 1 to NTags
        
        ReplaceSlugWithTag(Slugs(Index), Tags(Index))
        
    Next Index

End Sub

' =====================================================================
' FormatCenteredText
'
Sub FormatCenteredText

    Dim Doc As Object
    Dim Cursor As Object
    
    InitTables
    
    Doc = ThisComponent
    Cursor = Doc.Text.createTextCursor
    
    Dim SearchDesc As Object
    SearchDesc = Doc.createSearchDescriptor
    SearchDesc.searchString = "^" + CenteredLineTag
    SearchDesc.SearchBackwards = false
    SearchDesc.SearchRegularExpression = true
    SearchDesc.SearchCaseSensitive = false
    SearchDesc.SearchSimilarity = false
    
    Cursor = Doc.FindFirst(SearchDesc)
    
    Do Until IsNull(Cursor)
        
        Cursor.setString("")
        
        Cursor.collapseToStart()
        
        ' This code works but has undesirable side effects.
        Cursor.ParaStyleName = "CENTERED_BLOCK"
        
        Cursor = Doc.FindNext(Cursor.End, SearchDesc)           
    Loop
End Sub

' =====================================================================
' UnformatCenteredText
'
Sub UnformatCenteredText

    Dim Doc As Object
    Dim Enum1 As Object
    Dim TextElement As Object
    
    InitTables
     
    Doc = ThisComponent
    Enum1 = Doc.Text.createEnumeration
     
    ' loop over all paragraphs
    While Enum1.hasMoreElements
      TextElement = Enum1.nextElement
     
      If TextElement.supportsService("com.sun.star.text.Paragraph") Then
        
        If (TextElement.ParaStyleName = "CENTERED_BLOCK") Then
        	
           TextElement.ParaStyleName = "Default Paragraph Style"       
           TextElement.String = CenteredLineTag + TextElement.String  
        
        End If
     
      End If
    Wend
    
End Sub

' =====================================================================
' FormatStageDirectionBlocks
'
Sub FormatStageDirectionBlocks

    Dim Doc As Object
    Dim Cursor As Object
    
    InitTables
    
    Doc = ThisComponent
    Cursor = Doc.Text.createTextCursor
    
    Dim SearchDesc As Object
    SearchDesc = Doc.createSearchDescriptor
    SearchDesc.searchString = "^" + StageDirsTag
    SearchDesc.SearchBackwards = false
    SearchDesc.SearchRegularExpression = true
    SearchDesc.SearchCaseSensitive = false
    SearchDesc.SearchSimilarity = false
    
    Cursor = Doc.FindFirst(SearchDesc)
    
    Do Until IsNull(Cursor)
        
        Cursor.setString("")
        
        Cursor.collapseToStart()
        ' TODO: Make this style a parameter of the function
        Cursor.ParaStyleName = "STAGE_DIRECTION_BLOCK"
        
        Cursor = Doc.FindNext(Cursor.End, SearchDesc)           
    Loop

End Sub

' =====================================================================
' UnformatStageDirectionBlocks
'
Sub UnformatStageDirectionBlocks

    Dim Doc As Object
    Dim Enum1 As Object
    Dim TextElement As Object
    
    InitTables
     
    Doc = ThisComponent
    Enum1 = Doc.Text.createEnumeration
     
    ' loop over all paragraphs
    While Enum1.hasMoreElements
      TextElement = Enum1.nextElement
     
      If TextElement.supportsService("com.sun.star.text.Paragraph") Then
        
        If (TextElement.ParaStyleName = "STAGE_DIRECTION_BLOCK") Then
        	
            TextElement.ParaStyleName = "Default Paragraph Style"       
            TextElement.String = StageDirsTag + TextElement.String  
        
        End If
     
      End If
    Wend
    
End Sub

' =====================================================================
' UnformatManualBlockDirections
'
' Find paragraphs which have been manually formatted with a margin of 
' 2" or more, and unformat them, treating them as stage direction 
' blocks.
'
Sub UnformatManualBlockDirections

    Dim Doc As Object
    Dim Enum As Object
    Dim TextElement As Object
    
    InitTables
 
    Doc = ThisComponent
    Enum = Doc.Text.createEnumeration
 
    While Enum.hasMoreElements
        TextElement = Enum.nextElement
 
        If TextElement.supportsService("com.sun.star.text.Paragraph") Then
        
            StyleName = TextElement.ParaStyleName
            First20 = Left(TextElement.String, 20)              
            LeftMargin = TextElement.ParaLeftMargin
            
            If (StyleName <> "STAGE_DIRECTION_BLOCK") And (First20 <> "") _
                And (LeftMargin >= 5080) Then
            
                ' MsgBox "Para: '" + First20 + "' " + StyleName + " " + _
                ' LeftMargin
                
                TextElement.String = StageDirsTag + TextElement.String
                TextElement.ParaStyleName = "Default Paragraph Style"
                
            End If
        
        End If 
    Wend

End Sub

' =====================================================================
' FormatOverlineStageDirections
'
Sub FormatOverlineStageDirections

    Dim Doc As Object
    Dim Cursor As Object
    
    InitTables
    
    Doc = ThisComponent
    Cursor = Doc.Text.createTextCursor
    
    Dim SearchDesc As Object
    SearchDesc = Doc.createSearchDescriptor
    SearchDesc.searchString = OverlineDirsOpenTag
    SearchDesc.SearchBackwards = false
    SearchDesc.SearchRegularExpression = false
    SearchDesc.SearchCaseSensitive = false
    SearchDesc.SearchSimilarity = false
    
    Dim SearchDesc2 As Object
    SearchDesc2 = Doc.createSearchDescriptor
    SearchDesc2.searchString = OverlineDirsCloseTag
    SearchDesc2.SearchBackwards = false
    SearchDesc2.SearchRegularExpression = false
    SearchDesc2.SearchCaseSensitive = false
    SearchDesc2.SearchSimilarity = false
    
    Cursor = Doc.FindFirst(SearchDesc)
    
    Do Until IsNull(Cursor)
        
        Cursor.gotoEndOfParagraph(true)
        
        If HasOverlineStageDirection(Cursor.String) Then
        
            Cursor.collapseToStart()
            Cursor.goRight(Len(OverlineDirsOpenTag), true)
            Cursor.SetString("(")
            
            Cursor.collapseToEnd()
            Cursor = Doc.FindNext(Cursor.End, SearchDesc2)
            Cursor.SetString(")")
            Cursor.collapseToEnd()
            
            Doc.Text.insertControlCharacter(Cursor, _
                com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
                
            ' TODO: Make these styles into parameters
            Cursor.ParaStyleName = "LINE"
            Cursor.gotoEndOfParagraph(true)
            Cursor.SetString(Trim(Cursor.String))
            Cursor.collapseToStart()
            Cursor.goLeft(1, false)
            Cursor.ParaStyleName = "STAGE_DIRECTION_OVERLINE"
            Cursor.goRight(1, false)
            Cursor.gotoEndOfParagraph(false)
            
        End If
        
        Cursor = Doc.FindNext(Cursor.End, SearchDesc)
                
    Loop

End Sub

' =====================================================================
' UnformatOverlineStageDirections
'
Sub UnformatOverlineStageDirections

    Dim Doc As Object
    Dim Cursor As Object
    
    InitTables
    
    Doc = ThisComponent
    Cursor = Doc.Text.createTextCursor
    
    Dim SearchDesc As Object
    SearchDesc = Doc.createSearchDescriptor
    SearchDesc.searchString = "STAGE_DIRECTION_OVERLINE"
    SearchDesc.SearchBackwards = false
    SearchDesc.SearchRegularExpression = false
    SearchDesc.SearchCaseSensitive = false
    SearchDesc.SearchSimilarity = false
    SearchDesc.SearchStyles = true
    
    Cursor = Doc.FindFirst(SearchDesc)
    
    Do Until IsNull(Cursor)
        
        Cursor.SetString ( OverlineDirsOpenTag +  _
            StripParentheses ( Cursor.String ) +  _
            OverlineDirsCloseTag + " " )
        
        Cursor.gotoEndOfParagraph(false)        
        Cursor.goRight(1, true)
        Cursor.setString("")
        Cursor.ParaStyleName = "Default Paragraph Style"
        
        Cursor = Doc.FindNext(Cursor.End, SearchDesc)
                
    Loop

End Sub

' =====================================================================
' ApplyCharStyleToInlineStageDirections
'
Sub ApplyCharStyleToInlineStageDirections (inStyle as String)
	On Error Resume Next

    Dim Doc As Object
    Dim Cursor As Object
    Dim TextObj As Object
    
    InitTables
    
    Doc = ThisComponent
    Cursor = Doc.Text.createTextCursor
    
    Dim SearchDesc As Object
    SearchDesc = Doc.createSearchDescriptor
    SearchDesc.searchString = "\([^\(]*\)"
    SearchDesc.SearchBackwards = false
    SearchDesc.SearchRegularExpression = true
    SearchDesc.SearchCaseSensitive = false
    SearchDesc.SearchSimilarity = false
    SearchDesc.SearchStyles = false 
    
    Cursor = Doc.FindFirst(SearchDesc)
    
    Do Until IsNull(Cursor)
        
        If ( (Cursor.ParaStyleName = "LINE") ) Then
        
        	Cursor.CharStyleName = inStyle
        
        End If
        
        Cursor = Doc.FindNext(Cursor.End, SearchDesc)
                
    Loop

End Sub


' =====================================================================
' FormatInlineStageDirections
'
Sub FormatInlineStageDirections

    ApplyCharStyleToInlineStageDirections ( "STAGE_DIRECTION_INLINE" )
    
End Sub

' =====================================================================
' UnformatInlineStageDirections
'
Sub UnformatInlineStageDirections

    ApplyCharStyleToInlineStageDirections ( "Default Paragraph Style" )
    
End Sub

' =====================================================================
' ApplyTextDecorationFromMarkup
' 
' Replace markup with appropriate formatting
'
Sub ApplyTextDecorationFromMarkup ( OpenTag as String, CloseTag as String, _
                Posture as Integer, Underline as Integer, Weight as Integer, _
                Strikeout as Integer )

    Dim Doc As Object
    Dim Cursor As Object
    Dim TextLength As Integer
    
    Dim StartRange as Object
    Dim EndRange as Object
    Dim TextFormatCursor As Object
    
    InitTables
    
    Doc = ThisComponent
    Cursor = Doc.Text.createTextCursor
    TextFormatCursor = Doc.Text.createTextCursor
    
    Dim SearchDesc As Object
    SearchDesc = Doc.createSearchDescriptor
    SearchDesc.searchString = OpenTag
    SearchDesc.SearchBackwards = false
    SearchDesc.SearchRegularExpression = false
    SearchDesc.SearchCaseSensitive = false
    SearchDesc.SearchSimilarity = false
    
    Dim SearchDesc2 As Object
    SearchDesc2 = Doc.createSearchDescriptor
    SearchDesc2.searchString = CloseTag
    SearchDesc2.SearchBackwards = false
    SearchDesc2.SearchRegularExpression = false
    SearchDesc2.SearchCaseSensitive = false
    SearchDesc2.SearchSimilarity = false
    
    Cursor = Doc.FindFirst(SearchDesc)
    
    Do Until IsNull(Cursor)
        
        Cursor.gotoEndOfParagraph(true)
        
        If HasDelimiterEnclosedText ( _
            Cursor.String, OpenTag, CloseTag, True ) Then
        
            Cursor.collapseToStart() 
            
            Cursor.goRight(Len(OpenTag), true)
            Cursor.SetString("")
            
            Cursor.collapseToEnd()
            
            StartRange = Cursor.Start
            
            Cursor = Doc.FindNext(Cursor.Start, SearchDesc2)
                        
            Cursor.collapseToStart()
            Cursor.goRight(Len(CloseTag), true)
            
            Cursor.SetString("")
            Cursor.collapseToEnd()
            
            EndRange = Cursor.End
            
            TextFormatCursor.gotoRange(StartRange, False)
            TextFormatCursor.gotoRange(EndRange, True)
            
            If Posture > -1 Then
               TextFormatCursor.CharPosture = Posture 
            End If
               
            If Underline > -1 Then
               TextFormatCursor.CharUnderline = Underline
            End If
               
            If Weight > -1 Then
                TextFormatCursor.CharWeight = Weight
            End If
            
            If Strikeout > -1 Then
                TextFormatCursor.CharStrikeout = Strikeout
            End If
            
            'TextFormatCursor.CharPosture =  com.sun.star.awt.FontSlant.ITALIC    NONE
            'TextFormatCursor.CharUnderline = com.sun.star.awt.FontUnderline.SINGLE   NONE
            'TextFormatCursor.CharWeight = com.sun.star.awt.FontWeight.BOLD   NORMAL
            'TextFormatCursor.CharStrikeout = com.sun.star.awt.FontStrikeout.SINGLE ' SINGLE or NONE
            
        End If
        
        Cursor = Doc.FindNext(Cursor.End, SearchDesc)
                
    Loop
    
End Sub    

' =====================================================================
' ApplyBold
'
Sub ApplyBold

    InitTables
    ApplyTextDecorationFromMarkup ( BoldOpenTag, BoldCloseTag, -1, -1, _
        com.sun.star.awt.FontWeight.BOLD, -1  )

End Sub 

' =====================================================================
' ApplyUnderline
'
Sub ApplyUnderline

    InitTables
    ApplyTextDecorationFromMarkup ( UnderlineOpenTag, UnderlineCloseTag, _
        -1, com.sun.star.awt.FontUnderline.SINGLE, -1, -1  )

End Sub 

' =====================================================================
' ApplyItalic
'
Sub ApplyItalic

    InitTables
    ApplyTextDecorationFromMarkup ( ItalicOpenTag, ItalicCloseTag, _
        com.sun.star.awt.FontSlant.ITALIC, -1, -1, -1 )
        
End Sub 

' =====================================================================
' ApplyStrikeout
'
' TODO: Re-enable this when it is more robust.
Sub ApplyStrikeout

    InitTables
    ' ApplyTextDecorationFromMarkup ( StrikethroughOpenTag, _
    '    StrikethroughCloseTag, _
    '    -1, -1, -1, com.sun.star.awt.FontStrikeout.SINGLE )
        
End Sub       

' =====================================================================
' ReplaceTextDecorationWithMarkup
'
Sub ReplaceTextDecorationWithMarkup(DoBold As Boolean, DoItalic As Boolean, _
    DoUnderline As Boolean, DoStrikeout As Boolean)

    Dim Doc As Object
    Dim Cursor As Object
    Dim VCursor As Object
    
    Dim MyText As String
    Dim TextLen As Integer
    
    InitTables
    
    Dim NumberOfOptions As Integer
    Dim PageNum As Integer
    
    NumberOfOptions = 0
    If DoBold Then NumberOfOptions = NumberOfOptions + 1 
    If DoItalic Then NumberOfOptions = NumberOfOptions + 1 
    If DoUnderline Then NumberOfOptions = NumberOfOptions + 1
    If DoStrikeout Then NumberOfOptions = NumberOfOptions + 1 
    
    If NumberOfOptions <> 1 Then
        MsgBox "Wrong Number Of Options"
    End If
    
    Dim OpenTag As Variant
    Dim CloseTag As Variant
    
    Doc = ThisComponent
    Cursor = Doc.Text.createTextCursor
    VCursor = Doc.GetCurrentController.getViewCursor
    
    Dim SearchDesc As Object
    SearchDesc = Doc.createSearchDescriptor
    SearchDesc.searchString = ".*" '"[^     \n\r$]*" '"[A-Za-z]*"
    SearchDesc.SearchBackwards = false
    SearchDesc.SearchRegularExpression = true
    SearchDesc.SearchCaseSensitive = false
    SearchDesc.SearchSimilarity = false
    
    Dim SrchAttr(0) as new com.sun.star.beans.PropertyValue
    
    If DoBold Then
    
          SrchAttr(0).Name = "CharWeight"
          SrchAttr(0).Value = com.sun.star.awt.FontWeight.BOLD 
          OpenTag = BoldOpenTag
          CloseTag = BoldCloseTag
          
      ElseIf DoItalic Then
    
          SrchAttr(0).Name = "CharPosture"
          SrchAttr(0).Value = com.sun.star.awt.FontSlant.ITALIC
          OpenTag = ItalicOpenTag
          CloseTag = ItalicCloseTag
      
      ElseIf DoUnderline Then
    
          SrchAttr(0).Name = "CharUnderline" 
          SrchAttr(0).Value = com.sun.star.awt.FontUnderline.SINGLE
          OpenTag = UnderlineOpenTag
          CloseTag = UnderlineCloseTag
          
      ElseIf DoStrikeout Then
    
          SrchAttr(0).Name = "CharStrikeout" 
          SrchAttr(0).Value = com.sun.star.awt.FontStrikeout.SINGLE
          OpenTag = StrikethroughOpenTag
          CloseTag = StrikethroughCloseTag
      
      End If
      
      SearchDesc.SetSearchAttributes(SrchAttr())
    
    Cursor = Doc.FindFirst(SearchDesc)
    
    Do Until IsNull(Cursor)
    
        ' XXX: Experimental code. This is how to extract the page number 
        ' Of a block of text!
        VCursor.gotoRange(Cursor.Start, false)
        PageNum = VCursor.GetPage
        MyText = Cursor.GetString
        Cursor.SetString(OpenTag + MyText + CloseTag)
        
        If DoBold Then
            Cursor.CharWeight = com.sun.star.awt.FontWeight.NORMAL
        ElseIf DoItalic Then
            Cursor.CharPosture = com.sun.star.awt.FontSlant.NONE
        ElseIf DoUnderline Then
            Cursor.CharUnderline = com.sun.star.awt.FontUnderline.NONE
        ElseIf DoStrikeout Then
        	Cursor.CharStrikeout = com.sun.star.awt.FontStrikeout.NONE
        End If
        
        Cursor.collapseToEnd
        Cursor = Doc.FindNext(Cursor.End, SearchDesc)
        
    Loop    

End Sub

Sub ReplaceBoldWithMarkup
    ReplaceTextDecorationWithMarkup (True, False, False, False)
End Sub

Sub ReplaceItalicWithMarkup
    ReplaceTextDecorationWithMarkup (False, True, False, False)
End Sub

Sub ReplaceUnderlineWithMarkup    
    ReplaceTextDecorationWithMarkup (False, False, True, False)
End Sub 

Sub ReplaceStrikeoutWithMarkup    
    ReplaceTextDecorationWithMarkup (False, False, False, True)
End Sub                              

' =====================================================================
' ApplyFormatting
'
' Apply formatting for known tags, then scan for unknown tags and if
' there are any changes, apply formatting for them.
'
Sub ApplyFormatting

    CharacterSlugs
    
    If ( FindUnknownTags ) Then

       CharacterSlugs

    End If    
    
    FormatCenteredText
    FormatOverlineStageDirections   
    FormatStageDirectionBlocks
    FormatInlineStageDirections
    ApplyBold
    ApplyUnderline
    ApplyItalic
    ApplyStrikeout

End Sub

' =====================================================================
' StripFormatting
'
' Remove all relevant style information and apply tags to text.
'
Sub StripFormatting
    
    Dim Doc As Object
    Dim CurSelection As Object
    
    Doc = ThisComponent
    CurSelection = Doc.currentSelection
    
    CollapseContdSlugs
    Doc.LockControllers
    
    ReplaceStrikeoutWithMarkup
    ReplaceBoldWithMarkup
    ReplaceItalicWithMarkup
    ReplaceUnderlineWithMarkup
    UnformatInlineStageDirections
    UnformatStageDirectionBlocks
    UnformatOverlineStageDirections
    UnformatManualBlockDirections
    UnformatCenteredText
    CharacterTags
    
    Doc.CurrentController.Select(CurSelection)
    Doc.UnlockControllers

End Sub

' =====================================================================
' EditTags
' Launch an external editor to edit the tags file.
Sub EditTags

    Dim oSvc as object
    oSvc = createUnoService("com.sun.star.system.SystemShellExecute")

    oSvc.execute(ConvertToUrl("C:\windows\notepad.exe"), "tags.txt", 0)

End Sub

' =====================================================================
' CurrentLine
' Pretty much just a reference function now.
Function CurrentLine (oDoc)

   set mCurSelection = oDoc.currentSelection ' Record current selection
   ' Locking controller prevents flicker. But under ALL circumstances must be unlocked
   on local error goto finished:
   oDoc.lockControllers
   oVC = oDoc.getCurrentController.getViewCursor
   oVC.collapseToEnd  'Move view Cursor
   oVC.gotoStartofLine(true)
   nX = len(oVC.string)  'How many characters from the start of the line
   nY = 0
   'How many lines from top of page
   nPage = oVC.getPage
   while oVC.goUp(1,false) and oVC.getPage = nPage
       nY = nY + 1
   wend
   thisComponent.currentController.select(mCurSelection) 'Restore current selection
   finished:
   on error goto 0
   
   oDoc.unlockControllers
   CurrentLine = nY
   
end Function

' =====================================================================
' GetLineNumFromTextCursor
' 
' Given a text cursor, return the line number it is found on.
Function GetLineNumFromTextCursor (TCursor As TextCursor) As Integer

	Dim LineNum As Integer
	Dim PageNum As Integer
	Dim VCursor As Object
	VCursor = ThisComponent.getCurrentController.getViewCursor
	VCursor.gotoRange(TCursor, False)
	
	VCursor.collapseToEnd()
	VCursor.gotoStartOfLine(False)
	
	LineNum = 0
	PageNum = VCursor.getPage
	
	While (VCursor.goUp(1, False) And VCursor.getPage = PageNum)
		LineNum = LineNum + 1
	Wend
	
	GetLineNumFromTextCursor = LineNum
	
End Function

' =====================================================================
' GetPageNumFromTextCursor
'
' Given a text cursor, return the page number it is found on.
Function GetPageNumFromTextCursor (TCursor As TextCursor) As Integer

	Dim PageNum As Integer
	Dim VCursor As Object
	VCursor = ThisComponent.getCurrentController.getViewCursor
	VCursor.gotoRange(TCursor, False)
	
	VCursor.collapseToEnd()
	VCursor.gotoStartOfLine(False)
	
	PageNum = VCursor.getPage
	
	GetPageNumFromTextCursor = PageNum
	
End Function

' =====================================================================
' GetLengthOfSpeechFromTextCursor
'
' Given a text cursor assumed to be on a character slug, return the
' number of lines in the subsequent speech, defined as the number of 
' times that the visual cursor can be moved down and remain in a
' paragraph with the LINE or STAGE_DIRECTION_OVERLINE style.
Function GetLengthOfSpeechFromTextCursor (TCursor As TextCursor) As Integer

	Dim VCursor As Object
	Dim LocalTCursor As Object
	Dim LengthOfSpeech As Integer
	Dim MoveSuccessful As Boolean
	
	VCursor = ThisComponent.getCurrentController.getViewCursor
	LocalTCursor = ThisComponent.Text.createTextCursor
	VCursor.gotoRange(TCursor, False)

 	VCursor.gotoStartOfLine(false)
    LengthOfSpeech = 0
    MoveSuccessful = VCursor.goDown(1, false)
    While MoveSuccessful 
        LocalTCursor.gotoRange(VCursor, false)
        PStyle = LocalTCursor.ParaStyleName
        If (PStyle = "LINE" OR PStyle = "STAGE_DIRECTION_OVERLINE") Then
            LengthOfSpeech = LengthOfSpeech + 1
            MoveSuccessful = VCursor.goDown(1, false)
        Else
            MoveSuccessful = false
        End If      
    Wend
    
    GetLengthOfSpeechFromTextCursor = LengthOfSpeech

End Function

' =====================================================================
' BreakUpLongSpeeches
' NOTE: TextCursor API is here: http://wiki.services.openoffice.org/wiki/Text_cursor
'
' Notes for troubleshooting BreakupLongSpeeches:
' 1. The LINE style is marked for LibreOffice to try to keep the 
'    paragraph together, on the same page if possible, and keep it 
'    with its preceding CHARACTERNAME paragraph. Therefore, extremely
'    long speeches without paragraph breaks will initially appear on 
'    top of their own page, usually with very large areas of white space
'    on the preceding page. If you see this in your script,
'    BreakUpLongSpeeches is a good macro to invoke.
' 2. The macro will leave speeches of less than 16 lines alone, 
'    allowing LibreOffice to flow these on to another page and leaving
'    up to that many lines blank on your page.
' 3. The macro will also leave the speech alone if there are 8 or less
'    lines remaining on your page when the speech begins. 
' 3. If you see a lot of extra space (2-3 line feeds) between your 
'    "CONT'D" line and the continuation of the text, it just means that
'    the macro broke your speech in between paragraphs. It won't hurt to
'    just delete that extra space.
' 4. If you see a page begin in the middle of your long speech, without
'    a continuation line, this is because there was a line in the body
'    of the speech which doesn't have the "LINE" style applied to it.
'    The macro looks for the first line which isn't marked with that style,
'    in order to decide where the speech ends. To fix this,select the 
'    entire speech (except for the slug and any overline directions), 
'    apply the "LINE" paragraph syle, and then run the macro again.

Sub BreakUpLongSpeeches

    Dim Doc As Object
    Dim Enum1 As Object
    Dim TCursor As Object
    Dim SlugCursor As Object
    Dim VCursor As Object
    Dim SearchDesc As Object
    Dim CurSelection As Object
    
    Dim TextLength As Integer
    Dim SlugLineNum As Integer
    Dim SlugPageNum As Integer
    Dim CharacterName As String
    Dim MoveSuccessful As Boolean
    Dim PStyle As String
    
    Dim LengthOfSpeech As Integer
    Dim LinesLeftOnPage As Integer
    Dim Trace As String
    
    InitTables
    CollapseContdSlugs
     
    Doc = ThisComponent
    TCursor = Doc.Text.createTextCursor
    SlugCursor = Doc.Text.createTextCursor
    VCursor = Doc.getCurrentController.getViewCursor 
    
    CurSelection = Doc.currentSelection
	Doc.LockControllers   
    
    SearchDesc = Doc.createSearchDescriptor
    SearchDesc.searchString = "CHARACTERNAME"
    SearchDesc.SearchBackwards = false
    SearchDesc.SearchRegularExpression = false
    SearchDesc.SearchCaseSensitive = false
    SearchDesc.SearchSimilarity = false
    SearchDesc.SearchStyles = true
    
    LinesLeftOnPage = NLinesPerPage
    TCursor = Doc.FindFirst(SearchDesc)
    
    Do Until IsNull(TCursor)
        
        CharacterName = TCursor.GetString
        TextLength = Len(CharacterName)
        If TextLength > 1 Then            
            CharacterName = Left(CharacterName, TextLength - 1)
            TextLength = TextLength - 1
        End If
        
        While (UCase(Right(CharacterName, 9)) =  " (CONT'D)")
        	CharacterName = Left(CharacterName, TextLength - 9)
        	Textlength = TextLength - 9
        Wend
        
        VCursor.GotoRange(TCursor, false)
        SlugCursor.GotoRange(TCursor, false)
        SlugPageNum = GetPageNumFromTextCursor(TCursor)
        SlugLineNum = GetLineNumFromTextCursor(TCursor)
        
        LengthOfSpeech = GetLengthOfSpeechFromTextCursor(TCursor)
        
        If (LinesLeftOnPage = -1) Then
        	' Guess how many lines are left on the page
        	LinesLeftOnPage = NLinesPerPage - SlugLineNum - 1
        End If
        
        '
        ' IMPORTANT PARAMETERS CODED HERE
        If (LengthOfSpeech > 10 AND LengthOfSpeech > LinesLeftOnPage AND LinesLeftOnPage > 5) Then
        
        	'MsgBox("found a candidate: " + CharacterName + " lines: " + LengthOfSpeech + " starting from line# " + SlugLineNum)
        	
        	VCursor.gotoRange ( TCursor, false)
            VCursor.collapseToStart(false)
            VCursor.goDown(LinesLeftOnPage - 2, false)
            
            ' Now find the end of the line
            CurLineNum = CurrentLine(Doc)
            While (CurLineNum = CurrentLine(Doc))
                VCursor.goRight(1, false)
            Wend
            
            VCursor.goLeft(1, false)
            
            TCursor.gotoRange(VCursor, false)
            TCursor.collapseToEnd()
            Doc.Text.insertControlCharacter(TCursor, _
                com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
                
            'TCursor.goRight(1, true)    
            TCursor.ParaStyleName = "Default Paragraph Style"
            
            Doc.Text.insertControlCharacter(TCursor, _
                com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
                
            TCursor.ParaStyleName = "CHARACTERNAME"
            TCursor.SetString(CharacterName + " (CONT'D)")
            TCursor.collapseToEnd()
            
            Doc.Text.insertControlCharacter(TCursor, _
                com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
                
            TCursor.ParaStyleName = "LINE"
            TCursor.gotoEndOfParagraph(true)
            TCursor.SetString(Trim(TCursor.GetString()))
        
        End If
        
        LinesLeftOnPage = NLinesPerPage - SlugLineNum - LengthOfSpeech - 2 ' Include the slug and a blank
        
        ' Handle the case where we have more than 1 page worth of stuff here.
        If (LinesLeftOnPage <= 0) Then
        	LinesLeftOnPage = NLinesPerPage
        End If
        
        TCursor = Doc.FindNext(SlugCursor.End, SearchDesc)    
            
    Loop
    
    Doc.CurrentController.Select(CurSelection)
	Doc.UnlockControllers
    'MsgBox "Done."

End Sub
    
' =====================================================================
' EXPERIMENTAL BLOCK
'

' =====================================================================
' Remove all highlighting in the document.
Sub UnhighlightEverything

    Dim Doc As Object
    Dim Enum As Object
    Dim TextElement As Object
    
    InitTables
 
    Doc = ThisComponent
    Enum = Doc.Text.createEnumeration
 
    While Enum.hasMoreElements
        TextElement = Enum.nextElement
 
        If TextElement.supportsService("com.sun.star.text.Paragraph") Then
        	
			'TextElement.ParaBackColor, .ParaBackTransparent
            'TextElement.CharBackColor = RGB(255,255,255)
            TextElement.CharBackTransparent = True
        
        End If 
    Wend

End Sub

' =====================================================================
' HighlightBlockStageDirections
'
' Highlight only the block stage directions.
Sub HighlightBlockStageDirections

    Dim Doc As Object
    Dim SearchDesc As Object
    Dim Cursor As Object
    Dim CurSelection As Object
    Dim AppliedColor As Long ' TODO: Turn this into a parameter

    InitTables
    UnhighlightEverything

    Doc = ThisComponent
    Cursor = Doc.Text.CreateTextCursor

    CurSelection = Doc.currentSelection
    Doc.LockControllers

    SearchDesc = Doc.createSearchDescriptor
    SearchDesc.searchString = "STAGE_DIRECTION_BLOCK"
    SearchDesc.SearchBackwards = false
    SearchDesc.SearchRegularExpression = false
    SearchDesc.SearchCaseSensitive = false
    SearchDesc.SearchSimilarity = false
    SearchDesc.SearchStyles = true

    AppliedColor = RGB(255,255,0) ' Yellow

    Cursor = Doc.FindFirst(SearchDesc)

    Do Until IsNull(Cursor)

        Cursor.CharBackColor = AppliedColor

        Cursor = Doc.FindNext(Cursor.End, SearchDesc)

    Loop

    Doc.CurrentController.Select(CurSelection)
	Doc.UnlockControllers

End Sub

' =====================================================================
' HighlightCharacterLines
'
' Given a character name, highlight all that character's lines.
'
Sub HighlightCharacterLines ( CharName as String )

    Dim Doc As Object
    Dim Enum1 As Object
    Dim TextElement As Object

    Dim InCharacterLine As Boolean
    Dim HighlightColor As Long
    Dim Slug As String
    
    InitTables

    HighlightColor = RGB(255,255,0) ' Yellow
    InCharacterLine = False
     
    Doc = ThisComponent
    Enum1 = Doc.Text.createEnumeration
     
    ' loop over all paragraphs
    While Enum1.hasMoreElements

      TextElement = Enum1.nextElement
     
      If TextElement.supportsService("com.sun.star.text.Paragraph") Then
        
        If (TextElement.ParaStyleName = "CHARACTERNAME") Then
        
            ' Get the character name from TextElement.String
            Slug = TextElement.String

            ' Does it have a trailing "(CONT'D)" ?
            If (Right(Slug,8) = "(CONT'D)") Then
                Slug = Left(Slug, Len(Slug) - 8)
            End If

            Slug = Trim(UCase(Slug))

            ' Check it against the expected character name
            ' If it matches, highlight. Otherwise, remove highlighting
            If (UCase(CharName) = Slug) Then
                TextElement.CharBackColor = HighlightColor
                InCharacterLine = True
            Else
                TextElement.CharBackTransparent = True
                InCharacterLine = False
            End If

        ElseIf (TextElement.ParaStyleName = "LINE") OR _
        	(TextElement.ParaStyleName = "STAGE_DIRECTION_OVERLINE") Then

            If InCharacterLine Then
                TextElement.CharBackColor = HighlightColor
            Else
                TextElement.CharBackTransparent = True
            End If

        ElseIf (TextElement.ParaStyleName = "STAGE_DIRECTION_BLOCK") Then

            TextElement.CharBackTransparent = True
        
        End If
     
      End If

    Wend
    
End Sub

' =====================================================================
' HighlightOptions
'

Global OptionHighlightCharacter As Object

Sub HighlightOptions

    Dim Dlg As Object
    Dim OptionRemoveAllHighlights As Object
    Dim OptionHighlightStageDirections As Object
    'Dim OptionHighlightCharacter As Object
    Dim CharacterList As Object
    Dim OKButton As Object
    Dim CancelButton As Object
    
    Dim DialogResult As Integer
     ' 0 = remove all highlights, 1 = highlight directions, 2 = highlight character
    Dim OptionNumber As Integer
    Dim CharacterToHighlight As String
    Dim CharacterName As String
    
    InitTables
    
   	DialogLibraries.LoadLibrary("Standard")
    Dlg = CreateUnoDialog(DialogLibraries.Standard.HighlightOptions)
    
    Dlg.Title = "Script Highlighting Options"
    
    OptionRemoveAllHighlights = Dlg.getControl("OptionRemoveAllHighlights")
    OptionRemoveAllHighlights.State = True
    OptionHighlightStageDirections = Dlg.getControl("OptionHighlightStageDirections")
    OptionHighlightCharacter = Dlg.getControl("OptionHighlightCharacter")
    CharacterList = Dlg.getControl("CharacterList")
    
    For I=0 to NTags
    	CharacterName = Slugs(I)
    	CharacterList.AddItem(CharacterName, I)
    	If I = 0 Then
    		CharacterList.SelectItem(CharacterName, True)
    	End If
    Next I
    
    OKButton = Dlg.getControl("OKButton")
    CancelButton = Dlg.getControl("CancelButton")
    
    OptionRemoveAllHighlights.State = True     
    
    DialogResult = Dlg.Execute()
    
    If DialogResult = 1 Then
    
	    ' Determine what we wanted
	    If OptionRemoveAllHighlights.Model.State Then
	    	OptionNumber = 0
	   	Elseif OptionHighlightStageDirections.Model.State Then
	   		OptionNumber = 1
	   	Elseif OptionHighlightCharacter.Model.State Then
	   		OptionNumber = 2
	   		CharacterToHighlight = CharacterList.SelectedItem
	   	End If
	    
	    Dlg.dispose()
	    
	    ' Do the highlighting
	    If OptionNumber = 0 Then
	    	UnhighlightEverything
	    ElseIf OptionNumber = 1 Then
	    	HighlightBlockStageDirections
	    ElseIf OptionNumber = 2 Then
	    	HighlightCharacterLines (CharacterToHighlight)
	    End If
	    
	End If

End Sub

Sub OnCharacterListChanged

	OptionHighlightCharacter.State = True

End Sub






