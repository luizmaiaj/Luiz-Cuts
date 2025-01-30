Attribute VB_Name = "FormatADO"
'@luiz

'version 1.6
'after formatting a ticket from clipboard place cursor at the beginning of formatted text (after the P) to make it easier to input the ticket priority
'added User Story to the list of possible types with special treatment due to the space in the name
'remove the regex and wildcard settings from the search after executing a regex search for ticket number
'when adding link using the find by number first check if there isn’t a link already associated with the text, display warning

'version 1.5
'corrected some functions names
'optmised some functions to reuse others
'created one function for regex search
'added more functions that are simple reuses of the others with different options

Public Const MINIMUM_LENGTH As Integer = 5
'below the organisation and project need to be set and also other mentions of ORGANISATION and PROJECT in the code
Public Const BASE_URL As String = "https://dev.azure.com/ORGANISATION/PROJECT/_workitems/edit/"

'checks the string length to determine if it could be a ticket and also checks if there's a ':' inside the string
Private Function checkTicket(ticket As String) As Boolean

    checkTicket = True

    If Not checkTicketLength(ticket) Then
    
        checkTicket = False
    
        Exit Function
    End If
    
    If InStr(ticket, ":") < 1 Then
        MsgBox "Text doesn't appear to contain ADO ticket format: """ & ticket & """"
        
        checkTicket = False
    End If

End Function

'checks the string length to determine if it could be a ticket
Private Function checkTicketLength(ticket As String) As Boolean

    checkTicketLength = True

    If Len(ticket) < MINIMUM_LENGTH Then 'check if there is a reasonably long string
        MsgBox "No text or text shorter than " & CStr(MINIMUM_LENGTH) & " characters: """ & ticket & """"
        
        checkTicketLength = False
        
        Exit Function
    End If
    
End Function

'removes any kind of line break and shortens 'product backlog item' into PBI
Private Function cleanTicket(ticket As String) As String

    'remove any line break
    ticket = replace(ticket, vbCr, "")
    ticket = replace(ticket, vbLf, "")
    ticket = replace(ticket, vbCrLf, "") '// or vbNewLine
    
    ticket = replace(ticket, "Product Backlog Item", "PBI") 'replace Product Backlog Item by PBI
    cleanTicket = replace(ticket, "User Story", "User_Story") 'replace User Story by User_Story to avoid having the space in the name
    
End Function

Private Sub formatTicket(ticket As String, Optional ByVal simple As Boolean = False)

    If Not checkTicket(ticket) Then Exit Sub 'if the string is not MINIMUM_LENGTH at least then exit
    
    Dim ticketLength As Long 'store the ticket string length to determine if a line break was removed to put it back at the end of the formatting
    ticketLength = Len(ticket)
    
    ticket = cleanTicket(ticket) 'clean the string before formatting
    
    'Dim lengthChanged As Boolean
    'lengthChanged = Len(ticket) <> ticketLength
    
    Dim broken
    broken = Split(ticket) 'split string on spaces
    
    Dim workItemType As String
    Dim ticketNumber As String
    
    workItemType = broken(0) 'get work item type
    workItemType = replace(workItemType, "User_Story", "User Story") 'put back the User Story type with the space
    
    ticketNumber = replace(broken(1), ":", "", vbTextCompare) 'get ticket number and remove any :
    
    broken = Split(ticket, ":", 2, vbTextCompare)  'split string again but this time on the :
    
    Dim title As String
    title = broken(1)  'get the title
    
    Dim fullTitle As String 'to be constructed below depending on mode
    
    Dim address As String
    
    If Selection.Hyperlinks.Count > 0 Then
        address = Selection.Hyperlinks(1).address 'store the link if there is one
    Else
        address = BASE_URL + ticketNumber 'if no link found use base url and ticket number
    End If
    
    If Not simple Then 'formatting used on the weekly report
        fullTitle = "P " + workItemType + ":" + title + " [PROJECT-" + ticketNumber + "]: "  'create formatted ticket text
        Selection.Text = fullTitle  'additional step above because it will be reused later
        
        Selection.MoveEnd Unit:=wdCharacter, Count:=-(Len(" [PROJECT-" + ticketNumber + "]: ")) 'remove the ticket number from the selection so it does not get formatted to bold
        
        Selection.Font.Bold = True 'format current selection as bold (except the ticket number)
        
        Selection.Start = Selection.End 'move selection to just before the ticket number
        
        Selection.MoveEnd Unit:=wdCharacter, Count:=(Len(" [PROJECT-" + ticketNumber + "]: ") - 2) 'move the end of the selection to the end of the ticket number
        Selection.MoveStart Unit:=wdCharacter, Count:=1 'move the beginning of the selection by one character to not include the space between the title and ticket number
        
    Else 'formatting used on the release notes
        fullTitle = "[PROJECT-" + ticketNumber + "] -" + title 'create formatted ticket text
        Selection.Text = fullTitle  'additional step above because it will be reused later
        
        Selection.MoveEnd Unit:=wdCharacter, Count:=(1 - Len("] -" + title))
        
    End If
    
    If Len(address) > 0 Then 'set link
        ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, address:=address  'adds a link to the current selection that should be the ticket number
    End If
    
    If Not simple Then
        Selection.Move Unit:=wdCharacter, Count:=-(Len(fullTitle) - 3) 'move to the beginning, after the P for priority
    Else
        Selection.Move Unit:=wdCharacter, Count:=-Len("[PROJECT-" + ticketNumber + "] -") 'move to the beginning
    End If

    'If lengthChanged Then Selection.Text = Selection.Text + vbCrLf 'not working correctly yet, to fix ***

End Sub

Sub formatTicketFromText()

    formatTicket Selection.Text

End Sub

Sub simpleFormatURLFromClipboard()

    Dim cb As DataObject ' clipboard contents
    Set cb = New DataObject
    
    cb.GetFromClipboard ' get clipboard content to data object
        
    formatTicket cb.GetText, True

End Sub

Sub formatTicketFromClipboard()

    Dim cb As DataObject ' clipboard contents
    Set cb = New DataObject
    
    cb.GetFromClipboard ' get clipboard content to data object
    
    formatTicket cb.GetText

End Sub

'copies the selected text to clipboard and then highlights it in yellow
Sub copyAndHighlight()

    Dim cb As DataObject
    Set cb = New DataObject
    
    cb.SetText Selection.Text
    
    cb.PutInClipboard
    
    Selection.Range.HighlightColorIndex = wdYellow
    
End Sub

Sub makeLink()

    If Not checkTicketLength(Selection.Text) Then Exit Sub 'if the string is not MINIMUM_LENGTH at least then exit
    
    Dim ticketNumber As String: ticketNumber = getNumber(Selection.Text)
    
    If Len(ticketNumber) < MINIMUM_LENGTH Then
        MsgBox "No number found in selected text"
        
        Exit Sub 'define the lenght and display error***
    End If
    
    Selection.Text = "[PROJECT-" + ticketNumber + "]"
    
    addLinkToSelection getNumber(ticketNumber)

End Sub

Function getNumber(s As String) As String
    ' Variables needed (remember to use "option explicit").
    Dim retval As String    ' This is the return string.
    Dim i As Integer        ' Counter for character position.

    ' Initialise return string to empty
    retval = ""

    ' For every character in input string, copy digits to return string
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
            retval = retval + Mid(s, i, 1)
        End If
    Next
    
    getNumber = retval ' Then return the return string
End Function

Private Sub addLinkToSelection(ticketNumber As String)
    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, address:=BASE_URL & ticketNumber
End Sub

Private Function regExFind(regex As String) As Boolean
        
        'set the parameters and execute the search
        With Selection.Find
            .Text = regex
            .MatchWildcards = True
            .Forward = True
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchFuzzy = False
            .Execute Format:=False, Forward:=True
        End With
        
        regExFind = Selection.Find.Found

    If Not regExFind Then
        Dim result As VbMsgBoxResult: result = MsgBox("No ticket found, return to the beginning of the document?", vbYesNo)
        
        If result = vbYes Then
            Selection.HomeKey Unit:=wdStory
            'Selection.EndKey Unit:=wdStory
        End If
    End If

    'remove the wildcard settings for the search
        With Selection.Find
            .Text = ""
            .MatchWildcards = False
        End With

End Function

'search for ticket in text and add the link
Sub findTicket()

        regExFind "\[PROJECT-[0-9]{5,6}\]"
    
End Sub

Sub findTicketAndAddLink()

    If regExFind("\[PROJECT-[0-9]{5,6}\]") Then
        addLinkToSelection getNumber(Selection.Range)
    End If
    
End Sub

Sub findNumber()

    regExFind "[0-9]{5,6}"

End Sub

Sub findNumberAndAddLink()

    If regExFind("[0-9]{5,6}") Then
        Dim ticketNumber As String: ticketNumber = getNumber(Selection.Range)
    
        If Selection.Hyperlinks.Count = 0 Then

            Selection.Text = "[ORGANISATION-" + ticketNumber + "]"
            addLinkToSelection getNumber(Selection.Range)
            
        Else
            MsgBox "Found " + Selection.Text + "but it already has a link. Use Make Link to force adding a link"
            
        End If
    End If
    
End Sub











