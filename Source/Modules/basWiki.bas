Attribute VB_Name = "basWiki"
Option Explicit

Private Const WikiBase As String = "https://ddowiki.com/page/"
Private Const ImageBase As String = "https://ddowiki.com/images/"

Public Function MakeWiki(pstrName As String) As String
    Dim strReturn As String
    
    'Returns a WikiLink
    strReturn = pstrName
    If InStr(strReturn, " ") Then strReturn = Replace$(strReturn, " ", "_")
    MakeWiki = WikiBase & strReturn
End Function

Public Function MakeWikiItem(pstrName As String) As String
    Dim strReturn As String
    
    strReturn = pstrName
    If InStr(strReturn, " ") Then strReturn = Replace$(strReturn, " ", "_")
    MakeWikiItem = WikiBase & "Item:" & strReturn
End Function

Public Function UnWiki(pstrName As String) As String
    Dim strReturn As String
    Dim lngPos As Long
    
    strReturn = pstrName
    lngPos = InStr(strReturn, "/page/")
    If lngPos > 0 Then strReturn = Mid$(strReturn, lngPos + 6)
    If InStr(strReturn, "%27") Then strReturn = Replace$(strReturn, "%27", "'")
    If InStr(strReturn, "_") Then strReturn = Replace$(strReturn, "_", " ")
    UnWiki = strReturn
End Function

Public Function TierLink(plngTier As Long) As String
    Select Case plngTier
        Case 0: TierLink = "#Core_abilities"
        Case 1: TierLink = "#Tier_One"
        Case 2: TierLink = "#Tier_Two"
        Case 3: TierLink = "#Tier_Three"
        Case 4: TierLink = "#Tier_Four"
        Case 5: TierLink = "#Tier_Five"
        Case 6: TierLink = "#Tier_Six"
    End Select
End Function

Public Function WikiImage(pstrImage As String) As String
    WikiImage = ImageBase & pstrImage
End Function

Public Function WikiSearch(pstrSearch As String) As String
    WikiSearch = "https://ddowiki.com/index.php?search=" & Replace(pstrSearch, " ", "+")
End Function
