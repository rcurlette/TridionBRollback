<%
' BRollback v 1.0, Author Robert Curlette, www.curlette.com
' Rollback an item and all the children in the Blueprint to the previous version

Dim output
Dim tdse : set tdse = server.createObject("tds.tdse")
tdse.initialize()

if(Request.Form("sourceItemUri") <> "") then
	Dim sourceItemUri : sourceItemUri = Request.Form("sourceItemUri")
	output = "Begin rolling back " & sourceItemUri & "<br/>" & vbcrlf
	if(Request.Form("chkShowModifiedDate") = "Yes") then
		Call ShowModifiedDate(sourceItemUri)
	else 
		Call BRollback(sourceItemUri)
	end if
	
	output = output & ("<br/><br/><b>Item Rolled back</b>")
end if

set tdse = nothing	
Sub BRollback(sourceItemUri)
	Dim localizedXml : localizedXml  = ""
	Dim nodeItem
	
	Dim item : set item = tdse.GetObject(sourceItemUri, 1)
	Dim itemType : itemType = GetItemType(sourceItemUri)	
	Dim originalXml : originalXml = item.GetXml(1919)
	
	Dim localizedItemNodes : set localizedItemNodes = GetLocalizedItemNodes(sourceItemUri)
	' Process each localized item
	for each nodeItem in localizedItemNodes
		localizedXml = GetLocalizedXml(nodeItem.getAttribute("ID"))
		Call DisplayLastModifiedDate(nodeItem.getAttribute("ID"), pubUri)
		Call Rollback(nodeItem.getAttribute("ID"), pubUri)
	next
	
	Call Rollback(item.ID, pubUri)
	set item = nothing
	set nodeItem = nothing
	set localizedItemNodes = nothing
End Sub

Sub ShowModifiedDate(sourceItemUri)
	Dim localizedXml : localizedXml  = ""
	Dim nodeItem
	
	Dim item : set item = tdse.GetObject(sourceItemUri, 1)
	Dim itemType : itemType = GetItemType(sourceItemUri)	
	Dim originalXml : originalXml = item.GetXml(1919)
	
	Call DisplayLastModifiedDate(item.Id, pubUri)
	
	Dim localizedItemNodes : set localizedItemNodes = GetLocalizedItemNodes(sourceItemUri)
	' Process each localized item
	for each nodeItem in localizedItemNodes
		localizedXml = GetLocalizedXml(nodeItem.getAttribute("ID"))
		Call DisplayLastModifiedDate(nodeItem.getAttribute("ID"), pubUri)
	next
	
	set item = nothing
	set nodeItem = nothing
	set localizedItemNodes = nothing
End Sub

Sub DisplayLastModifiedDate(uri, pubUri)
	dim tridionItem : set tridionItem = tdse.getObject(uri, 1, pubUri)
	output = output & "<div class='modified-date'>" & tridionItem.ID & ", " & tridionItem.Info.RevisionDate & "</div>" & vbcrlf
	set tridionItem = nothing
End Sub

Sub Rollback(uri, pubUri)
	dim tridionItem : set tridionItem = tdse.getObject(uri, 1, pubUri)
	dim newVersion : newVersion = tridionItem.Info.Version - 1
	if(newVersion > 0) then
		Call tridionItem.Rollback(newVersion, False)
		output = output & "<div class='rollback'>" & tridionItem.ID & " rolled back to version #" & newVersion & "</div>" & vbcrlf
	end if
	set tridionItem = nothing
End Sub

Function IsCheckoutable(itemType)
	if((itemType = 16) or (itemType = 64)) then
		IsCheckoutable = true
	else
		IsCheckoutable = false
	end if
End Function

Function GetLocalizedXml(localizeditemUri)
	Dim localizedItem
	' get localized item xml
	set localizedItem = tdse.getObject(localizeditemUri,1)
	GetLocalizedXml = localizedItem.GetXml(1919)
	set localizedItem = nothing
End Function

Function GetLocalizedItemNodes(itemUri)
	Dim tridionItem : set tridionItem = tdse.GetObject(itemUri,1) 
	Dim rowFilter : set rowFilter = tdse.CreateListRowFilter()
	call rowFilter.SetCondition("ItemType", GetItemType(itemUri))
	call rowFilter.SetCondition("InclLocalCopies", true)
	Dim usingItemsXml : usingItemsXml = tridionItem.Info.GetListUsingItems(1919, rowFilter)
	
	Dim domDoc : set domDoc = GetNewDOMDocument()  
	domDoc.LoadXml(usingItemsXml)
	Dim nodeList : set nodeList = domDoc.SelectNodes("/tcm:ListUsingItems/tcm:Item[@CommentToken='LocalCopy']")
	
	set tridionItem = nothing
	set domDoc = nothing
	set GetLocalizedItemNodes = nodeList
End Function

Function GetPubUriFromitemUri(uri)
	Dim parts : parts = split(uri, "-")
	GetPubUriFromitemUri = "tcm:0-" & Replace(parts(0), "tcm:", "") & "-1"
End Function

'GetNewDOMDocument
' borrowed from Tridion PowerTools Utils.asp
Function GetNewDomDocument ()
   Dim domDoc
   On Error Resume Next
   Set domDoc = Server.CreateObject("MSXML2.DomDocument.4.0")
   If Err.number <> 0 Then
		' MSXML4.0 is not installed
		Response.Write "Please install MSXML 4.0<br/>"
		Set GetTridionDomDocument = Nothing
		Response.End
		Exit Function
   End If
   domDoc.async = False
   domDoc.setProperty "SelectionLanguage", "XPath"
   domDoc.setProperty "SelectionNamespaces", "xmlns:tcmapi='http://www.tridion.com/ContentManager/5.0/TCMAPI' xmlns:tcm='http://www.tridion.com/ContentManager/5.0' xmlns:xlink='http://www.w3.org/1999/xlink' xmlns:xsl='http://www.w3.org/1999/XSL/Transform'"
   Set GetNewDomDocument = domDoc
End Function

Function GetItemType(uri)
	Dim parts : parts = Split(uri, "-")
	if(UBound(parts) < 2) then
		GetItemType = 16
	else
		GetItemType = parts(2)
	end if
End Function
%>

<html>
<head>
    <meta charset="utf-8">
    <title>BRollback, Rollack an item 1 version and all its' Blueprint-related children</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="BRollback, Rollack an item 1 version and all its' Blueprint-related children">
    <meta name="author" content="Robert Curlette">
	<script src="bootstrap1/jquery-1.6.2.min.js"></script>
	<meta http-equiv="Content-Type" content="text/html;charset=UTF-8" />
    <link rel="stylesheet/less" type="text/css" href="bootstrap1/bootstrap-1.0.0.min.css">
	<script src="bootstrap1/less-1.1.3.min.js" type="text/javascript"></script>
	<script src="bootstrap1/jquery.tablesorter.min.js"></script>
  </head>
<body>
	<div class="result" id="result" style="dispay:none;"></div>
	<div id="errorLog"></div>
	<div id="errContent"></div>
	<div class="container">
		<section id="forms">
			<div class="span12 columns">
				<form class="form-stacked" id="frm" method="post">
					<fieldset>
						<!--<h1>View Localized Items</h1>-->
						<h2>BRollback, Rollack an item 1 version and all its' Blueprint-related children</h2>
						<div class="clearfix">
							<label>URI of Item to Rollback (Creates a new version, not deleting previous version)</label>
							<div class="input">
							  <input class="medium" id="sourceItemUri" name="sourceItemUri" size="30" type="text" value="<%=sourceItemUri%>" />
							</div>
						</div>
						</div>
						<input type="checkbox" name="chkShowModifiedDate" id="chkShowModifiedDate" value="Yes"/>Show modified date<br/>
						<br/>
						<input type="submit" class="btn primary" id="btnRollback" value="Rollback" />
					</fieldset>
				</form>
				<span><%=output%></span>
			</div>
		</section>
	</div>
</body>
</html>