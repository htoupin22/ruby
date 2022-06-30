<html>
<body>
<div>
<div class="navbar navbar-expand-sm navbar-light navbar-lewagon">
  <a class="navbar-brand" href="/">
    <h1 class="app-name">Cook &amp; Share</h1>
</a>
 <div class="app-name">
    <h1>Cook & Share</h1>
  </div> 

  <button class="navbar-toggler" type="button" data-toggle="collapse" data-target="#navbarSupportedContent" aria-controls="navbarSupportedContent" aria-expanded="false" aria-label="Toggle navigation">
    <span class="navbar-toggler-icon"></span>
  </button>


  <div class="collapse navbar-collapse" id="navbarSupportedContent">
    <ul class="navbar-nav mr-auto">
        <li class="nav-item">
          <a class="nav-link" href="/users/sign_in">Se connecter</a>
        </li>
    </ul>
  </div>
</div>
</body>
Sub DVP_InsererEtOuActualiserUnIndexThematique()
	Dim aLstRecettes() As String
	Dim aLstTypes() As String
	Dim aLstPrix() As String
	Dim aLstDiff() As String
					    
	ReDim aLstRecettes(0 To 0) As String
	ReDim aLstTypes(0 To 0) As String
	ReDim aLstPrix(0 To 0) As String
	ReDim aLstDiff(0 To 0) As String
					    
	'// On recupère la TdM
	'// Attention, on considère que :
	'//     1°) La TdM qui nous interresse est la 1ere
	'//     2°) La TdM est à jour
	'// La mise à jour de la table des matières qu'avec le style "Titre 1" 
	'//(pour ne pas avoir à gérer les variantes) ne fonctionne pas
	'//     (ils restent toujours pris en compte) ==&gt; traitement manuel
	ActiveDocument.Range(Start:=ActiveDocument.TablesOfContents(1).Range.Start, _
			End:=ActiveDocument.TablesOfContents(1).Range.End).Select
		aTdM = ActiveDocument.TablesOfContents(1).Range.Text
					    
		'// On stocke la TdM sous forme d'un tableau
		aNbRecettes = 0
		aTmpRecettes = ""
		While InStr(aTdM, vbCr) &lt;&gt; 0
			aNbRecettes = aNbRecettes + 1
			aTmpRecettes = aTmpRecettes + Left(aTdM, InStr(aTdM, vbTab) - 1) + "$" + Left(Mid(aTdM, _
						InStr(aTdM, vbTab) + 1), InStr(Mid(aTdM, InStr(aTdM, vbTab) + 1), vbCr) - 1) + "£"				        
			aTdM = Mid(aTdM, InStr(aTdM, vbCr) + 1)
		Wend
		ReDim aLstRecettes(0 To aNbRecettes)
		For aI = 0 To aNbRecettes - 1
			aLstRecettes(aI) = Left(aTmpRecettes, InStr(aTmpRecettes, "£") - 1)
			aTmpRecettes = Mid(aTmpRecettes, Len(aLstRecettes(aI)) + 2)
		Next
					         
		'// On parcourt la liste des recettes pour retrouver les catégories
		Selection.find.ClearFormatting
		Selection.find.Replacement.ClearFormatting
		For aI = 0 To aNbRecettes - 1
			Selection.HomeKey Unit:=wdStory
			'// On regarde si la recette est une recette principale
			Selection.find.Style = "Titre 1"
			With Selection.find
				.Text = Left(aLstRecettes(aI), InStr(aLstRecettes(aI), "$") - 1) + "^p"
				.Forward = True
				.Wrap = wdFindContinue
				.Format = True
				.MatchCase = False
				.MatchWholeWord = False
				.MatchWildcards = False
				.MatchSoundsLike = False
				.MatchAllWordForms = False
			End With
			Selection.find.Execute
			If Selection.find.Found Then
				Selection.Next(Unit:=wdTable, Count:=1).Select
			Else '// On regarde si la recette est une variante
				Selection.find.Style = "Titre 1 - Variante"
				Selection.find.Execute
				If Selection.find.Found Then
					Selection.Previous(Unit:=wdTable, Count:=1).Select
				End If
			End If
					        
			If Selection.Information(wdWithInTable) Then
			'// Attention, on considère que le champ de formulaire (type de plat : entrée, plat principal...) 
			'// qui nous interresse est le 1er
			'// ==&gt; On va récupérer tous les types de plats pour créer des catégories
				aPasTrouve = True
					            
				aJ = LBound(aLstTypes)
				While (aJ &lt; UBound(aLstTypes)) And (aPasTrouve)
					 If (Left(aLstTypes(aJ), InStr(aLstTypes(aJ), "$") - 1) = Selection.FormFields(1).result) Then
					      aPasTrouve = False
					 Else
					      aJ = aJ + 1
					 End If
				Wend
					If aPasTrouve Then
					     aLstTypes(aJ) = Selection.FormFields(1).result + "$"
					     ReDim Preserve aLstTypes(0 To (aJ + 1))
					End If
					aLstTypes(aJ) = aLstTypes(aJ) + aLstRecettes(aI) + "£"
			End If
		Next
					        
					        
		'// On supprime le contenu de la table précédente si elle existe
		Selection.GoTo What:=wdGoToBookmark, Name:="TitreTableDIndexParCategorie"
		Selection.Move Unit:=wdCharacter, Count:=2
		Selection.MoveEnd Unit:=wdSection, Count:=1
		Selection.MoveLeft Unit:=wdCharacter, Count:=3, Extend:=wdExtend
					    
					    
		For aI = LBound(aLstTypes) To UBound(aLstTypes) - 1
			With Selection
				.TypeText Left(aLstTypes(aI), InStr(aLstTypes(aI), "$") - 1)
				.TypeParagraph
				.MoveUp Unit:=wdParagraph, Count:=1, Extend:=wdExtend
			End With
			Selection.Style = ActiveDocument.Styles("Catégorie de plats")
			Selection.MoveRight Unit:=wdCharacter, Count:=1
					        
			With Selection
				.TypeText Mid(aLstTypes(aI), InStr(aLstTypes(aI), "$") + 1)
				.TypeParagraph
				.MoveUp Unit:=wdParagraph, Count:=1, Extend:=wdExtend
			End With
			Selection.find.ClearFormatting
			Selection.find.Replacement.ClearFormatting
			With Selection.find
				.Text = "$"
				.Replacement.Text = "^t"
				.Forward = True
				.Wrap = wdFindStop
				.Format = False
				.MatchCase = False
				.MatchWholeWord = False
				.MatchWildcards = False
				.MatchSoundsLike = False
				.MatchAllWordForms = False
			End With
			Selection.find.Execute Replace:=wdReplaceAll
			With Selection.find
				.Text = "£"
				.Replacement.Text = "^p"
				.Forward = True
				.Wrap = wdFindStop
				.Format = False
				.MatchCase = False
				.MatchWholeWord = False
				.MatchWildcards = False
				.MatchSoundsLike = False
				.MatchAllWordForms = False
			End With
			Selection.find.Execute Replace:=wdReplaceAll
	
			Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(18), _
					Alignment:=wdAlignTabRight, Leader:=wdTabLeaderDots
	
			Selection.MoveRight Unit:=wdCharacter, Count:=1
		Next
End Sub

