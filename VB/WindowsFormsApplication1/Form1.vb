Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms
Imports DevExpress.XtraBars.Ribbon
Imports DevExpress.Office
Imports DevExpress.XtraRichEdit.API.Native

Namespace WindowsFormsApplication1
	Partial Public Class Form1
		Inherits RibbonForm

		Public Sub New()
			InitializeComponent()
		End Sub

		Private someDT As New DataTable()
		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
			Dim loremIpsumString As String = "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna." & vbCrLf & "Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus." & vbCrLf & "Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Proin pharetra nonummy pede. Mauris et orci."


			someDT.Columns.Add("Header", GetType(String))
			someDT.Columns.Add("Content", GetType(String))

			For i As Integer = 0 To 9
				someDT.Rows.Add(New Object() { "Header " & i.ToString(), loremIpsumString })
			Next i
		End Sub

		Private Sub barButtonItem1_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles barButtonItem1.ItemClick
			For Each dataRow As DataRow In someDT.Rows
				Dim headerParagraph As Paragraph = richEditControl1.Document.Paragraphs.Append()
				SetParagraphStyleSettings(headerParagraph, 1)
				Dim headerRange As DocumentRange = richEditControl1.Document.AppendText(dataRow("Header").ToString())
				Dim singleParagraph As Paragraph = richEditControl1.Document.Paragraphs.Append()
				SetParagraphStyleSettings(singleParagraph, 0)
				ChangeHeaderCharacterProeprties(headerRange)
				richEditControl1.Document.AppendText(dataRow("Content").ToString())
				richEditControl1.Document.AppendText(Characters.PageBreak.ToString())
			Next dataRow

			InsertTOC("\h")
		End Sub

		Private Sub InsertTOC(ByVal switches As String)
			Dim field As Field = richEditControl1.Document.Fields.Create(richEditControl1.Document.Range.Start, "TOC " & switches)
			Dim cp As CharacterProperties = richEditControl1.Document.BeginUpdateCharacters(field.Range)
			cp.Bold = False
			cp.FontSize = 12
			cp.ForeColor = Color.Blue
			richEditControl1.Document.EndUpdateCharacters(cp)
			richEditControl1.Document.InsertSection(field.Range.End)
			field.Update()
		End Sub

		Private levelIndent As Integer = 1
		Private Sub SetParagraphStyleSettings(ByVal currentParagraph As Paragraph, ByVal level As Integer)
			If level > 0 Then
				Dim styleName As String = "Paragraph Level " & levelIndent.ToString()
				Dim paragraphStyle As ParagraphStyle = richEditControl1.Document.ParagraphStyles(styleName)

				If paragraphStyle Is Nothing Then
					paragraphStyle = richEditControl1.Document.ParagraphStyles.CreateNew()
					paragraphStyle.Name = styleName
					paragraphStyle.Parent = richEditControl1.Document.ParagraphStyles("Normal")
					paragraphStyle.OutlineLevel = levelIndent
					richEditControl1.Document.ParagraphStyles.Add(paragraphStyle)
				End If
				currentParagraph.Style = paragraphStyle
			Else
				currentParagraph.Style = richEditControl1.Document.ParagraphStyles("Normal")
			End If
		End Sub

		Private Sub ChangeHeaderCharacterProeprties(ByVal headerRange As DocumentRange)
			Dim cp As CharacterProperties = richEditControl1.Document.BeginUpdateCharacters(headerRange)
			cp.Bold = True
			cp.ForeColor = Color.Blue
			richEditControl1.Document.EndUpdateCharacters(cp)
		End Sub
	End Class
End Namespace
