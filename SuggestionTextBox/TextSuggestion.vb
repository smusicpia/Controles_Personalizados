Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Drawing

Public Class TextSuggestion
	Inherits TextBox

#Region "Propiedades públicas"
	<DefaultValue(10)>
	Public Property MaxNumOfSuggestions As Integer

	Public Property SuggestDataSource As IEnumerable(Of String)
#End Region

#Region "Variables privadas"
	Private _listBox As ListBox
	Private _listBoxAddedToForm As Boolean = False
#End Region

#Region "Constructor"
	Public Sub New()
		MaxNumOfSuggestions = 10
		_listBox = New ListBox()
		AddHandler KeyDown, AddressOf This_KeyDown
		AddHandler KeyUp, AddressOf This_KeyUp
		AddHandler LostFocus, AddressOf This_LostFocus
		AddHandler _listBox.Click, AddressOf ListBox_Click
	End Sub
#End Region

#Region "ListBox"
	Private Sub ShowListBox()
		If Not _listBoxAddedToForm Then
			Me.TopLevelControl.Controls.Add(_listBox)
			Dim controlLocation As Point _
				= Me.TopLevelControl.PointToClient(Me.Parent.PointToScreen(Me.Location))
			_listBox.Left = controlLocation.X
			_listBox.Top = controlLocation.Y + Me.Height
			_listBox.Font = Me.Font
			_listBox.Width = Me.Width
			_listBox.MinimumSize = New Size(Me.Width, _listBox.MinimumSize.Height)
			_listBox.Height = _listBox.ItemHeight * (MaxNumOfSuggestions + 1)
			_listBoxAddedToForm = True
		End If
		_listBox.Visible = True
		_listBox.BringToFront()
	End Sub

	Private Sub HideListBox()
		_listBox.Visible = False
	End Sub

	Private Sub UpdateListBox()
		If SuggestDataSource IsNot Nothing AndAlso Not String.IsNullOrEmpty(Me.Text) Then
			Dim result As IEnumerable(Of String) = SuggestDataSource _
				.Where(Function(s) s.StartsWith(Me.Text, StringComparison.OrdinalIgnoreCase) _
					AndAlso Not s.Equals(Me.Text, StringComparison.OrdinalIgnoreCase)) _
				.OrderBy(Function(s) s) _
				.Take(MaxNumOfSuggestions)
			If result.Count() > 0 Then
				_listBox.DataSource = result.ToList()
				ShowListBox()
			Else
				HideListBox()
			End If
		Else
			HideListBox()
		End If
	End Sub

	Private Sub ListBox_Click(sender As Object, e As EventArgs)
		If _listBox.SelectedIndex >= 0 Then
			Text = _listBox.SelectedItem.ToString()
		End If
		HideListBox()
	End Sub

#End Region

#Region "Entrada de teclado"
	Private Sub This_KeyDown(sender As Object, e As KeyEventArgs)
		Select Case e.KeyCode
			Case Keys.Down
				If _listBox.Visible AndAlso _listBox.SelectedIndex < _listBox.Items.Count - 1 Then
					_listBox.SelectedIndex += 1
				End If
				e.SuppressKeyPress = True
			Case Keys.Up
				If _listBox.Visible AndAlso _listBox.SelectedIndex >= 0 Then
					_listBox.SelectedIndex -= 1
				End If
				e.SuppressKeyPress = True
			Case Keys.Enter
				If _listBox.Visible Then
					If _listBox.SelectedIndex >= 0 Then
						Text = _listBox.SelectedItem.ToString()
						SelectAll()
					End If
					HideListBox()
					e.SuppressKeyPress = True
				End If
		End Select
	End Sub

	Dim _lastText As String
	Private Sub This_KeyUp(sender As Object, e As KeyEventArgs)
		If Me.Text <> _lastText Then
			UpdateListBox()
			_lastText = Me.Text
		End If
	End Sub
#End Region

#Region "LostFocus"
	Protected Overrides Sub OnLostFocus(e As EventArgs)
		If Not _listBox.ContainsFocus Then
			MyBase.OnLostFocus(e)
		End If
	End Sub

	Private Sub This_LostFocus(sender As Object, e As EventArgs)
		HideListBox()
	End Sub
#End Region

End Class
