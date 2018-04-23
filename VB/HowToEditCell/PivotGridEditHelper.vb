Imports System
Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Input
Imports DevExpress.Mvvm
Imports DevExpress.Xpf.Editors
Imports DevExpress.Xpf.PivotGrid
Imports DevExpress.Xpf.PivotGrid.Internal
Imports System.Windows.Media


Namespace HowToEditCell

    Friend Class PivotGridEditHelper
        Inherits DependencyObject

        Public Delegate Sub OnCellEditHandler(ByVal sender As DependencyObject, ByVal args As PivotCellEditEventArgs)

        Public Shared ReadOnly OnCellEditEvent As RoutedEvent
        Public Shared ReadOnly UseHelperProperty As DependencyProperty
        Public Shared ReadOnly LostFocusCommandProperty As DependencyProperty
        Public Shared ReadOnly EditedCellProperty As DependencyProperty

        Shared Sub New()
            OnCellEditEvent = EventManager.RegisterRoutedEvent("OnCellEdit", RoutingStrategy.Direct, GetType(OnCellEditHandler), GetType(PivotGridEditHelper))
            LostFocusCommandProperty = DependencyProperty.RegisterAttached("LostFocusCommand", GetType(ICommand), GetType(PivotGridEditHelper), New PropertyMetadata(AddressOf OnLostFocusCommandPropertyChanged))
            EditedCellProperty = DependencyProperty.RegisterAttached("EditedCell", GetType(TextEdit), GetType(PivotGridEditHelper), New PropertyMetadata(Nothing, Sub(d, e) OnEditedCellPropertyChanged(d, e)))
            UseHelperProperty = DependencyProperty.RegisterAttached("UseHelper", GetType(PivotGridControl), GetType(PivotGridEditHelper), New PropertyMetadata(Nothing, Sub(d, e) OnUseHelperChanged(e)))
        End Sub

        Public Shared Sub AddOnCellEditHandler(ByVal d As DependencyObject, ByVal handler As OnCellEditHandler)
            Dim element As UIElement = TryCast(d, UIElement)
            If element IsNot Nothing Then
                element.AddHandler(OnCellEditEvent, handler)
            End If
        End Sub
        Public Shared Sub RemoveOnCellEditHandler(ByVal d As DependencyObject, ByVal handler As OnCellEditHandler)
            Dim element As UIElement = TryCast(d, UIElement)
            If element IsNot Nothing Then
                element.RemoveHandler(OnCellEditEvent, handler)
            End If
        End Sub

        Private Shared Sub OnLostFocusCommandPropertyChanged(ByVal obj As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim element As UIElement = TryCast(obj, UIElement)
            If element Is Nothing Then
                Return
            End If

            If e.NewValue Is Nothing Then
                RemoveHandler element.LostFocus, AddressOf OnEditLostFocus
                RemoveHandler element.LostKeyboardFocus, AddressOf OnEditLostFocus
            Else
                AddHandler element.LostFocus, AddressOf OnEditLostFocus
                AddHandler element.LostKeyboardFocus, AddressOf OnEditLostFocus
            End If
        End Sub
        Private Shared Sub OnEditedCellPropertyChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim edit As TextEdit = TryCast(e.OldValue, TextEdit)
            If edit IsNot Nothing Then
                edit.EditMode = EditMode.InplaceInactive
            End If
        End Sub
        Private Shared Sub OnUseHelperChanged(ByVal e As DependencyPropertyChangedEventArgs)
            Dim newPivot As PivotGridControl = TryCast(e.NewValue, PivotGridControl)
            If newPivot IsNot Nothing Then
                AddHandler newPivot.GridLayout, AddressOf OnEditInactive
                DependencyPropertyDescriptor.FromProperty(PivotGridControl.LeftTopCoordProperty, GetType(PivotGridControl)).AddValueChanged(newPivot, AddressOf OnEditInactive)
            End If
            Dim oldPivot As PivotGridControl = TryCast(e.OldValue, PivotGridControl)
            If oldPivot IsNot Nothing Then
                RemoveHandler oldPivot.GridLayout, AddressOf OnEditInactive
                DependencyPropertyDescriptor.FromProperty(PivotGridControl.LeftTopCoordProperty, GetType(PivotGridControl)).RemoveValueChanged(oldPivot, AddressOf OnEditInactive)
            End If
        End Sub
        Private Shared Sub OnEditInactive(ByVal sender As Object, ByVal e As EventArgs)
            SetEditedCell(DirectCast(sender, PivotGridControl), Nothing)
        End Sub

        Public Shared Function GetLostFocusCommand(ByVal obj As DependencyObject) As ICommand
            Return DirectCast(obj.GetValue(LostFocusCommandProperty), ICommand)
        End Function
        Public Shared Sub SetLostFocusCommand(ByVal obj As DependencyObject, ByVal value As Object)
            obj.SetValue(LostFocusCommandProperty, value)
        End Sub

        Public Shared Function GetEditedCell(ByVal obj As DependencyObject) As TextEdit
            Return DirectCast(obj.GetValue(EditedCellProperty), TextEdit)
        End Function
        Public Shared Sub SetEditedCell(ByVal obj As DependencyObject, ByVal value As TextEdit)
            obj.SetValue(EditedCellProperty, value)
        End Sub

        Public Shared Function GetUseHelper(ByVal obj As DependencyObject) As Boolean
            Return DirectCast(obj.GetValue(UseHelperProperty), Boolean)
        End Function
        Public Shared Sub SetUseHelper(ByVal obj As DependencyObject, ByVal value As Boolean)
            obj.SetValue(UseHelperProperty, value)
        End Sub

        Private Shared Sub OnEditLostFocus(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim edit As TextEdit = DirectCast(sender, TextEdit)
            If edit.IsKeyboardFocusWithin Then
                Return
            End If
            DirectCast(edit.GetValue(LostFocusCommandProperty), ICommand).Execute(sender)
        End Sub
        Private Shared Function IsParent(ByVal parent As TextEdit, ByVal child As DependencyObject) As Boolean
            If child Is Nothing OrElse parent Is Nothing Then
                Return False
            End If
            Do While child IsNot parent AndAlso child IsNot Nothing
                child = VisualTreeHelper.GetParent(child)
            Loop
            Return child Is parent
        End Function




        Public Shared ReadOnly Property Enter() As ICommand(Of Object)
            Get
                Return New DelegateCommand(Of Object)(AddressOf OnEditValue, Nothing, False)
            End Get
        End Property
        Public Shared ReadOnly Property StartEdit() As ICommand(Of Object)
            Get
                Return New DelegateCommand(Of Object)(AddressOf OnTextEditMouseDown, Nothing, False)
            End Get
        End Property

        Private Shared Sub OnEditValue(ByVal sender As Object)
            Dim pivotGrid As PivotGridControl = If(sender Is Nothing, Nothing, FindParentPivotGrid(DirectCast(sender, DependencyObject)))
            If pivotGrid IsNot Nothing Then
                If GetEditedCell(pivotGrid) Is Nothing Then
                    Return
                End If
                SetEditedCell(pivotGrid, Nothing)
            End If

            Dim edit As TextEdit = TryCast(sender, TextEdit)

            If edit Is Nothing OrElse TryCast(edit.DataContext, CellsAreaItem) Is Nothing Then
                Return
            End If
            Dim item As CellsAreaItem = TryCast(edit.DataContext, CellsAreaItem)
            Dim newValue As Decimal = Nothing
            Dim oldValue As Decimal = Nothing

            If edit.EditValue IsNot Nothing AndAlso Decimal.TryParse(edit.EditValue.ToString(), newValue) Then
                If item.Value Is Nothing OrElse (Not Decimal.TryParse(item.Value.ToString(), oldValue)) Then
                    Return
                End If

                If pivotGrid Is Nothing Then
                    Return
                End If

                Dim args As RoutedEventArgs = New PivotCellEditEventArgs(OnCellEditEvent, item, pivotGrid, newValue, oldValue)
                pivotGrid.RaiseEvent(args)
                pivotGrid.RefreshData()
            Else
            End If
        End Sub
        Private Shared Sub OnTextEditMouseDown(ByVal sender As Object)
            Dim edit As TextEdit = TryCast(sender, TextEdit)
            Dim cell As CellsAreaItem = TryCast(edit.DataContext, CellsAreaItem)
            If edit Is Nothing OrElse cell Is Nothing OrElse cell.PivotGrid Is Nothing Then
                Return
            End If
            '#Region "EditingOfLastLevelCell"
            Dim lastLevelCell As Boolean = cell.Item.RowValueType = DevExpress.XtraPivotGrid.PivotGridValueType.Value AndAlso cell.Item.IsFieldValueExpanded(cell.Item.RowField) AndAlso cell.Item.ColumnValueType = DevExpress.XtraPivotGrid.PivotGridValueType.Value AndAlso cell.Item.IsFieldValueExpanded(cell.Item.ColumnField)
            If Not lastLevelCell Then
                Return
            End If
            '#End Region
            edit.EditMode = EditMode.InplaceActive
            edit.Focus()
            Keyboard.Focus(edit)
            SetEditedCell(cell.PivotGrid, edit)
        End Sub
        Private Shared Function FindParentPivotGrid(ByVal item As DependencyObject) As PivotGridControl
            Dim parent As DependencyObject = System.Windows.Media.VisualTreeHelper.GetParent(item)
            If parent Is Nothing Then
                Return Nothing
            End If
            Dim pivot As PivotGridControl = TryCast(parent, PivotGridControl)
            If pivot IsNot Nothing Then
                Return pivot
            End If
            Return FindParentPivotGrid(parent)
        End Function
    End Class

    Public Class PivotCellEditEventArgs
        Inherits RoutedEventArgs

        Private ReadOnly cellItem As CellsAreaItem
        Private ReadOnly pivotGrid As PivotGridControl


        Private ReadOnly newValue_Renamed, oldValue_Renamed As Decimal

        Public ReadOnly Property NewValue() As Decimal
            Get
                Return newValue_Renamed
            End Get
        End Property
        Public ReadOnly Property OldValue() As Decimal
            Get
                Return oldValue_Renamed
            End Get
        End Property

        Friend Sub New(ByVal routedEvent As RoutedEvent, ByVal cellItem As CellsAreaItem, ByVal pivot As PivotGridControl, ByVal newValue As Decimal, ByVal oldValue As Decimal)
            MyBase.New(routedEvent)
            Me.cellItem = cellItem
            Me.pivotGrid = pivot
            Me.newValue_Renamed = newValue
            Me.oldValue_Renamed = oldValue
            Source = pivotGrid
        End Sub

        Public Function CreateDrillDownDataSource() As PivotDrillDownDataSource
            Return pivotGrid.CreateDrillDownDataSource(cellItem.ColumnIndex, cellItem.RowIndex)
        End Function
    End Class
End Namespace
