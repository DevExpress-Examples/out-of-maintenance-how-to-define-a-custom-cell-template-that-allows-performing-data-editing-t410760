Imports System
Imports System.Collections.ObjectModel
Imports System.Windows
Imports DevExpress.Xpf.Core
Imports DevExpress.Xpf.PivotGrid

Namespace HowToEditCell
	''' <summary>
	''' Interaction logic for MainWindow.xaml
	''' </summary>
	Partial Public Class MainWindow
		Inherits ThemedWindow

		Public Property OrderSourceList() As ObservableCollection(Of MyOrderRow)

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub Window_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
			OrderSourceList = DatabaseHelper.CreateData()
			pivotGridControl1.DataSource = OrderSourceList
			pivotGridControl1.BestFit()
		End Sub

		Private Sub pivotGridControl1_OnCellEdit(ByVal sender As DependencyObject, ByVal args As PivotCellEditEventArgs)
			Dim pivotGrid As PivotGridControl = CType(sender, PivotGridControl)
			Dim fieldExtendedPrice As PivotGridField = pivotGrid.Fields("ExtendedPrice")
			Dim ds As PivotDrillDownDataSource = args.CreateDrillDownDataSource()
			Dim difference As Decimal = args.NewValue - args.OldValue
			Dim factor As Decimal = If(difference = args.NewValue, (difference / ds.RowCount), (difference / args.OldValue))
			For i As Integer = 0 To ds.RowCount - 1
				Dim value As Decimal = Convert.ToDecimal(ds(i)(fieldExtendedPrice))
				Dim newValue As Decimal = If(value = 0D, factor, value * (1D + factor))
				ds.SetValue(i, fieldExtendedPrice, newValue)
			Next i
		End Sub
	End Class
End Namespace
