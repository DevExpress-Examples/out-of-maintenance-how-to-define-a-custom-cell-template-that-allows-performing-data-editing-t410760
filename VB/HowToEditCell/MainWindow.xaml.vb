Imports System
Imports System.Windows
Imports DevExpress.Xpf.PivotGrid
Imports HowToEditCell.NwindDataSetTableAdapters

Namespace HowToEditCell
    ''' <summary>
    ''' Interaction logic for MainWindow.xaml
    ''' </summary>
    Partial Public Class MainWindow
        Inherits Window

        Private salesPersonDataTable As New NwindDataSet.SalesPersonDataTable()
        Private salesPersonDataAdapter As New SalesPersonTableAdapter()

        Public Sub New()
            InitializeComponent()
            pivotGridControl1.DataSource = salesPersonDataTable
        End Sub

        Private Sub Window_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
            salesPersonDataAdapter.Fill(salesPersonDataTable)
        End Sub

        Private Sub pivotGridControl1_OnCellEdit(ByVal sender As DependencyObject, ByVal agrs As PivotCellEditEventArgs)
            Dim pivotGrid As PivotGridControl = CType(sender, PivotGridControl)
            Dim fieldExtendedPrice As PivotGridField = pivotGrid.Fields("Extended Price")
            Dim ds As PivotDrillDownDataSource = agrs.CreateDrillDownDataSource()
            Dim difference As Decimal = agrs.NewValue - agrs.OldValue
            Dim factor As Decimal = If(difference = agrs.NewValue, (difference / ds.RowCount), (difference / agrs.OldValue))
            For i As Integer = 0 To ds.RowCount - 1
                Dim value As Decimal = Convert.ToDecimal(ds(i)(fieldExtendedPrice))
                ds(i)(fieldExtendedPrice) = If(value = 0D, factor, value * (1D + factor))
            Next i
        End Sub
    End Class
End Namespace
