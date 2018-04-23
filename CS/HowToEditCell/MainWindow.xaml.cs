using System;
using System.Windows;
using DevExpress.Xpf.PivotGrid;
using HowToEditCell.NwindDataSetTableAdapters;

namespace HowToEditCell {
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window {

		NwindDataSet.SalesPersonDataTable salesPersonDataTable = new NwindDataSet.SalesPersonDataTable();
		SalesPersonTableAdapter salesPersonDataAdapter = new SalesPersonTableAdapter();

		public MainWindow() {
			InitializeComponent();
			pivotGridControl1.DataSource = salesPersonDataTable;
		}

		void Window_Loaded(object sender, RoutedEventArgs e) {
			salesPersonDataAdapter.Fill(salesPersonDataTable);
		}

		void pivotGridControl1_OnCellEdit(DependencyObject sender, PivotCellEditEventArgs agrs) {
			PivotGridControl pivotGrid = (PivotGridControl)sender;
			PivotGridField fieldExtendedPrice = pivotGrid.Fields["Extended Price"];
			PivotDrillDownDataSource ds = agrs.CreateDrillDownDataSource();
			decimal difference = agrs.NewValue - agrs.OldValue;
			decimal factor = (difference == agrs.NewValue) ? (difference / ds.RowCount) : (difference / agrs.OldValue);
			for(int i = 0; i < ds.RowCount; i++) {
				decimal value = Convert.ToDecimal(ds[i][fieldExtendedPrice]);
				ds[i][fieldExtendedPrice] = (value == 0m) ? factor : value * (1m + factor);
			}
		}
	}
}
