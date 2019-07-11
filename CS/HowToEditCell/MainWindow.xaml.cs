using System;
using System.Collections.ObjectModel;
using System.Windows;
using DevExpress.Xpf.Core;
using DevExpress.Xpf.PivotGrid;

namespace HowToEditCell {
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : ThemedWindow {

        public ObservableCollection<MyOrderRow> OrderSourceList { get; set; }

        public MainWindow() {
			InitializeComponent();
		}

		void Window_Loaded(object sender, RoutedEventArgs e) {
            OrderSourceList = DatabaseHelper.CreateData();
            pivotGridControl1.DataSource = OrderSourceList;
            pivotGridControl1.BestFit();
        }

		void pivotGridControl1_OnCellEdit(DependencyObject sender, PivotCellEditEventArgs args) {
			PivotGridControl pivotGrid = (PivotGridControl)sender;
			PivotGridField fieldExtendedPrice = pivotGrid.Fields["ExtendedPrice"];
			PivotDrillDownDataSource ds = args.CreateDrillDownDataSource();
			decimal difference = args.NewValue - args.OldValue;
			decimal factor = (difference == args.NewValue) ? (difference / ds.RowCount) : (difference / args.OldValue);
			for(int i = 0; i < ds.RowCount; i++) {
				decimal value = Convert.ToDecimal(ds[i][fieldExtendedPrice]);
                decimal newValue = (value == 0m) ? factor : value * (1m + factor);
                ds.SetValue(i, fieldExtendedPrice, newValue);
			}
		}
	}
}
