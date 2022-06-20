<!-- default badges list -->
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T410760)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
# Pivot Grid for WPF - How to Edit a Cell with the Cell Editing Template

The Pivot Grid control does not support the data editing functionality out of the box because it displays aggregated data. This example demonstrates how to implement a custom [CellTemplate](https://docs.devexpress.com/WPF/DevExpress.Xpf.PivotGrid.PivotGridField.CellTemplate) with the in-place editing functionality. 

This example implements the base set of features. It does not allow to go to the next cell by pressing the tab or arrow keys.

![screenshot](/images/screenshot.png)

The `PivotGridEditHelper` class implements the in-place cell editor. When editing is finished, the editor calls the `pivotGridControl1_OnCellEdit` method. This method retrieves the underlying data for the edited cell and adjusts them so that their sum equals the new value entered in the cell.

## Files to Look At

* [MainWindow.xaml](./CS/HowToEditCell/MainWindow.xaml) (VB: [MainWindow.xaml](./VB/HowToEditCell/MainWindow.xaml))
* [MainWindow.xaml.cs](./CS/HowToEditCell/MainWindow.xaml.cs) (VB: [MainWindow.xaml.vb](./VB/HowToEditCell/MainWindow.xaml.vb))
* [PivotGridEditHelper.cs](./CS/HowToEditCell/PivotGridEditHelper.cs) (VB: [PivotGridEditHelper.vb](./VB/HowToEditCell/PivotGridEditHelper.vb))

## Documentation

* [Drill Down to the Underlying Data](https://docs.devexpress.com/WPF/8056)
* [Delegate Commands](https://docs.devexpress.com/WPF/17353)
