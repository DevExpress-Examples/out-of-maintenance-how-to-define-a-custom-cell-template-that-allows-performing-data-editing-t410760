<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128578603/15.2.11%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T410760)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
[![](https://img.shields.io/badge/💬_Leave_Feedback-feecdd?style=flat-square)](#does-this-example-address-your-development-requirementsobjectives)
<!-- default badges end -->
<!-- default file list -->
*Files to look at*:

* [MainWindow.xaml](./CS/HowToEditCell/MainWindow.xaml) (VB: [MainWindow.xaml](./VB/HowToEditCell/MainWindow.xaml))
* [MainWindow.xaml.cs](./CS/HowToEditCell/MainWindow.xaml.cs) (VB: [MainWindow.xaml.vb](./VB/HowToEditCell/MainWindow.xaml.vb))
* [PivotGridEditHelper.cs](./CS/HowToEditCell/PivotGridEditHelper.cs) (VB: [PivotGridEditHelper.vb](./VB/HowToEditCell/PivotGridEditHelper.vb))
<!-- default file list end -->
# How to define a custom cell template that allows performing data editing 


<p>The DXPivotGrid control does not support the data editing functionality out of the box. This example demonstrate how to define a custom <a href="https://documentation.devexpress.com/WPF/DevExpressXpfPivotGridPivotGridField_CellTemplatetopic.aspx">CellTemplate</a>, which provides the in-place editing functionality. Note that only the base set of features is implemented in this example: e.g. it is impossible to go to the next cell by pressing the tab or arrow keys. The PivotGridControl is optimized to provide the best performance possible during displaying and navigation through a huge number of cells, thus it does not provide a way to implement some of these features at all.</p>

<br/>


<!-- feedback -->
## Does this example address your development requirements/objectives?

[<img src="https://www.devexpress.com/support/examples/i/yes-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=wpf-pivot-grid-define-custom-cell-template-to-performing-data-editing&~~~was_helpful=yes) [<img src="https://www.devexpress.com/support/examples/i/no-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=wpf-pivot-grid-define-custom-cell-template-to-performing-data-editing&~~~was_helpful=no)

(you will be redirected to DevExpress.com to submit your response)
<!-- feedback end -->
