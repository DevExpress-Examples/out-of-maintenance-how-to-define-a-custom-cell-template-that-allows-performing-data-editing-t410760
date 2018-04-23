# How to define a custom cell template that allows performing data editing 


<p>The DXPivotGrid control does not support the data editing functionality out of the box. This example demonstrate how to define a custom <a href="https://documentation.devexpress.com/WPF/DevExpressXpfPivotGridPivotGridField_CellTemplatetopic.aspx">CellTemplate</a>, which provides the in-place editing functionality. Note that only the base set of features is implemented in this example: e.g. it is impossible to go to the next cell by pressing the tab or arrow keys. The PivotGridControl is optimized to provide the best performance possible during displaying and navigation through a huge number of cells, thus it does not provide a way to implement some of these features at all.</p>

<br/>


