using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Input;
using DevExpress.Mvvm;
using DevExpress.Xpf.Editors;
using DevExpress.Xpf.PivotGrid;
using DevExpress.Xpf.PivotGrid.Internal;
using System.Windows.Media;


namespace HowToEditCell {

	class PivotGridEditHelper : DependencyObject {
		public delegate void OnCellEditHandler(DependencyObject sender, PivotCellEditEventArgs args);

		public static readonly RoutedEvent OnCellEditEvent;
		public static readonly DependencyProperty UseHelperProperty;
		public static readonly DependencyProperty LostFocusCommandProperty;
		public static readonly DependencyProperty EditedCellProperty;

		static PivotGridEditHelper() {
			OnCellEditEvent = EventManager.RegisterRoutedEvent("OnCellEdit", RoutingStrategy.Direct, typeof(OnCellEditHandler), typeof(PivotGridEditHelper));
			LostFocusCommandProperty = DependencyProperty.RegisterAttached("LostFocusCommand", typeof(ICommand), typeof(PivotGridEditHelper), new PropertyMetadata(OnLostFocusCommandPropertyChanged));
			EditedCellProperty = DependencyProperty.RegisterAttached("EditedCell", typeof(TextEdit), typeof(PivotGridEditHelper), new PropertyMetadata(null, (d, e) => OnEditedCellPropertyChanged(d, e)));
			UseHelperProperty = DependencyProperty.RegisterAttached("UseHelper", typeof(PivotGridControl), typeof(PivotGridEditHelper), new PropertyMetadata(null, (d, e) => OnUseHelperChanged(e)));
		}

		public static void AddOnCellEditHandler(DependencyObject d, OnCellEditHandler handler) {
			UIElement element = d as UIElement;
			if(element != null) {
				element.AddHandler(OnCellEditEvent, handler);
			}
		}
		public static void RemoveOnCellEditHandler(DependencyObject d, OnCellEditHandler handler) {
			UIElement element = d as UIElement;
			if(element != null) {
				element.RemoveHandler(OnCellEditEvent, handler);
			}
		}

		static void OnLostFocusCommandPropertyChanged(DependencyObject obj, DependencyPropertyChangedEventArgs e) {
			UIElement element = obj as UIElement;
			if(element == null)
				return;

			if(e.NewValue == null) {
				element.LostFocus -= OnEditLostFocus;
				element.LostKeyboardFocus -= OnEditLostFocus;
			} else {
				element.LostFocus += OnEditLostFocus;
				element.LostKeyboardFocus += OnEditLostFocus;
			}
		}
		static void OnEditedCellPropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e) {
			TextEdit edit = e.OldValue as TextEdit;
			if(edit != null) {
				edit.EditMode = EditMode.InplaceInactive;
			}
		}
		static void OnUseHelperChanged(DependencyPropertyChangedEventArgs e) {
			PivotGridControl newPivot = e.NewValue as PivotGridControl;
			if(newPivot != null) {
				newPivot.GridLayout += OnEditInactive;
				DependencyPropertyDescriptor.FromProperty(PivotGridControl.LeftTopCoordProperty, typeof(PivotGridControl)).AddValueChanged(newPivot, OnEditInactive);
			}
			PivotGridControl oldPivot = e.OldValue as PivotGridControl;
			if(oldPivot != null) {
				oldPivot.GridLayout -= OnEditInactive;
				DependencyPropertyDescriptor.FromProperty(PivotGridControl.LeftTopCoordProperty, typeof(PivotGridControl)).RemoveValueChanged(oldPivot, OnEditInactive);
			}
		}
		static void OnEditInactive(object sender, EventArgs e) {
			SetEditedCell((PivotGridControl)sender, null);
		}

		public static ICommand GetLostFocusCommand(DependencyObject obj) {
			return (ICommand)obj.GetValue(LostFocusCommandProperty);
		}
		public static void SetLostFocusCommand(DependencyObject obj, object value) {
			obj.SetValue(LostFocusCommandProperty, value);
		}

		public static TextEdit GetEditedCell(DependencyObject obj) {
			return (TextEdit)obj.GetValue(EditedCellProperty);
		}
		public static void SetEditedCell(DependencyObject obj, TextEdit value) {
			obj.SetValue(EditedCellProperty, value);
		}

		public static bool GetUseHelper(DependencyObject obj) {
			return (bool)obj.GetValue(UseHelperProperty);
		}
		public static void SetUseHelper(DependencyObject obj, bool value) {
			obj.SetValue(UseHelperProperty, value);
		}

		static void OnEditLostFocus(object sender, RoutedEventArgs e) {
			TextEdit edit = (TextEdit)sender;
			if(edit.IsKeyboardFocusWithin)
				return;
			((ICommand)edit.GetValue(LostFocusCommandProperty)).Execute(sender);
		}
		static bool IsParent(TextEdit parent, DependencyObject child) {
			if(child == null || parent == null)
				return false;
			while(child != parent && child != null)
				child = VisualTreeHelper.GetParent(child);
			return child == parent;
		}

		public static ICommand Enter {
			get {
				return new DelegateCommand<object>(OnEditValue, null);
			}
		}
		public static ICommand StartEdit {
			get {
				return new DelegateCommand<object>(OnTextEditMouseDown, null);
			}
		}

		static void OnEditValue(object sender) {
			PivotGridControl pivotGrid = sender == null ? null : FindParentPivotGrid((DependencyObject)sender);
			if(pivotGrid != null) {
				if(GetEditedCell(pivotGrid) == null)
					return;
				SetEditedCell(pivotGrid, null);
			}

			TextEdit edit = sender as TextEdit;

			if(edit == null || edit.DataContext as CellsAreaItem == null)
				return;
			CellsAreaItem item = edit.DataContext as CellsAreaItem;
			decimal newValue;
			decimal oldValue;

			if(edit.EditValue != null && decimal.TryParse(edit.EditValue.ToString(), out newValue)) {
				if(item.Value == null || !decimal.TryParse(item.Value.ToString(), out oldValue))
					return;

				if(pivotGrid == null)
					return;

				RoutedEventArgs args = new PivotCellEditEventArgs(OnCellEditEvent, item, pivotGrid, newValue, oldValue);
				pivotGrid.RaiseEvent(args);
				pivotGrid.RefreshData();
			} else {
			}
		}
		static void OnTextEditMouseDown(object sender) {
			TextEdit edit = sender as TextEdit;
            CellsAreaItem cell = edit.DataContext as CellsAreaItem;
            if (edit == null || cell == null || cell.PivotGrid == null)
				return;
            #region EditingOfLastLevelCell             
            bool lastLevelCell = cell.Item.RowValueType == DevExpress.XtraPivotGrid.PivotGridValueType.Value && cell.Item.IsFieldValueExpanded(cell.Item.RowField) &&
                cell.Item.ColumnValueType == DevExpress.XtraPivotGrid.PivotGridValueType.Value && cell.Item.IsFieldValueExpanded(cell.Item.ColumnField);
            if ( !lastLevelCell)
				return;
            #endregion
            edit.EditMode = EditMode.InplaceActive;
			edit.Focus();
			Keyboard.Focus(edit);
			SetEditedCell(cell.PivotGrid, edit);
		}
		static PivotGridControl FindParentPivotGrid(DependencyObject item) {
			DependencyObject parent = System.Windows.Media.VisualTreeHelper.GetParent(item);
			if(parent == null)
				return null;
			PivotGridControl pivot = parent as PivotGridControl;
			if(pivot != null)
				return pivot;
			return FindParentPivotGrid(parent);
		}
	}

	public class PivotCellEditEventArgs : RoutedEventArgs {
		readonly CellsAreaItem cellItem;
		readonly PivotGridControl pivotGrid;
		readonly decimal newValue, oldValue;

		public decimal NewValue {
			get { return newValue; }
		}
		public decimal OldValue {
			get { return oldValue; }
		}

		internal PivotCellEditEventArgs(RoutedEvent routedEvent, CellsAreaItem cellItem, PivotGridControl pivot, decimal newValue, decimal oldValue)
			: base(routedEvent) {
			this.cellItem = cellItem;
			this.pivotGrid = pivot;
			this.newValue = newValue;
			this.oldValue = oldValue;
			Source = pivotGrid;
		}

		public PivotDrillDownDataSource CreateDrillDownDataSource() {
			return pivotGrid.CreateDrillDownDataSource(cellItem.ColumnIndex, cellItem.RowIndex);
		}
	}
}
