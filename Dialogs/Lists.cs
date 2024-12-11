using System;
using System.Linq;
using ExcelDna.Integration;
using ExcelDna.Integration.Helpers.DialogBox;

namespace Dialogs
{
  [ExcelCommand(MenuName = "Lists")]
  public static class Lists
  {

    /// <summary>
    /// An example illustrating different ways to populate a
    /// ListControl: Add / AddRange, Constructor, xlfSetName, Formulas
    /// </summary>
    [ExcelCommand(MenuText = "List items")]
    public static void ListItems() {
      try {

        // 'fillMethod' is used in the DropDown constructor
        var fillMethod = new object[] { "Enumerable", "SetName", "Formula" };

        // An array of objects, i.e. items
        var items = new object[] { "Item1", "Item2", "Item3", "Item4", "Item5" };

        var dialog = new Dialog { Width = 300, Height = 230, Title = "List Example", }
          .Add(new StaticText { Left = 15, Top = 5, Text = "Choose fill method:", })
          .Add(new DropDown(fillMethod, "dd") { Left = 15, Top = 21, Width = 270, IsTrigger = true, })
          .Add(new StaticText { Left = 15, Top = 48, Text = "Items:", })
          .Add(
            new ListBox("lb") { Left = 15, Top = 64, Width = 270, Height = 120, }
            .AddItemRange(items) // Initially fill the ListBox with the AddItemRange method
          )
          .Add(new OkButton { Left = 110, Top = 196, Width = 80, Height = 24, Text = "OK", })
        ;

        // Set the focus
        dialog.SetFocus("lb");

        // Get a reference to the ListBox control
        var lb = dialog.GetControl<ListBox>("lb");

        // The TriggerHandler delegate
        bool Handled(Dialog d) {

          if (d.TriggerId == "dd") {

            var dd = d.TriggerControl as DropDown;

            if (dd.SelectedItem.Equals("Enumerable")) {
              lb.Formula = null;
              lb.ClearItems();
              lb.AddItemRange(items);
            }
            else if (dd.SelectedItem.Equals("SetName")) {
              // WARN: if Formula is set, ClearItems throws
              if (lb.Formula is null) lb.ClearItems();
              // WARN: the second argument must be an array; an Enumerable crashes Excel
              var nitems = items.Select(s => "Name" + s).ToArray();
              var name = "LbItems";
              XlCall.Excel(XlCall.xlfSetName, name, nitems);
              lb.Formula = name;
            }
            else if (dd.SelectedItem.Equals("Formula")) {
              // WARN: if Formula is set, ClearItems throws
              if (lb.Formula is null) lb.ClearItems();
              // Fill the ListBox using an Excel formula
              lb.Formula = "\"FormulaItem\"&SEQUENCE(5)";
            }

            lb.SelectedIndex = 0;
            d.SetFocus("lb");
            return false;

          }

          return true;

        }

        var ans = dialog.Show(Handled);

      }
      catch (Exception ex) {
        XlCall.Excel(XlCall.xlcAlert, ex.Message, 3);
      }
    }

  }
}
