using ExcelDna.Integration;
using ExcelDna.Integration.Helpers.DialogBox;

namespace Dialogs
{
  [ExcelCommand(MenuName = "ExcelSDK")]
  public static class ExcelSDK
  {

    /// <summary>
    /// The dialog box from GENERIC.c (Excel SDK)
    /// </summary>
    [ExcelCommand(MenuText = "GENERIC")]
    public static void GENERIC() {

      var dialog = new Dialog { Width = 494, Height = 210, Title = "Generic Sample Dialog" }
        .Add(new OkButtonDef { Left = 330, Top = 174, Width = 88 })
        .Add(new CancelButton { Left = 225, Top = 174, Width = 88 })
        .Add(new StaticText { Left = 19, Top = 11, Text = "&Name" })
        .Add(new TextEditBox { Left = 19, Top = 29, Width = 251 })
        .Add(new GroupBox { Left = 305, Top = 15, Width = 154, Height = 73, Text = "&College" })
        .Add(new OptionGroup("og"))
        .Add(new OptionButton { Text = "&Harvard" })
        .Add(new OptionButton { Text = "&Other" })
        .Add(new StaticText { Left = 19, Top = 50, Text = "&Reference:" })
        .Add(new RefEditBox { Left = 19, Top = 67, Width = 253 })
        .Add(new GroupBox { Left = 209, Top = 93, Width = 250, Height = 63, Text = "&Qualifications" })
        .Add(new CheckBox { Text = "&BA / BS", Value = true })
        .Add(new CheckBox { Text = "&MA / MS", Value = true })
        .Add(new CheckBox { Text = "&PhD / Other Grad" })
        .Add(new ListBox("lb") { Left = 19, Top = 99, Width = 160, Height = 96 })
      ;

      var og = dialog.GetControl<OptionGroup>("og");
      og.SelectedIndex = 0;

      var listItems = new object[] { "Bake", "Broil", "Sizzle", "Fry", "Saute" };
      var lb = dialog.GetControl<ListBox>("lb");
      lb.AddItemRange(listItems);

      //XlCall.Excel(XlCall.xlfSetName, "GENERIC_List1", listItems);
      //lb.Formula = "GENERIC_List1";

      var ans = dialog.Show();

    }

  }
}
