using System;
using ExcelDna.Integration;
using ExcelDna.Integration.Helpers.DialogBox;
using WinForms = System.Windows.Forms;

[ExcelCommand(MenuName = "Workbench")]
public static class Workbench
{

  const string TITLE = "Workbench";

  static readonly Action<string> ShowError = msg => WinForms.MessageBox.Show(msg, TITLE, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
  static readonly Action<string> ShowInfo = msg => WinForms.MessageBox.Show(msg, TITLE, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);

  [ExcelCommand(MenuText = "Raw Test")]
  public static void RawTest() {
    try {

      var items = new object[] { 1, 2, 3 };
      var ans = XlCall.Excel(XlCall.xlfSetName, "items", items);

      var d = new object[2, 7];
      d[0, 5] = "List Crash";
      d[1, 0] = 15;
      d[1, 5] = "SEQUENCE(5)";

      do {
        ans = XlCall.Excel(XlCall.xlfDialogBox, d);
      } while ((bool)ans);

    }
    catch (XlCallException cex) {
      ShowError($"xlReturn: {cex.xlReturn}\n{cex.StackTrace}");
    }
    catch (Exception ex) {
      ShowError(ex.ToString());
    }
  }

  [ExcelCommand(MenuText = "Dialog Test")]
  public static void DialogTest() {
    try {

      var text = @"PRESS ESC TO EXIT

xlfDialogBox always fill result[0, 6] with
Max(1,selected trigger) so pressing
the enter key is distinguishable only
if there is a non-trigger control in the
first position.";

      // var dialog1 = new Dialog { Title = "Enter Key", }
      //   .Add(new StaticText { Left = 10, Top = 10, Width = 300, Height = 94, Text = text })
      //   .Add(new OptionGroup("og") { Text = "Option Group", IsTrigger = true })
      //   .Add(new OptionButton("ob1") { Text = "Option 1" })
      //   .Add(new OptionButton("ob2") { Text = "Option 2" })
      //   .Add(new OptionButton("ob3") { Text = "Option 3" })
      // ;
      // 
      // var dialog2 = new Dialog { Title = "Enter Key", }
      //   .Add(new OptionGroup("og") { Text = "Option Group", IsTrigger = true })
      //   .Add(new OptionButton("ob1") { Text = "Option 1" })
      //   .Add(new OptionButton("ob2") { Text = "Option 2" })
      //   .Add(new OptionButton("ob3") { Text = "Option 3" })
      //   .Add(new StaticText { Left = 10, Top = 80, Width = 300, Height = 94, Text = text })
      // ;

      var dialog1 = new Dialog { Title = "Enter Key", }
        .Add(new StaticText { Left = 10, Top = 10, Width = 300, Height = 88, Text = text })
        .Add(new TextEditBox("teb"))
        .Add(new OkButton("ok") { Top = 132, Text = "OK" })
      ;

      var dialog2 = new Dialog { Title = "Enter Key", }
        .Add(new OkButton("ok") { Text = "OK" })
        .Add(new TextEditBox("teb"))
        .Add(new StaticText { Left = 10, Top = 64, Width = 300, Height = 94, Text = text })
      ;

      var ogFirst = false;
      var loop = true;

      bool Handled(Dialog d) {
        XlCall.Excel(XlCall.xlcAlert, d.TriggerId ?? "Enter Key", 1);
        ogFirst = !ogFirst;
        return true;
      }

      while (loop) {
        if (ogFirst)
          loop = dialog2.Show(Handled);
        else
          loop = dialog1.Show(Handled);
      }

    }
    catch (Exception ex) {
      ShowError(ex.ToString());
    }
  }

  [ExcelCommand(MenuText = "Empty Id")]
  public static void EmptyId() {

    var dialog = new Dialog { Title = "Empty Id" }
      .Add(new OkButton("  "))
    ;

    var c = dialog.GetControl("");

  }

}
