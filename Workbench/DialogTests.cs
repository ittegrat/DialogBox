using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using ExcelDna.Integration.Helpers.DialogBox;
using WinForms = System.Windows.Forms;

namespace DialogTests
{

  [ExcelCommand(MenuName = "Dialog Tests")]
  public static class XlmDialogTests
  {

    const string TITLE = "Run Tests";

    static readonly Action<string> ShowError = msg => WinForms.MessageBox.Show(msg, TITLE, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
    static readonly Action<string> ShowInfo = msg => WinForms.MessageBox.Show(msg, TITLE, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);

    [ExcelCommand(MenuText = TITLE)]
    public static void RunTests() {
      try {

        string ans;

        ans = Controls.RunTests();
        if (ans != null) { ShowError($"Test failed: {nameof(Controls)}.{ans}"); return; };

        ans = DialogOps.RunTests();
        if (ans != null) { ShowError($"Test failed: {nameof(DialogOps)}.{ans}"); return; };
        
        ans = DialogAdd.RunTests();
        if (ans != null) { ShowError($"Test failed: {nameof(DialogAdd)}.{ans}"); return; };
        
        ans = DialogShow.RunTests();
        if (ans != null) { ShowError($"Test failed: {nameof(DialogShow)}.{ans}"); return; };

        if (ans == null) ShowInfo($"Tests passed.");

      }
      catch (XlCallException cex) {
        ShowError($"xlReturn: {cex.xlReturn}\n{cex.StackTrace}");
      }
      catch (Exception ex) {
        ShowError(ex.ToString());
      }
    }

  }

  internal static class Controls
  {

    public static string RunTests() {
      string test = null;
      try {

        test = nameof(Id);
        if (!Id()) return test;

        test = nameof(OptionGroup);
        if (!OptionGroup()) return test;

        test = nameof(ListControl);
        if (!ListControl()) return test;

        return null;

      }
      catch (Exception ex) {
        throw new InvalidOperationException($"Unexpected failure: {nameof(Controls)}.{test}.", ex);
      }
    }

    public static bool Id() {
      try {
        var control = new StaticText("");
        return false;
      }
      catch (ArgumentException) { }
      try {
        var control = new StaticText("   ");
        return false;
      }
      catch (ArgumentException) { }
      {
        var control = new StaticText(" ctrl ");
        if (control.Id != "ctrl") return false;
      }
      {
        var control = new StaticText();
        if (control.Id != null) return false;
      }
      return true;
    }
    public static bool OptionGroup() {
      try {
        var control = new OptionGroup();
        control.SelectedId = "";
        return false;
      }
      catch (ArgumentException) { }
      try {
        var control = new OptionGroup();
        control.SelectedId = "   ";
        return false;
      }
      catch (ArgumentException) { }
      try {
        var control = new OptionGroup();
        control.SelectedId = "ctrl";
        return false;
      }
      catch (ArgumentOutOfRangeException) { }
      try {
        var control = new OptionGroup();
        var dialog = new Dialog()
          .Add(control)
          .Add(new OptionButton("a"))
          .Add(new OptionButton("b"))
        ;
        control.SelectedId = "c";
        return false;
      }
      catch (ArgumentOutOfRangeException) { }
      {
        var control = new OptionGroup();
        var dialog = new Dialog()
          .Add(control)
          .Add(new OptionButton("a"))
          .Add(new OptionButton("b"))
        ;
        control.SelectedId = "b";
        if (control.SelectedIndex != 1)
          return false;
        control.SelectedId = null;
        if (control.SelectedIndex.HasValue)
          return false;
        if (!Object.ReferenceEquals(dialog.GetControl<OptionButton>("a").OptionGroup, control))
          return false;
      }
      return true;
    }
    public static bool ListControl() {
      try {
        var control = new ListBox();
        control.Formula = "";
        return false;
      }
      catch (ArgumentException) { }
      try {
        var control = new ListBox();
        control.Formula = "   ";
        return false;
      }
      catch (ArgumentException) { }
      try {
        var control = new ListBox();
        control.Formula = "DefinedName";
        control.SelectedItem = new { };
        return false;
      }
      catch (InvalidOperationException) { }
      try {
        var control = new ListBox();
        control.Formula = "DefinedName";
        var obj = control.SelectedItem;
        return false;
      }
      catch (InvalidOperationException) { }
      {
        var control = new ListBox();
        control.Formula = "  F(a,b,c)  ";
        if (control.Formula != "F(a,b,c)")
          return false;
      }
      {
        var control = new ListBox().AddItemRange(Enumerable.Range(0, 10).Cast<object>());
        control.SelectedIndex = 5;
        if (5 != (int)control.SelectedItem)
          return false;
        control.RemoveItems(4, 3);
        if (control.SelectedItem != null)
          return false;
        control.SelectedIndex = 4;
        if (7 != (int)control.SelectedItem)
          return false;
        control.RemoveItems(2, 2);
        if (7 != (int)control.SelectedItem)
          return false;
        control.RemoveItems(4);
        if (7 != (int)control.SelectedItem)
          return false;
        control.RemoveItems(2);
        if (control.SelectedItem != null)
          return false;
      }
      return true;
    }

  }

  internal static class DialogOps
  {

    public static string RunTests() {
      string test = null;
      try {

        test = nameof(GetControl);
        if (!GetControl()) return test;

        test = nameof(SetFocus);
        if (!SetFocus()) return test;

        return null;

      }
      catch (Exception ex) {
        throw new InvalidOperationException($"Unexpected failure: {nameof(DialogAdd)}.{test}.", ex);
      }
    }

    public static bool GetControl() {
      try {
        var dialog = new Dialog();
        var c = dialog.GetControl(null);
        return false;
      }
      catch (ArgumentException) { }
      try {
        var dialog = new Dialog();
        var c = dialog.GetControl("");
        return false;
      }
      catch (ArgumentException) { }
      try {
        var dialog = new Dialog();
        var c = dialog.GetControl("  ");
        return false;
      }
      catch (ArgumentException) { }
      try {
        var dialog = new Dialog();
        var c = dialog.GetControl("id");
        return false;
      }
      catch (KeyNotFoundException) { }
      {
        var dialog = new Dialog()
          .Add(new StaticText("ctrl"))
        ;
        if (dialog.GetControl(" ctrl ").Id != "ctrl")
          return false;
      }
      return true;
    }
    public static bool SetFocus() {
      try {
        var dialog = new Dialog();
        dialog.SetFocus(null);
        return false;
      }
      catch (ArgumentException) { }
      try {
        var dialog = new Dialog();
        dialog.SetFocus("");
        return false;
      }
      catch (ArgumentException) { }
      try {
        var dialog = new Dialog();
        dialog.SetFocus("   ");
        return false;
      }
      catch (ArgumentException) { }
      try {
        var dialog = new Dialog();
        dialog.SetFocus("id");
        return false;
      }
      catch (KeyNotFoundException) { }
      {
        var dialog = new Dialog()
          .Add(new StaticText("ctrl"))
        ;
        dialog.SetFocus(" ctrl ");
      }
      return true;
    }

  }

  internal static class DialogAdd
  {

    public static string RunTests() {
      string test = null;
      try {

        test = nameof(Control);
        if (!Control()) return test;

        test = nameof(OptionButton);
        if (!OptionButton()) return test;

        test = nameof(LinkedListBox);
        if (!LinkedListBox()) return test;

        test = nameof(LinkedDropDown);
        if (!LinkedDropDown()) return test;

        test = nameof(FileListBox);
        if (!FileListBox()) return test;

        test = nameof(DriveDirBox);
        if (!DriveDirBox()) return test;

        return null;

      }
      catch (Exception ex) {
        throw new InvalidOperationException($"Unexpected failure: {nameof(DialogAdd)}.{test}.", ex);
      }
    }

    public static bool Control() {
      try {
        var dialog = new Dialog()
          .Add(new StaticText("id"))
          .Add(new StaticText(" id "))
        ;
        return false;
      }
      catch (ArgumentException) { }
      return true;
    }
    public static bool OptionButton() {
      var dialog = new Dialog()
        .Add(new OptionGroup())
        .Add(new OptionButton())
        .Add(new OptionButton())
      ;
      try {
        var xdialog = new Dialog()
          .Add(new OptionButton())
          .Add(new OptionButton())
        ;
        return false;
      }
      catch (ArgumentException) { }
      return true;
    }
    public static bool LinkedListBox() {
      var dialog = new Dialog()
        .Add(new RefEditBox())
        .Add(new StaticText())
        .Add(new LinkedListBox())
      ;
      try {
        var xdialog = new Dialog()
          .Add(new StaticText())
          .Add(new LinkedListBox())
        ;
        return false;
      }
      catch (ArgumentException) { }
      return true;
    }
    public static bool LinkedDropDown() {
      var dialog = new Dialog()
        .Add(new RefEditBox())
        .Add(new StaticText())
        .Add(new LinkedDropDown())
      ;
      try {
        var xdialog = new Dialog()
          .Add(new StaticText())
          .Add(new LinkedDropDown())
        ;
        return false;
      }
      catch (ArgumentException) { }
      return true;
    }
    public static bool FileListBox() {
      var dialog = new Dialog()
        .Add(new TextEditBox())
        .Add(new StaticText())
        .Add(new FileListBox())
      ;
      try {
        var xdialog = new Dialog()
          .Add(new StaticText())
          .Add(new FileListBox())
        ;
        return false;
      }
      catch (ArgumentException) { }
      try {
        var xdialog = new Dialog()
          .Add(new RefEditBox())
          .Add(new StaticText())
          .Add(new FileListBox())
        ;
        return false;
      }
      catch (ArgumentException) { }
      return true;
    }
    public static bool DriveDirBox() {
      var dialog = new Dialog()
        .Add(new TextEditBox())
        .Add(new StaticText())
        .Add(new FileListBox())
        .Add(new DriveDirBox())
      ;
      try {
        var xdialog = new Dialog()
          .Add(new StaticText())
          .Add(new DriveDirBox())
        ;
        return false;
      }
      catch (ArgumentException) { }
      try {
        var xdialog = new Dialog()
          .Add(new TextEditBox())
          .Add(new DriveDirBox())
        ;
        return false;
      }
      catch (ArgumentException) { }
      return true;
    }

  }

  internal static class DialogShow
  {

    public static string RunTests() {
      string test = null;
      try {

        test = nameof(NoControls);
        if (!NoControls()) return test;

        test = nameof(Buttons);
        if (!Buttons()) return test;

        test = nameof(EditBoxes);
        if (!EditBoxes()) return test;

        test = nameof(Options);
        if (!Options()) return test;

        test = nameof(ListControls);
        if (!ListControls()) return test;

        test = nameof(Icons);
        if (!Icons()) return test;

        test = nameof(FileControls);
        if (!FileControls()) return test;

        return null;

      }
      catch (Exception ex) {
        throw new InvalidOperationException($"Unexpected failure: {nameof(DialogShow)}.{test}.", ex);
      }
    }

    public static bool NoControls() {
      try {
        var dialog = new Dialog();
        dialog.Show();
        return false;
      }
      catch (InvalidOperationException) { }
      return true;
    }
    public static bool Buttons() {
      var dialog = new Dialog { Title = "Buttons" }
        .Add(new StaticText { Height = 32, Width = 180, Text = "Press ESC or any Cancel\nbuttons to continue." })
        .Add(new OkButtonDef())
        .Add(new CancelButton())
        .Add(new OkButton())
        .Add(new CancelButtonDef())
      ;
      var ans = dialog.Show(d => false);
      return !ans;
    }
    public static bool EditBoxes() {
      var dialog = new Dialog { Width = 210, Title = "EditBoxes" }
        .Add(new StaticText { Text = "Press ESC to continue." })
        .Add(new TextEditBox())
        .Add(new IntegerEditBox())
        .Add(new NumberEditBox())
        .Add(new FormulaEditBox())
        .Add(new RefEditBox())
      ;
      var ans = dialog.Show(d => false);
      return !ans;
    }
    public static bool Options() {
      var dialog = new Dialog { Width = 210, Title = "Options" }
        .Add(new StaticText { Text = "Press ESC to continue." })
        .Add(new OptionGroup { Top = 30, Text = "Option Group" })
        .Add(new OptionButton { Text = "Option1" })
        .Add(new OptionButton { Text = "Option2" })
        .Add(new OptionButton { Text = "Option3" })
        .Add(new GroupBox { Height = 70, Text = "Group Box" })
        .Add(new CheckBox { Text = "Flag1", Value = true })
        .Add(new CheckBox { Text = "Flag2" })
        .Add(new CheckBox { Text = "Flag3", Value = null })
      ;
      var ans = dialog.Show(d => false);
      return !ans;
    }
    public static bool ListControls() {
      var items = new object[] { "Item1", "Item2", "Item3" };
      var dialog = new Dialog { Width = 350, Title = "ListControls" }
        .Add(new StaticText { Left = 10, Text = "Press ESC to continue." })
        .Add(new ListBox(items) { Left = 10, Top = 20 })
        .Add(new TextEditBox { Left = 180 })
        .Add(new LinkedListBox(items) { Height = 70 })
        .Add(new DropDown(items) { Left = 10, Top = 116 })
        .Add(new TextEditBox { Left = 180 })
        .Add(new LinkedDropDown(items))
      ;
      var ans = dialog.Show(d => false);
      return !ans;
    }
    public static bool Icons() {
      var dialog = new Dialog { Width = 210, Title = "Icons" }
        .Add(new StaticText { Text = "Press ESC to continue." })
        .Add(new Icon(Icon.Style.Question) { Left = 16, Top = 32 })
        .Add(new Icon(Icon.Style.Warning) { Left = 76 })
        .Add(new Icon(Icon.Style.Information) { Left = 138 })
      ;
      var ans = dialog.Show(d => false);
      return !ans;
    }
    public static bool FileControls() {
      var dialog = new Dialog { Width = 210, Title = "FileControls" }
        .Add(new StaticText { Text = "Press ESC to continue." })
        .Add(new TextEditBox())
        .Add(new FileListBox())
        .Add(new DriveDirBox())
        .Add(new DirectoryTextBox())
      ;
      var ans = dialog.Show(d => false);
      return !ans;
    }

  }

}
