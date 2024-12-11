using System;
using ExcelDna.Integration;
using ExcelDna.Integration.Helpers.DialogBox;

namespace Dialogs
{
  [ExcelCommand(MenuName = "Dialogs")]
  public static class XlmDialogs
  {

    [ExcelCommand(MenuText = "Edit Controls")]
    public static void EditControls() {
      try {

        var dialog = new Dialog { Width = 290, Height = 182, Title = "Edit Controls Example", }
          .Add(new StaticText { Left = 10, Top = 5, Text = "This is static text, a.k.a. a label.", })
          .Add(new StaticText { Left = 10, Top = 26, Text = "Text:", })
          .Add(new TextEditBox { Left = 86, Top = 23, Width = 185, Value = "some text", })
          .Add(new StaticText { Left = 10, Top = 51, Text = "Integer:", })
          .Add(new IntegerEditBox { Left = 86, Top = 48, Width = 185, Value = 32767, })
          .Add(new StaticText { Left = 10, Top = 76, Text = "Number:", })
          .Add(new NumberEditBox { Left = 86, Top = 73, Width = 185, Value = 3.141592654, })
          .Add(new StaticText { Left = 10, Top = 101, Text = "Formula:", })
          .Add(new FormulaEditBox { Left = 86, Top = 98, Width = 185, Value = "=R1C1", })
          .Add(new StaticText { Left = 10, Top = 126, Text = "Reference:", })
          .Add(new RefEditBox { Left = 86, Top = 123, Width = 185, Value = "R2C2:R4C4", })
          .Add(new OkButtonDef { Left = 105, Top = 150, Width = 80, Height = 24, Text = "OK", })
        ;
        var ans = dialog.Show();

      }
      catch (Exception ex) {
        XlCall.Excel(XlCall.xlcAlert, ex.Message, 3);
      }
    }

    [ExcelCommand(MenuText = "Options")]
    public static void OptControls() {
      try {

        var dialog = new Dialog { Width = 302, Height = 116, Title = "Options & Checkboxes", }
          .Add(new OptionGroup("og") { Left = 8, Top = 1, Width = 140, Height = 74, Text = "Options", })
          .Add(new OptionButton { Left = 16, Top = 17, Text = "Option 1", })
          .Add(new OptionButton("ob") { Left = 16, Top = 35, Text = "Option 2", IsTrigger = true, })
          .Add(new OptionButton { Left = 16, Top = 53, Text = "Option 3", })
          .Add(new GroupBox { Left = 156, Top = 1, Width = 140, Height = 74, Text = "Check boxes", })
          .Add(new CheckBox { Left = 166, Top = 17, Text = "Check 1", Value = true })
          .Add(new CheckBox { Left = 166, Top = 35, Text = "Check 2", })
          .Add(new CheckBox("cb") { Left = 166, Top = 53, Text = "Check 3", IsTrigger = true, })
          .Add(new OkButtonDef { Left = 111, Top = 83, Width = 80, Height = 24, Text = "OK", })
        ;

        dialog.GetControl<OptionGroup>("og").SelectedIndex = 2;
        dialog.GetControl<CheckBox>("cb").Value = null;

        var ans = dialog.Show(d => {
          if (d.TriggerId == "ob") {
            XlCall.Excel(XlCall.xlcAlert, (d.TriggerControl as OptionButton).Text, 2);
            return false;
          }
          else if (d.TriggerId == "cb") {
            XlCall.Excel(XlCall.xlcAlert, "Tristate Checkbox", 2);
            (d.TriggerControl as CheckBox).Value = null;
            return false;
          }
          return true;
        });

      }
      catch (Exception ex) {
        XlCall.Excel(XlCall.xlcAlert, ex.Message, 3);
      }
    }

    [ExcelCommand(MenuText = "Icons")]
    public static void Icons() {
      try {

        var dialog = new Dialog { Width = 200, Title = "Icons", }
          .Add(new Icon(Icon.Style.Question) { Left = 20, Top = 10, })
          .Add(new Icon(Icon.Style.Information) { Left = 80, })
          .Add(new Icon(Icon.Style.Warning) { Left = 140, })
          .Add(new OkButton { Left = 55, Top = 50, Width = 88, Text = "OK", })
        ;

        var ans = dialog.Show();

      }
      catch (Exception ex) {
        XlCall.Excel(XlCall.xlcAlert, ex.Message, 3);
      }
    }

    [ExcelCommand(MenuText = "FileDialog")]
    public static void FileDialog() {
      try {

        var dialog = new Dialog { Width = 494, Height = 276, Title = "File Finder", }
          .Add(new GroupBox { Left = 10, Top = 8, Width = 472, Height = 45, Text = "Current directory at launch of dialog. Use ⭮ button to refresh.", })
          .Add(new OkButton { Left = 20, Top = 24, Width = 32, Text = "⭮", })
          .Add(new DirectoryTextBox { Left = 58, Top = 28, Width = 400, })
          .Add(new GroupBox { Left = 10, Top = 55, Width = 472, Height = 186, Text = "File selector. Use *.* to search for all files in a folder.", })
          .Add(new TextEditBox { Left = 20, Top = 72, Width = 220, })
          .Add(new FileListBox { Left = 250, Top = 72, Width = 220, Height = 160, })
          .Add(new DriveDirBox { Left = 20, Top = 96, Width = 220, Height = 136, })
          .Add(new OkButton { Left = 158, Top = 246, Width = 80, Text = "&OK", })
          .Add(new CancelButton { Left = 252, Top = 246, Width = 80, Text = "&Cancel", })
        ;

        var ans = dialog.Show();

      }
      catch (Exception ex) {
        XlCall.Excel(XlCall.xlcAlert, ex.Message, 3);
      }
    }

    [ExcelCommand(MenuText = "About")]
    public static void About() {
      try {

        var dialog = new Dialog { Width = 320, Height = 186, Title = "About Dialogs", }
          .Add(new GroupBox { Left = 10, Top = 5, Width = 300, Height = 140, Text = "Assembly Versions", })
          .Add(new ListBox("lb") { Left = 20, Top = 22, Width = 280, Height = 110, })
          .Add(new OkButtonDef { Left = 120, Top = 152, Width = 80, Height = 24, Text = "&OK", })
        ;

        var lb = dialog.GetControl<ListBox>("lb");

        var an = typeof(Ribbon).Assembly.GetName();
        lb.AddItem($"{an.Name}: {an.Version}");

        an = typeof(DnaLibrary).Assembly.GetName();
        lb.AddItem($"{an.Name}: {an.Version}");

        lb.SelectedIndex = null;

        var ans = dialog.Show();

      }
      catch (Exception ex) {
        XlCall.Excel(XlCall.xlcAlert, ex.Message, 3);
      }
    }

    [ExcelCommand(MenuText = "Version")]
    public static void Version() {
      try {

        var dialog = new Dialog { Width = 313, Height = 200, Title = "Version Info", }
          .Add(new GroupBox { Left = 13, Top = 13, Width = 287, Height = 130, Text = "Nonsense function library", })
          .Add(new StaticText { Left = 31, Top = 39, Text = "Library version", })
          .Add(new TextEditBox { Left = 31, Top = 58, Width = 250, Enabled = false, })
          .Add(new StaticText { Left = 31, Top = 91, Text = "Library compile date", })
          .Add(new TextEditBox { Left = 31, Top = 110, Width = 250, Enabled = false, })
          .Add(new OkButtonDef { Left = 31, Top = 160, Width = 100, Text = "&OK", })
        ;

        var ans = dialog.Show();

      }
      catch (Exception ex) {
        XlCall.Excel(XlCall.xlcAlert, ex.Message, 3);
      }
    }

  }
}
