using System;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Excel = Microsoft.Office.Interop.Excel;
using WinForms = System.Windows.Forms;

namespace Dialogs
{
  /// <summary>
  /// A dialog box can be invoked from the ribbon calling the 'Application.Run'
  /// or the 'ExcelAsyncUtil.QueueMacro' methods.
  /// </summary>
  [ComVisible(true)]
  [ProgId("Dialogs.Ribbon")]
  [Guid("9B58B466-20CD-43F3-A73A-61522617AD7F")]
  public class Ribbon : ExcelRibbon
  {

    public Ribbon() {
      var asm = typeof(Ribbon).Assembly;
      FriendlyName = (asm.GetCustomAttributes(typeof(AssemblyTitleAttribute),false)[0] as AssemblyTitleAttribute).Title;
      Description = (asm.GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false)[0] as AssemblyDescriptionAttribute).Description;
    }

    public override string GetCustomUI(string RibbonID) {
      var ribbon = @"
        <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
          <ribbon>
            <tabs>
              <tab id='Dialogs' label='Dialogs'>
                <group id='Runners' label='Show Dialog'>
                  <button id='Runner.Run' onAction='OnRunner' label='AppRun' size='large' imageMso='MacroPlay' />
                  <button id='Runner.QAM' onAction='OnRunner' label='Queue' size='large' imageMso='MacroPlay' />
                </group>
              </tab>
            </tabs>
          </ribbon>
        </customUI>
      ";
      return ribbon;
    }

    public void OnRunner(IRibbonControl control) {
      try {
        var id = control.Id;
        var macro = "About";
        if (id.EndsWith("Run")) {
          var excel = ExcelDnaUtil.Application as Excel.Application;
          excel.Run(macro);
        }
        else if (id.EndsWith("QAM")) {
          ExcelAsyncUtil.QueueMacro(macro);
        }
      }
      catch (Exception ex) {
        WinForms.MessageBox.Show($"ERROR: {ex.Message}", "OnRunner", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
      }
    }

  }
}
