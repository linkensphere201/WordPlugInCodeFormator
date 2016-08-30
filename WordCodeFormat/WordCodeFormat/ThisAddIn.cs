using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Tools;
using System.Windows.Forms;

namespace WordCodeFormat
{
    public partial class ThisAddIn
    {
        public CustomTaskPane _MyCustomTaskPane = null;
        public Word.Range LastSelected = null;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //
            //Globals.ThisAddIn.Application
            //MessageBox.Show("startup");
            UserControl1 taskpane = new UserControl1();
            _MyCustomTaskPane = this.CustomTaskPanes.Add(taskpane,"mytaskpane");
            this.Application.WindowSelectionChange += new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
            _MyCustomTaskPane.Width = 150;
            _MyCustomTaskPane.Visible = false;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                this.Application.WindowSelectionChange -= new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
            }
            catch 
            { }
        }
        //支持单一范围选择，假如按住ctrl，选了多个范围呢？ - 解决方法 ： 
        //1、现有的有没有类似Selections一个数组保存着我们要的多选的范围
        //2、没有的话，那我们只能自己每次Select行为产生后，自己进行多选的管理
        //目前先忽略多选，假设每次只选中一段话，认为它是一段代码，格式要变
        void Application_WindowSelectionChange(Word.Selection Sel) {
           // MessageBox.Show(Sel.Range.Start.ToString() +" "+ Sel.Range.End.ToString());
            if (Sel.Range.Start != Sel.Range.End)
            {
                LastSelected = this.Application.ActiveDocument.Range(Sel.Range.Start, Sel.Range.End);
               // MessageBox.Show(LastSelected.Start.ToString() + ":" + LastSelected.End.ToString());
            }
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
