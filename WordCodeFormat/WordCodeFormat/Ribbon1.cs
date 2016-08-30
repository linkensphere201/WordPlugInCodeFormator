using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;


//10.19:  关键字着色，英文后面的中文符号转英文例如括号等，大括号缩进，首行加上一个小“code”提示

namespace WordCodeFormat
{
    public partial class Ribbon1
    {
        
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
           // System.Windows.Forms.MessageBox.Show("hello");
            if (Globals.ThisAddIn._MyCustomTaskPane != null) {
                Globals.ThisAddIn._MyCustomTaskPane.Visible = true;
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn._MyCustomTaskPane != null)
            {
                Globals.ThisAddIn._MyCustomTaskPane.Visible = false;
            }
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.LastSelected != null)
            {
                //here we are going to make some different！
               
                /*
                *              sel.HighlightColorIndex = Word.WdColorIndex.wdGreen;

                                ------------set number format-------------**11.17**drop

                                11.17 drop
                                sel.ListFormat.ApplyNumberDefault();
                                try
                                {
                                    sel.ListFormat.ListTemplate.ListLevels[1].NumberFormat = "%1 .";
                                    设置编号：每次设置编号格式之前都必须有前面那句话。
                                }
                                catch
                                {
                                    MessageBox.Show("Format Setting Error!");
                                }
                11.17 add

                                -----type a string“<Code>”: 键入--------------
                                Globals.ThisAddIn.Application.Selection.End = sel.Start - 1;
                                Globals.ThisAddIn.Application.Selection.TypeParagraph();
                                Globals.ThisAddIn.Application.Selection.TypeText("<Code>");

                                try...catch 运行机制？
                                -------------set Font & Font Size
                               sel.Font.Name = "Courier New";
                                sel.Font.Size = 9;
                              ------------set shading format---------------
                                sel.ParagraphFormat.Shading.Texture = Word.WdTextureIndex.wdTextureNone;
                                sel.ParagraphFormat.Shading.ForegroundPatternColor = Word.WdColor.wdColorAutomatic;
                                int paraCount = sel.Paragraphs.Count;
                                while (paraCount > 0)
                                {
                                    sel.Paragraphs[paraCount].Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;
                                    if (paraCount - 1 > 0)
                                        sel.Paragraphs[--paraCount].Shading.BackgroundPatternColor = Word.WdColor.wdColorGray05;
                                    paraCount--;
                                }

                                sel.Start -= 7;
                                -----------set "<Code>" Color & Size
                               sel.Paragraphs[1].Range.Font.Size = 8;
                                sel.Paragraphs[1].Range.Font.ColorIndex = Word.WdColorIndex.wdBrightGreen;

                               ----------set paragraph line unit Before & After段前，段后---------
                               sel.Paragraphs[1].LineUnitBefore = sel.Paragraphs[sel.Paragraphs.Count].LineUnitAfter += 0.5F;//如果没有后缀，则表示一个双精度浮点数

                                MessageBox.Show(sel.Text);
                               ----------set Left Indent 缩进需要大改！检测括号---------------------
                                paraCount = sel.Paragraphs.Count;
                                while (paraCount > 0) {
                                    sel.Paragraphs[paraCount--].CharacterUnitLeftIndent += 1.0F;
                                }

                                内容匹配的均可以通过正则表达式实现 - 不过效率不知如何？
                                ----------set keywords color--------------------------------------
                                ----------set comment color----------------------------------
                       
                * */
                Word.Range sel = Globals.ThisAddIn.LastSelected;
                Myformat k = new Myformat();
                k.SetForm(sel);
               // k.SetColorForKeyWords(sel);
                //
            }
        }//End of Button3_Click


    }
}
