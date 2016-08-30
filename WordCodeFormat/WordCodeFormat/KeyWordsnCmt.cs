using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace WordCodeFormat
{
    public class Myformat
    {
      private const string key_word_list = 
          "auto|struct|break|else|long|switch|case|enum|register|typedef|extern|return|union|const|unsigned|continue|for|signed|default|goto|sizeof|volatile| do|if|while|static";//关键字 蓝色
      private const string var_type_list =
          @"\s{1,}char|int|bool|float|double|size_t|string|short|void\s{1,}";//几个主要类型 深绿色


      //----------------------
      public Word.WdColor keyword_color = Word.WdColor.wdColorBlue;
      public Word.WdColor vartype_color = Word.WdColor.wdColorDarkGreen;

      public void SetColorForKeyWords(Word.Range sel)
      {
          //MessageBox.Show(keywordsarr.Length.ToString());
          //MessageBox.Show(vartypesarr.Length.ToString());
          MatchCollection keywords = Regex.Matches(sel.Text,key_word_list);
         // MessageBox.Show(keywords.Count.ToString());
          if (keywords.Count > 0) {
              foreach (Match _m_wd in keywords) {
                  try
                  {
                      Word.Range kw_rg = Globals.ThisAddIn.Application.ActiveDocument.Range(
                          sel.Start + _m_wd.Index, sel.Start + _m_wd.Index +_m_wd.Length
                          );
                      MessageBox.Show(kw_rg.Text);
                      kw_rg.Font.Color = keyword_color;
                      //这里的Font.Color居然TM没有在输入"."的时候弹出来！！MSDN上也没有！！是通过录制宏逆推出来的！#%￥...VSTO。。。

                  }
                  catch (Exception e) {
                      MessageBox.Show(e.Message);
                  }
              }
          }
          //----------
          MatchCollection vartypes = Regex.Matches(sel.Text, @"\s+char\s+");
          MessageBox.Show(keywords.Count.ToString());
          if (vartypes.Count > 0)
          {
              foreach (Match _v_tp in vartypes)
              {
                  try
                  {
                      Word.Range tp_rg = Globals.ThisAddIn.Application.ActiveDocument.Range(
                          sel.Start + _v_tp.Index, sel.Start + _v_tp.Index + _v_tp.Length
                          );
                      tp_rg.Font.Color = vartype_color;
                  }
                  catch (Exception e) {
                      MessageBox.Show(e.Message);
                  }
              }
          }
          //----------

      }
      public Word.Table SetForm(Word.Range sel) {

        //  MessageBox.Show(sel.Paragraphs.Count.ToString());
          int selS = sel.Start, selE = sel.End;
          int paraC = sel.Paragraphs.Count;
          Word.Application This_App = Globals.ThisAddIn.Application;
          Word.Selection This_Slt = This_App.Selection;
          Word.WdColor Bdr_cl = Word.WdColor.wdColorGreen;
          Word.WdLineStyle Bdr_ls = Word.WdLineStyle.wdLineStyleThinThickSmallGap;

          int where = sel.End;
          This_Slt.Start = where;
          This_Slt.End = where;
         // MessageBox.Show(sel.Paragraphs.Count.ToString());
          Word.Table _t = This_App.ActiveDocument.Tables.Add(This_Slt.Range, paraC + 1, 2);
         // MessageBox.Show(sel.Paragraphs.Count.ToString());//BUG： 只要选择的范围sel的范围是包括文档结尾，这里的paragraphs.count会因为table.add而改变
          _t.Rows[1].Select();
          This_Slt.Cells.Merge();
          _t.Cell(1, 1).Range.Text = "<Code: " + paraC.ToString() + "L>  ";
          //_t.Cell(1, 1).Range.Font.Name = "Courier New";
         
         // MessageBox.Show("1");
          for (int i = 2; i <= _t.Rows.Count; i++)
          {
              _t.Cell(i, 1).Range.Text = (i-1).ToString()+".";
              _t.Cell(i, 2).Range.Text = " " + sel.Paragraphs[i-1].Range.Text.TrimEnd("\r".ToCharArray());//经过测试发现，selection.Paragraphs[i].Range.text中结尾字符为换行'\r'而不是'\n'
             /* MessageBox.Show(sel.Paragraphs[i - 1].LeftIndent.ToString() + ":" + sel.Paragraphs[i - 1].CharacterUnitLeftIndent.ToString() + ":"
                 + sel.Paragraphs[i - 1].CharacterUnitFirstLineIndent.ToString() + ":" + sel.Paragraphs[i - 1].FirstLineIndent.ToString());*/
              _t.Cell(i, 2).LeftPadding = sel.Paragraphs[i - 1].FirstLineIndent + sel.Paragraphs[i - 1].LeftIndent;//缩进 -->ok -->？？自动检测代码块缩进 -->？？自动检测括号
              //这里Borders为类似数组的结构
              //_t.Cell(i, 2).Borders[Word.WdBorderType.BorderUp].LineStyle = ...
              //VB宏类似这样：  Selection.Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
          }
          //_t.Columns[1].Select();//选中第一列
          
          //This_Slt.Borders[Word.WdBorderType.wdBorderRight].Color = Word.WdColor.wdColorOliveGreen; **Gill Sans MT Condensed
          //结果显示，**标注的语句执行没有改变效果(在第一个表格生成之后)，第二个表格则可用 
          //--> 不仅跟第二个表格有关，还跟颜色有关，目前只有wdColorGreen/Orange/Red可以用 
          //--> 仅仅与颜色有关,green就可以，olivegreen不行。
          _t.Cell(1, 1).Range.Font.Name = "Century Gothic";
          _t.Cell(1, 1).Range.Font.Size = (float)9;
          _t.Cell(1, 1).Range.Font.Bold = 1;
          _t.Cell(1, 1).Range.Shading.Texture = Word.WdTextureIndex.wdTextureNone;
          _t.Cell(1, 1).Range.Shading.ForegroundPatternColor = Word.WdColor.wdColorAutomatic;//
          _t.Cell(1, 1).Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray20;
          //MessageBox.Show(_t.Cell(1, 1).Range.Start.ToString() + "1" + _t.Cell(1, 1).Range.End.ToString());
          Word.Range pic = This_App.ActiveDocument.Range(_t.Cell(1, 1).Range.End - 1,_t.Cell(1, 1).Range.End - 1);
          //pic.InlineShapes.AddPicture("../../resource/th.png",false,true); 《== 应该将图片作为一个资源，以资源形式添加到工程中，每次以资源形式访问。
          //MessageBox.Show("1");
          for (int i = 2; i <= _t.Rows.Count; i++)
          {
              _t.Cell(i, 1).Borders[Word.WdBorderType.wdBorderRight].LineStyle = Bdr_ls;
              _t.Cell(i, 1).Borders[Word.WdBorderType.wdBorderRight].Color = Bdr_cl;

              _t.Cell(i, 1).Range.Font.Name = "Century Gothic";
              _t.Cell(i, 1).Range.Font.Size = (float)8;

              _t.Cell(i, 2).Range.Font.Name = "Consolas";
              _t.Cell(i, 2).Range.Font.Size = (float)8;
              //----set number shading color
              _t.Cell(i, 1).Range.Shading.Texture = Word.WdTextureIndex.wdTextureNone;
              _t.Cell(i, 1).Range.Shading.ForegroundPatternColor = Word.WdColor.wdColorAutomatic;
              _t.Cell(i, 1).Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray30;
              //----set shading colot end
              _t.Cell(i, 2).Range.Shading.Texture = Word.WdTextureIndex.wdTextureNone;
              _t.Cell(i, 2).Range.Shading.ForegroundPatternColor = Word.WdColor.wdColorAutomatic;
              _t.Cell(i, 2).Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
              if (i % 2 == 1) {
                  _t.Cell(i, 2).Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray20;
              }
          }
          _t.Rows.SetLeftIndent(21,Word.WdRulerStyle.wdAdjustNone);//整个表格设置左缩进
          _t.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);//调整宽度
          This_App.ActiveDocument.Range(selS,selE).Delete();

          return _t;
      }
    }
}
