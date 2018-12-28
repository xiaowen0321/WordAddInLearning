using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;


namespace WordAddInLearning
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var doc = Globals.ThisAddIn.Application.ActiveDocument;

            switch (editBox1.Text)
            {
                case "1":
                    //选中0-7之间的7个字符
                    doc.Range(0, 7).Select();
                    break;
                case "2":
                    //选中所有字符
                    doc.Content.Select();
                    break;
                case "3":
                    //选中第二个句子
                    doc.Sentences[1].Select();
                    break;
                case "4":
                    //选中第一和第二个句子
                    var sentences = doc.Sentences;
                    if (sentences.Count >= 2)
                    {
                        doc.Range(sentences[1].Start,sentences[2].End).Select();
                    }
                    break;
                case "5":
                    //获取第二个句子的起止范围
                    var rang = doc.Sentences[2];
                    System.Windows.Forms.MessageBox.Show($"{rang.Start}+{rang.End}");
                    break;
                case "6":
                    ReplaceParagraphText();
                    break;
                case "7":
                    ExpandRange();
                    break;
                case "8":
                    ResetRange();
                    break;
                case "9":
                    CollapseRange();
                    break;
                case "10":
                    SelectionInsertText();
                    break;
                case "11":
                    RangeFormat();
                    break;
                case "12":
                    GetAllStyles();
                    break;
                case "13":
                    CreateCustomerStyle();
                    break;
                case "14":
                    CreateBookMark();
                    break;
                default:
                    break;
            }
            

            
            
        }


        private void ReplaceParagraphText()
        {
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            var firstRange = doc.Paragraphs[1].Range;
            var secondRange = doc.Paragraphs[2].Range;

            var firstString = firstRange.Text;
            var secondString = secondRange.Text;

            firstRange.Select();
            MessageBox.Show(firstRange.Text);
            secondRange.Select();
            MessageBox.Show(secondRange.Text);

            object charUnit = Word.WdUnits.wdCharacter;
            object move = -1;

            firstRange.MoveEnd(ref charUnit, ref move);

            firstRange.Text = $"第一段新内容——{DateTime.Now.ToString()}";
            secondRange.Text = $"第二段新内容——{DateTime.Now.ToString()}";

            firstRange.Select();
            MessageBox.Show(firstRange.Text);
            secondRange.Select();
            MessageBox.Show(secondRange.Text);

            move = 1;
            firstRange.MoveEnd(ref charUnit, ref move);
            secondRange.Delete();
            firstRange.Text = firstString;
            firstRange.InsertAfter(secondString);
            firstRange.Select();


        }

        private void ExpandRange()
        {
            var range = Globals.ThisAddIn.Application.ActiveDocument.Range(0, 7);
            //起始位置向右移7个字符
            range.MoveStart(Word.WdUnits.wdCharacter, 7);
            //终止位置向右移7个字符
            range.MoveEnd(Word.WdUnits.wdCharacter, 7);
            range.Select();
        }

        private void ResetRange()
        {
            var range = Globals.ThisAddIn.Application.ActiveDocument.Range(0, 7);
            range.SetRange(Globals.ThisAddIn.Application.ActiveDocument.Sentences[2].Start,
                Globals.ThisAddIn.Application.ActiveDocument.Sentences[5].End);
            range.Select();
        }

        private void CollapseRange()
        {
            var range = Globals.ThisAddIn.Application.ActiveDocument.Paragraphs[1].Range;
            range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            range.Text = "段前新文本";
            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            range.Text = "段后新文本";
            range.Select();
        }

        private void SelectionInsertText()
        {
            var currentSelection = Globals.ThisAddIn.Application.Selection;
            var userOvertype = Globals.ThisAddIn.Application.Options.Overtype;
            if (Globals.ThisAddIn.Application.Options.Overtype)
            {
                Globals.ThisAddIn.Application.Options.Overtype = false;
            }
            if (currentSelection.Type == Word.WdSelectionType.wdSelectionIP)
            {
                currentSelection.TypeText("《插入的新文本》");
                currentSelection.TypeParagraph();//插入新行
            }
            else if (currentSelection.Type == Word.WdSelectionType.wdSelectionNormal)
            {
                if (Globals.ThisAddIn.Application.Options.ReplaceSelection)
                {
                    object direction = Word.WdCollapseDirection.wdCollapseStart;
                    currentSelection.Collapse(ref direction);
                }
                currentSelection.TypeText("《插入的新文本》");
                currentSelection.TypeParagraph();
            }
            else
            {

            }
            //还原用户初始设置
            Globals.ThisAddIn.Application.Options.Overtype = userOvertype;
        }

        private void RangeFormat()
        {
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            var range = doc.Paragraphs[1].Range;

            range.Font.Size = 14;
            range.Font.Name = "宋体";
            range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

            range.Select();
            MessageBox.Show("格式化第一段");

            object numTimes = 3;
            doc.Undo(ref numTimes);
            range.Select();
            MessageBox.Show("撤销 3 步操作");

            object indentStyle = "引用";
            range.set_Style(ref indentStyle);

            range.Select();
            MessageBox.Show("应用样式");

            object numTimes1 = 1;
            doc.Undo(ref numTimes1);

            range.Select();
            MessageBox.Show("撤销 1 步");
        }

        private void GetAllStyles()
        {
            //var styleString = Globals.ThisAddIn.Application.ActiveDocument.Paragraphs[1].Range.get_Style();
            //MessageBox.Show(styleString.NameLocal);

            var selection = Globals.ThisAddIn.Application.Selection;
            foreach (Word.Style style in Globals.ThisAddIn.Application.ActiveDocument.Styles)
            {
                if (style.Type == Word.WdStyleType.wdStyleTypeParagraph)
                {
                    selection.TypeText(style.NameLocal + "\t");
                }
            }
        }

        private void CreateCustomerStyle()
        {
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            doc.Styles.Add("我的样式", Word.WdStyleType.wdStyleTypeParagraph);
            Word.Style myStyle = doc.Styles["我的样式"];
            myStyle.AutomaticallyUpdate = true;
            myStyle.Font.Name = "楷体";
            myStyle.Font.Size = 22;
            myStyle.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
        }

        private void CreateBookMark()
        {
            Globals.ThisAddIn.Application.ActiveDocument.Range(0, 7).Bookmarks.Add("测试书签");
            
        }
    }
}
