using Microsoft.Office.Tools.Ribbon;
using NHunspell;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace KhmerSpellCheck
{
    public partial class KhmerSpellCheck
    {
        //The word application
        Word.Application wordApp;
        List<IgnoreWord> ignoreWords = new List<IgnoreWord>();
        List<IgnoreWord> ignoreAllWords = new List<IgnoreWord>();
        char[] KhmerCharacters = new char[] {'\u0041','\u0042','\u0043','\u0044', '\u0045', '\u0046', '\u0047',
            '\u0048', '\u0049', '\u004A', '\u004B', '\u004C', '\u004D', '\u004E', '\u004F', '\u0050', '\u0051', '\u0052',
            '\u0053', '\u0054', '\u0055', '\u0056', '\u0057', '\u0058', '\u0059', '\u005A', '\u005B', '\u005C', '\u005D',
            '\u0061', '\u0062', '\u0063', '\u0064', '\u0065', '\u0066', '\u0067', '\u0068', '\u0069', '\u006A', '\u006B',
            '\u006C', '\u006D', '\u006E', '\u006F', '\u0070', '\u0071', '\u0072', '\u0073', '\u0074', '\u0075', '\u0076',
            '\u0077', '\u0078', '\u0079', '\u007A', '\u007B', '\u007C', '\u007D'};

       
        char[] CommonEnglishPunctuations = new char[] {
            '~', '`', '!', '@', '#', '$', '%', '^', '&', '*', '(', ')', '-', '_', '=', '+', '{', '[', '}', ']',
            '|', '\\', ':', ';', '"', '\'', '<', ',', '>', '.', '?', '/', '1', '2', '3', '4', '5', '6', '7',
            '8', '9', '0', '•', '‣', '‥', '…', '«', '»', '£', '¥', '©', '®', '¶', '·'};

        private void KhmerSpellCheck_Load(object sender, RibbonUIEventArgs e)
        {
            //Get the word application object
            wordApp = Globals.ThisAddIn.Application;
        }

        private void BtnSpellCheck_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //Exit if there is no active document
                if (wordApp.ActiveDocument == null)
                    return;

                //Get all the words from the active document
                Word.Document doc = wordApp.ActiveDocument;
                //Word.Range entirerange = activeDoc.Range();
                //string[] KhmerWords = entirerange.Text.Split(new char[] { '\u200B', ' ', '\r', '\a', '.' }, StringSplitOptions.RemoveEmptyEntries);

                #region Locate Dictionary files
                //Get the deployment directory
                System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();

                //Location is where the assembly is run from 
                string assemblyLocation = assemblyInfo.Location;

                //CodeBase is the location of the ClickOnce deployment files
                Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);
                string InstallationLocation = Path.GetDirectoryName(uriCodeBase.LocalPath.ToString());

                //Khmer dictionaries
                string affFile = Path.Combine(InstallationLocation, "ia.aff");
                string dictFile = Path.Combine(InstallationLocation, "ia.dic");
                #endregion

                //Loop through all the words to find the first mis-spell
                using (Hunspell hunspell = new Hunspell(affFile, dictFile))
                {
                    //Process all Paragraphs in the documents
                    Object oMissing = System.Reflection.Missing.Value;
                    object WdLine = Microsoft.Office.Interop.Word.WdUnits.wdLine; // change a line; 
                    object moveExtend = Microsoft.Office.Interop.Word.WdMovementType.wdExtend;

                    //bool isInitial = true; //Is first line

                    doc.ActiveWindow.Selection.HomeKey(Word.WdUnits.wdStory, ref oMissing);

                    ////Read the entire document
                    //while (doc.ActiveWindow.Selection.Bookmarks.Exists(@"\EndOfDoc") == false)
                    //{
                    //    if (!isInitial)
                    //    {
                    //        doc.ActiveWindow.Selection.MoveDown(ref WdLine, ref oMissing, ref oMissing);
                    //        doc.ActiveWindow.Selection.HomeKey(ref WdLine, ref oMissing);
                    //    }

                    //    isInitial = false;

                    //    //Skiping table content
                    //    if (doc.ActiveWindow.Selection.get_Information(Word.WdInformation.wdEndOfRangeColumnNumber).ToString() != "-1")
                    //    {
                    //        while (doc.ActiveWindow.Selection.get_Information(Word.WdInformation.wdEndOfRangeColumnNumber).ToString() != "-1")
                    //        {
                    //            if (doc.ActiveWindow.Selection.Bookmarks.Exists(@"\EndOfDoc"))
                    //                break;

                    //            doc.ActiveWindow.Selection.MoveDown(ref WdLine, ref oMissing, ref oMissing);
                    //            doc.ActiveWindow.Selection.HomeKey(ref WdLine, ref oMissing);
                    //        }
                    //        doc.ActiveWindow.Selection.HomeKey(ref WdLine, ref oMissing);
                    //    }

                    //    //Select the line and get all the khmer words from the line
                    //    doc.ActiveWindow.Selection.EndKey(ref WdLine, ref moveExtend);

                    //    //Set the repeatcheck and stopcheck
                    //    bool repeatcheck, stopcheck;

                    //    //Check Khmer Check
                    //    do
                    //    {
                    //        string SelectedText = doc.ActiveWindow.Selection.Text;
                    //        string[] KhmerWords = GetKhmerWords(SelectedText);

                    //        CheckWords(doc, hunspell, SelectedText, KhmerWords, out stopcheck, out repeatcheck);

                    //        //Break if user selects to exit
                    //        if (stopcheck) return;
                    //    }
                    //    while (repeatcheck);
                    //}

                    foreach (Word.Paragraph para in doc.Paragraphs)
                    {
                        para.Range.Select();

                        //Set the repeatcheck and stopcheck
                        bool repeatcheck, stopcheck;

                        //Check Khmer Check
                        do
                        {
                            string SelectedText = doc.ActiveWindow.Selection.Text;
                            string[] KhmerWords = GetKhmerWords(SelectedText);

                            CheckWords(doc, hunspell, SelectedText, KhmerWords, out stopcheck, out repeatcheck);

                            //Break if user selects to exit
                            if (stopcheck) return;
                        }
                        while (repeatcheck);
                    }

                    ////Processing all tables in the documents
                    //for (int iCounter = 1; iCounter <= doc.Tables.Count; iCounter++)
                    //{
                    //    foreach (Word.Row aRow in doc.Tables[iCounter].Rows)
                    //    {
                    //        foreach (Word.Cell aCell in aRow.Cells)
                    //        {
                    //            aCell.Select();

                    //            //Set the repeatcheck and stopcheck
                    //            bool repeatcheck, stopcheck;

                    //            //Check Khmer Check
                    //            do
                    //            {
                    //                //Get the selected text and khmer words
                    //                string SelectedText = aCell.Range.Text;
                    //                string[] KhmerWords = GetKhmerWords(SelectedText);

                    //                CheckWords(doc, hunspell, SelectedText, KhmerWords, out stopcheck, out repeatcheck, "table", aCell);

                    //                //Break if user selects to exit
                    //                if (stopcheck) return;
                    //            }
                    //            while (repeatcheck);
                    //        }
                    //    }
                    //}

                    var shapes = doc.Shapes;
                    //Finds text within textboxes, then changes them
                    foreach (Microsoft.Office.Interop.Word.Shape shape in shapes)
                    {
                        shape.Select();

                        //Set the repeatcheck and stopcheck
                        bool repeatcheck, stopcheck;

                        //Check Khmer Check
                        do
                        {
                            //Get the selected text and khmer words
                            string SelectedText = shape.TextFrame.TextRange.Text;
                            string[] KhmerWords = GetKhmerWords(SelectedText);

                            CheckWords(doc, hunspell, SelectedText, KhmerWords, out stopcheck, out repeatcheck, "shape", shape);

                            //Break if user selects to exit
                            if (stopcheck) return;
                        }
                        while (repeatcheck);
                    }


                    MessageBox.Show("Khmer Spelling Check is complete", "Khmer Spell Check", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void CheckWords(Word.Document doc, Hunspell hunspell, string SelectedText, string[] KhmerWords, out bool stopcheck, out bool repeatcheck, string objecttype = "", object wordobject = null)
        {
            int StartPosition = 0;
            Object oMissing = System.Reflection.Missing.Value;
            stopcheck = repeatcheck = false;

            //Check all the Khmer words from the selected line
            foreach (string KhmerWord in KhmerWords)
            {
                DialogResult dialogResult = DialogResult.None;
                frmKhmer frmKhmer = null;
                String newKhmerWord = String.Empty;

                if (!hunspell.Spell(KhmerWord))
                {
                    if (!ignoreAllWords.Any(ignoreAllWord => ignoreAllWord.KhmerWord == KhmerWord))
                    {
                        if (!ignoreWords.Contains(new IgnoreWord { document = doc.Name, KhmerWord = KhmerWord, SelectedText = SelectedText, StartPosition = StartPosition, IgnoreAll = false }))
                        {
                            Word.Range start = null;
                            Word.WdColorIndex highlightcolorindex = Word.WdColorIndex.wdNoHighlight;
                            Word.WdUnderline fontunderline = Word.WdUnderline.wdUnderlineNone;
                            Word.WdColor fontcolor = Word.WdColor.wdColorBlack;
                            Word.Range selectionRange = null;

                            //Select the erroneous word on the main document
                            if (String.IsNullOrWhiteSpace(objecttype))
                            {
                                //Set the initial selection
                                start = doc.ActiveWindow.Selection.Range;

                                //Set the search area
                                doc.ActiveWindow.Selection.Start += StartPosition;
                                Word.Selection searchArea = doc.ActiveWindow.Selection;

                                //Set the find object
                                Word.Find findObject = searchArea.Find;
                                findObject.ClearFormatting();
                                findObject.Text = KhmerWord;


                                //Find the mis-spelled word
                                findObject.Execute(ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                                //Temp store the current formatting
                                highlightcolorindex = doc.ActiveWindow.Selection.Range.HighlightColorIndex;
                                fontunderline = doc.ActiveWindow.Selection.Range.Font.Underline;
                                fontcolor = doc.ActiveWindow.Selection.Range.Font.UnderlineColor;

                                //Highlight the selection
                                doc.ActiveWindow.Selection.Range.HighlightColorIndex = Word.WdColorIndex.wdYellow;
                                doc.ActiveWindow.Selection.Range.Font.Underline = Word.WdUnderline.wdUnderlineWavy;
                                doc.ActiveWindow.Selection.Range.Font.UnderlineColor = Word.WdColor.wdColorRed;
                                selectionRange = doc.ActiveWindow.Selection.Range;
                                doc.ActiveWindow.Selection.Collapse();
                            }
                            else
                            {
                                if (objecttype == "table")
                                {
                                    start = ((Word.Cell)wordobject).Range;
                                }
                                else if (objecttype == "shape")
                                {
                                    start = ((Word.Shape)wordobject).TextFrame.TextRange;
                                    start.Start += StartPosition;
                                }

                                //Set the find object
                                Word.Find findObject = start.Find;
                                findObject.ClearFormatting();
                                findObject.Text = KhmerWord;

                                //Temp store the current formatting
                                highlightcolorindex = start.HighlightColorIndex;
                                fontunderline = start.Font.Underline;
                                fontcolor = start.Font.UnderlineColor;

                                //Find the mis-spelled word
                                findObject.Execute(ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                                //Highlight the selection
                                start.HighlightColorIndex = Word.WdColorIndex.wdYellow;
                                start.Font.Underline = Word.WdUnderline.wdUnderlineWavy;
                                start.Font.UnderlineColor = Word.WdColor.wdColorRed;
                                start.Select();
                            }

                            bool isObject = !String.IsNullOrWhiteSpace(objecttype);
                            frmKhmer = new frmKhmer(SelectedText, KhmerWord, StartPosition, hunspell.Suggest(KhmerWord), isObject);
                            dialogResult = frmKhmer.ShowDialog();

                            //Select the line again
                            if (String.IsNullOrWhiteSpace(objecttype))
                            {
                                //Revert the highlights
                                selectionRange.Select();
                                doc.ActiveWindow.Selection.Range.HighlightColorIndex = highlightcolorindex;
                                doc.ActiveWindow.Selection.Range.Font.Underline = fontunderline;
                                doc.ActiveWindow.Selection.Range.Font.UnderlineColor = fontcolor;

                                if (dialogResult != DialogResult.Abort) start.Select();
                            }
                            else
                            {
                                start.HighlightColorIndex = highlightcolorindex;
                                start.Font.Underline = fontunderline;
                                start.Font.UnderlineColor = fontcolor;

                                if (dialogResult != DialogResult.Abort)
                                {
                                    if (objecttype == "table")
                                    {
                                        ((Word.Cell)wordobject).Select();
                                    }
                                    else if (objecttype == "shape")
                                    {
                                        ((Word.Shape)wordobject).Select();
                                    }
                                }
                            }
                        }
                    }
                }

                #region Cancel Button Clicked
                //Return if the user hits Cancel Button
                if (dialogResult == DialogResult.Cancel || dialogResult == DialogResult.Abort)
                {
                    stopcheck = true;
                    repeatcheck = false;
                    return;
                }
                #endregion

                #region Ignore or Ignore All Clicked
                //Ignore the word
                if (dialogResult == DialogResult.Ignore)
                {
                    if (frmKhmer.IgnoreAll)
                    {
                        ignoreAllWords.Add(new IgnoreWord { KhmerWord = KhmerWord, IgnoreAll = frmKhmer.IgnoreAll });
                    }
                    else
                    {
                        ignoreWords.Add(new IgnoreWord { document = doc.Name, KhmerWord = KhmerWord, SelectedText = SelectedText, StartPosition = StartPosition });
                    }
                }
                #endregion

                #region Change or Change All Clicked
                if (dialogResult == DialogResult.Yes)
                {
                    if (String.IsNullOrWhiteSpace(objecttype))
                    {
                        //Set the initial selection
                        Word.Range start = doc.ActiveWindow.Selection.Range;

                        //Set the searcharea
                        if (frmKhmer.changeAll)
                        {
                            doc.Content.Select();
                        }
                        Word.Selection searchArea = doc.ActiveWindow.Selection;

                        //Set the find object
                        Word.Find findObject = searchArea.Find;
                        findObject.ClearFormatting();
                        findObject.Text = KhmerWord;
                        findObject.Replacement.ClearFormatting();
                        findObject.Replacement.Text = frmKhmer.selectedSuggestion;

                        object replaceAll = frmKhmer.changeAll ? Word.WdReplace.wdReplaceAll : Word.WdReplace.wdReplaceOne;

                        findObject.Execute(ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                            ref replaceAll, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                        newKhmerWord = frmKhmer.selectedSuggestion;

                        //Set back the selection
                        start.Select();

                        //Set repeatcheck to true
                        if (frmKhmer.changeAll)
                        {
                            stopcheck = false;
                            repeatcheck = true;
                            return;
                        }
                    }
                    else
                    {
                        var resultingText = SelectedText.Replace(KhmerWord, frmKhmer.selectedSuggestion);

                        if (objecttype == "table")
                        {
                            Word.Range range = ((Word.Cell)wordobject).Range;
                            range.Text = resultingText;
                        }
                        else if (objecttype == "shape")
                        {
                            Word.Shape shape = (Word.Shape)wordobject;
                            shape.TextFrame.TextRange.Text = resultingText;
                        }

                        stopcheck = false;
                        repeatcheck = true;
                        return;
                    }
                }
                #endregion

                StartPosition += String.IsNullOrWhiteSpace(newKhmerWord) ? KhmerWord.Length : newKhmerWord.Length;
            }
        }

        private string[] GetKhmerWords(string SelectedText)
        {
            string[] KhmerWords =
                SelectedText
                    .Split(new char[] { '\u200B', '\u200C', ' ', '\r', '\a', '.', '\t', '\v' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(KhmerWord => KhmerWord.Trim(CommonEnglishPunctuations))
                    .Where(KhmerWord => String.IsNullOrWhiteSpace(KhmerWord) == false)
                    .Where(KhmerWord => KhmerCharacters.Contains(KhmerWord[0]))
                    .ToArray();
            return KhmerWords;
        }
    }
}
