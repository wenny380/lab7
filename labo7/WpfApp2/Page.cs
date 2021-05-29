﻿using System;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;

namespace WpfApp2
{
    class Page
    {
        private Word.Application objword;

        public Page()
        {
            objword = new Word.Application(); // Create an object for working with Word
            objword.Visible = true;
            objword.WindowState = Word.WdWindowState.wdWindowStateNormal; // Normal window position
        }

        public void SetParagraph(Word.Paragraph objpara)
        {
            objpara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            objpara.Range.Font.Name = "Times New Roman"; // font name
            objpara.Range.Font.Size = 14; //font size
            objpara.Range.Paragraphs.Space1(); // single spacing
        }

        public void FormTitlePage(string cafedra, int labNum, string theme, string discipline, string student, string teacher, int year)
        {
            // Department
            // Lab number
            //Subject
            //Discipline
            //Student
            //Teacher
            //Year
            try
            {
                Word.Document objdoc = objword.Documents.Add();
                Word.Paragraph objpara;
                objpara = objdoc.Paragraphs.Add();
                objpara.Range.Text = "МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ" +
                    "   РОССИЙСКОЙ ФЕДЕРАЦИИ";
                objpara = objdoc.Paragraphs.Add();
                SetParagraph(objpara);

                objpara = objdoc.Paragraphs.Add();
                objpara.Range.Text = "ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ ВЫСШЕГО ОБРАЗОВАНИЯ" +
                    " «ОРЛОВСКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ";
                objpara = objdoc.Paragraphs.Add();
                SetParagraph(objpara);

                objpara = objdoc.Paragraphs.Add();
                objpara.Range.Text = " ИМЕНИ И.С. ТУРГЕНЕВА»";
                objpara = objdoc.Paragraphs.Add();
                SetParagraph(objpara);

                objpara = objdoc.Paragraphs.Add();
                objpara.Range.Text = $"\nКафедра {cafedra}";
                objpara = objdoc.Paragraphs.Add();
                SetParagraph(objpara);

                objpara = objdoc.Paragraphs.Add();
                objpara.Range.Text = "\n\n\nОТЧЕТ";
                objpara = objdoc.Paragraphs.Add();
                SetParagraph(objpara);
                objpara.Range.Bold = 2;

                objpara = objdoc.Paragraphs.Add();
                objpara.Range.Text = $"По лабораторной работе №{labNum}";
                objpara = objdoc.Paragraphs.Add();
                SetParagraph(objpara);

                objpara = objdoc.Paragraphs.Add();
                objpara.Range.Text = $"на тему: «{theme}»";
                objpara = objdoc.Paragraphs.Add();
                SetParagraph(objpara);

                objpara = objdoc.Paragraphs.Add();
                objpara.Range.Text = $"по дисциплине: «{discipline}»";
                objpara = objdoc.Paragraphs.Add();
                SetParagraph(objpara);

                objpara = objdoc.Paragraphs.Add();
                objpara.Range.Text = $"\n\n\n\nВыполнил: {student}";
                objpara = objdoc.Paragraphs.Add();
                SetParagraph(objpara);
                objpara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

                objpara = objdoc.Paragraphs.Add();
                objpara.Range.Text = "Институт приборостроения, автоматизации и информационных технологий" +
                    " Направление: 09.03.04 «Программная инженерия»";
                objpara = objdoc.Paragraphs.Add();
                SetParagraph(objpara);
                objpara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

                objpara = objdoc.Paragraphs.Add();
                objpara.Range.Text = "Группа: 92-ПГ";
                objpara = objdoc.Paragraphs.Add();
                SetParagraph(objpara);
                objpara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

                objpara = objdoc.Paragraphs.Add();
                objpara.Range.Text = $"Проверил: {teacher}";
                objpara = objdoc.Paragraphs.Add();
                SetParagraph(objpara);
                objpara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

                objpara = objdoc.Paragraphs.Add();
                objpara.Range.Text = $"\nОтметка о зачете:                                    Дата: «____» __________ {year} г.";
                objpara = objdoc.Paragraphs.Add();
                SetParagraph(objpara);
                objpara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

                objpara = objdoc.Paragraphs.Add();
                objpara.Range.Text = $"\n\n\n\n\nОрел, {year}";
                objpara = objdoc.Paragraphs.Add();
                SetParagraph(objpara);

                objdoc.SaveAs("C:\\MSDOC.docx"); //Сохранить файл
                objdoc.Close();
                objword.Quit();
            }
            catch(Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }
    }
}
