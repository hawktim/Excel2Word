using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Excel2Word.Entities;
using System.Collections.Generic;

namespace Excel2Word.Activity
{
    public class BuildWord : GeneratedClass
    {
        private List<Department> _data;

        public void WriteFile(List<Department> data, string filePath)
        {
            _data = data;
            CreatePackage(filePath);
        }
        protected override void DataBinding(Table table1)
        {
            base.DataBinding(table1);

            foreach (var department in _data)
            {
                table1.Append(AppendDepartment(department.DepartmentName, department.AllProblem));

                foreach (var staff in department.Staffs)
                    table1.Append(AppendStaff(staff.ToString(), staff.CountProblem));
            }
        }

        private TableRow AppendDepartment(string departmentName, int allProblem)
        {
            TableRow tableRow2 = new TableRow() { RsidTableRowAddition = "003B602A", RsidTableRowProperties = "003B602A" };

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "4672", Type = TableWidthUnitValues.Dxa };
            Shading shading3 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D9D9D9", ThemeFill = ThemeColorValues.Background1, ThemeFillShade = "D9" };

            tableCellProperties3.Append(tableCellWidth3);
            tableCellProperties3.Append(shading3);

            Paragraph paragraph5 = new Paragraph() { RsidParagraphMarkRevision = "003B602A", RsidParagraphAddition = "003B602A", RsidParagraphProperties = "004D2B18", RsidRunAdditionDefault = "003B602A" };
            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
            paragraphMarkRunProperties4.Append(new Bold());
            paragraphProperties4.Append(paragraphMarkRunProperties4);
            Run run4 = new Run() { RsidRunProperties = "003B602A" };

            RunProperties runProperties4 = new RunProperties();
            runProperties4.Append(new Bold());
            Text text4 = new Text
            {
                Space = SpaceProcessingModeValues.Preserve,
                Text = departmentName
            };

            run4.Append(runProperties4);
            run4.Append(text4);

            paragraph5.Append(paragraphProperties4);
            paragraph5.Append(run4);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph5);

            TableCell tableCell4 = new TableCell();

            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "4673", Type = TableWidthUnitValues.Dxa };
            Shading shading4 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D9D9D9", ThemeFill = ThemeColorValues.Background1, ThemeFillShade = "D9" };

            tableCellProperties4.Append(tableCellWidth4);
            tableCellProperties4.Append(shading4);

            Paragraph paragraph6 = new Paragraph() { RsidParagraphMarkRevision = "003B602A", RsidParagraphAddition = "003B602A", RsidParagraphProperties = "004641AF", RsidRunAdditionDefault = "003B602A" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            Justification justification4 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
            Bold bold8 = new Bold();

            paragraphMarkRunProperties5.Append(bold8);

            paragraphProperties5.Append(justification4);
            paragraphProperties5.Append(paragraphMarkRunProperties5);

            Run run6 = new Run() { RsidRunProperties = "003B602A" };

            RunProperties runProperties6 = new RunProperties();
            Bold bold9 = new Bold();

            runProperties6.Append(bold9);
            Text text6 = new Text { Text = allProblem.ToString() };

            run6.Append(runProperties6);
            run6.Append(text6);

            paragraph6.Append(paragraphProperties5);
            paragraph6.Append(run6);

            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph6);

            tableRow2.Append(tableCell3);
            tableRow2.Append(tableCell4);

            return tableRow2;
        }
        private TableRow AppendStaff(string fio, int countProblem)
        {

            #region RowFio
            TableRow tableRow3 = new TableRow() { RsidTableRowAddition = "003B602A", RsidTableRowProperties = "003B602A" };

            TableCell tableCell5 = new TableCell();

            TableCellProperties tableCellProperties5 = new TableCellProperties();
            TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "4672", Type = TableWidthUnitValues.Dxa };

            tableCellProperties5.Append(tableCellWidth5);

            Paragraph paragraph7 = new Paragraph() { RsidParagraphAddition = "003B602A", RsidRunAdditionDefault = "003B602A" };

            Run run7 = new Run();
            Text text7 = new Text { Text = fio };

            run7.Append(text7);

            paragraph7.Append(run7);

            tableCell5.Append(tableCellProperties5);
            tableCell5.Append(paragraph7);

            TableCell tableCell6 = new TableCell();

            TableCellProperties tableCellProperties6 = new TableCellProperties();
            TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "4673", Type = TableWidthUnitValues.Dxa };

            tableCellProperties6.Append(tableCellWidth6);

            Paragraph paragraph8 = new Paragraph() { RsidParagraphAddition = "003B602A", RsidParagraphProperties = "004641AF", RsidRunAdditionDefault = "003B602A" };

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();
            Justification justification5 = new Justification() { Val = JustificationValues.Center };

            paragraphProperties6.Append(justification5);

            Run run8 = new Run();
            Text text8 = new Text { Text = countProblem.ToString() };

            run8.Append(text8);

            paragraph8.Append(paragraphProperties6);
            paragraph8.Append(run8);

            tableCell6.Append(tableCellProperties6);
            tableCell6.Append(paragraph8);

            tableRow3.Append(tableCell5);
            tableRow3.Append(tableCell6);
            #endregion
            return tableRow3;
        }
    }
}
