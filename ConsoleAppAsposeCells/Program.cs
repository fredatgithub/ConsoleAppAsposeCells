// See https://aka.ms/new-console-template for more information
using Aspose.Cells;
using System.Drawing;

Action<string> display = Console.WriteLine;
display("ASPOSE.Cells sample app");
Workbook workbook = new Workbook();
if (workbook.Worksheets.Count == 0)
{
  workbook.Worksheets.Add("sheet1");
}
else
{
  workbook.Worksheets[0].Name = "sheet1";
}

// Create a Worksheet and get the first sheet.
Worksheet worksheet = workbook.Worksheets[0];

// Create a Cells object ot fetch all the cells.
Cells cells = worksheet.Cells;

// Unmerge the cells.
cells.UnMerge(5, 2, 2, 3);

//Merge some Cells (C6:E7) into a single C6 Cell.
cells.Merge(5, 2, 2, 3);

//Input data into C6 Cell.
worksheet.Cells[5, 2].PutValue("cells merged");

//Create a Style object to fetch the Style of C6 Cell.
Style style = worksheet.Cells[5, 2].GetStyle();

//Create a Font object
Aspose.Cells.Font font = style.Font;

//Set the name.
font.Name = "Times New Roman";

//Set the font size.
font.Size = 18;

//Set the font color
font.Color = Color.Blue;

//Bold the text
font.IsBold = true;

//Make it italic
font.IsItalic = true;

//Set the backgrond color of C6 Cell to Red
style.ForegroundColor = Color.Red;

style.Pattern = BackgroundType.Solid;

//Apply the Style to C6 Cell.
cells[5, 2].SetStyle(style);

worksheet.Cells[0, 0].PutValue("Header1");
worksheet.Cells[0, 1].PutValue("Header2 long name");
worksheet.Cells[0, 2].PutValue("Header3 short name");
worksheet.Cells[0, 3].PutValue("Header4");
worksheet.Cells[0, 4].PutValue("Header5");
worksheet.Cells[0, 5].PutValue("Header6");
worksheet.Cells[0, 6].PutValue("Header7");
worksheet.Cells[0, 7].PutValue("Header8");
worksheet.Cells[0, 8].PutValue("Header9");
worksheet.Cells[0, 9].PutValue("Header10");
worksheet.Cells[0, 10].PutValue("Header11");

worksheet.Cells[1, 0].PutValue("Header1");
worksheet.Cells[1, 1].PutValue("Header2 long name");
worksheet.Cells[1, 2].PutValue("Header3 short name");
worksheet.Cells[1, 3].PutValue("Header4");
worksheet.Cells[1, 4].PutValue("Header5");
worksheet.Cells[1, 5].PutValue("Header6");
worksheet.Cells[1, 6].PutValue("Header7");
worksheet.Cells[1, 7].PutValue("Header8");
worksheet.Cells[1, 8].PutValue("Header9");
worksheet.Cells[1, 9].PutValue("Header10");
worksheet.Cells[1, 10].PutValue("Header11");

worksheet.Cells[2, 0].PutValue("Header1");
worksheet.Cells[2, 1].PutValue("Header2 long name");
worksheet.Cells[2, 2].PutValue("Header3 short name");
worksheet.Cells[2, 3].PutValue("Header4");
worksheet.Cells[2, 4].PutValue("Header5");
worksheet.Cells[2, 5].PutValue("Header6");
worksheet.Cells[2, 6].PutValue("Header7");
worksheet.Cells[2, 7].PutValue("Header8");
worksheet.Cells[2, 8].PutValue("Header9");
worksheet.Cells[2, 9].PutValue("Header10");
worksheet.Cells[2, 10].PutValue("Header11");

worksheet.Cells[3, 0].PutValue("Header1");
worksheet.Cells[3, 1].PutValue("Header2 long name");
worksheet.Cells[3, 2].PutValue("Header3 short name");
worksheet.Cells[3, 3].PutValue("Header4");
worksheet.Cells[3, 4].PutValue("Header5");
worksheet.Cells[3, 5].PutValue("Header6");
worksheet.Cells[3, 6].PutValue("Header7");
worksheet.Cells[3, 7].PutValue("Header8");
worksheet.Cells[3, 8].PutValue("Header9");
worksheet.Cells[3, 9].PutValue("Header10");
worksheet.Cells[3, 10].PutValue("Header11");

Style headerStyle = worksheet.Cells[0, 0].GetStyle();
Aspose.Cells.Font headerFont = style.Font;
headerFont.IsBold = true;
headerFont.IsItalic = false;
headerStyle.ForegroundColor = Color.Blue;
//headerStyle.BackgroundColor = Color.LightBlue;
headerStyle.Pattern = BackgroundType.Solid;
cells[0, 0].SetStyle(headerStyle);

SetColor(0, 1, Color.BlueViolet);
SetColor(0, 2, Color.AliceBlue);
SetColor(0, 3, Color.CadetBlue);
SetColor(0, 4, Color.CornflowerBlue);
SetColor(0, 5, Color.DarkBlue);
SetColor(0, 6, Color.DarkSlateBlue);
SetColor(0, 7, Color.DeepSkyBlue);
SetColor(0, 8, Color.BlueViolet);
SetColor(0, 9, Color.AliceBlue);
SetColor(0, 10, Color.CadetBlue);
SetColor(1, 0, Color.CornflowerBlue);
SetColor(1, 1, Color.DarkBlue);
SetColor(1, 2, Color.DarkSlateBlue);
SetColor(1, 3, Color.DeepSkyBlue);
SetColor(1, 4, Color.DodgerBlue);
SetColor(1, 5, Color.LightBlue);
SetColor(1, 6, Color.LightSkyBlue);
SetColor(1, 7, Color.LightSteelBlue);
SetColor(1, 8, Color.MediumBlue);
SetColor(1, 9, Color.MediumSlateBlue);
SetColor(1, 10, Color.MidnightBlue);
SetColor(2, 0, Color.PowderBlue);
SetColor(2, 1, Color.RoyalBlue);
SetColor(2, 2, Color.SkyBlue);
SetColor(2, 3, Color.SlateBlue);
SetColor(2, 4, Color.SteelBlue);
SetColor(3, 0, Color.DeepSkyBlue);
for (int i = 0; i < 12; i++)
{
  SetColor(3, i, Color.DeepSkyBlue);
  SetBold(3, i);
}
worksheet.AutoFitColumns();

worksheet.IsSelected = true;
workbook.Save("test.xlsx");
display("Press any key to exit:");
//Console.ReadKey();

void SetColor(int rowNumber, int columnNumber, Color color)
{
  headerStyle.ForegroundColor = color;
  cells[rowNumber, columnNumber].SetStyle(headerStyle);
}

void SetBold(int rowNumber, int columnNumber)
{
  headerFont.IsBold = true;
  cells[rowNumber, columnNumber].SetStyle(headerStyle);
}

void UnSetBold(int rowNumber, int columnNumber)
{
  headerFont.IsBold = false;
  cells[rowNumber, columnNumber].SetStyle(headerStyle);
}
