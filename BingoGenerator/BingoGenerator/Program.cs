using System.Diagnostics;
using System.Globalization;
using OfficeOpenXml;

Console.WriteLine("Bingo Generator v1.0");
Console.WriteLine();
Console.Write("Enter the number of cards to generate: ");
int cardCount = int.Parse(Console.ReadLine()!, NumberStyles.Integer, CultureInfo.InvariantCulture);

string[] source = File.ReadAllLines("source.txt").Select(s => s.Trim()).Where(s => !string.IsNullOrEmpty(s)).ToHashSet().ToArray();
List<int> sample = Enumerable.Range(0, source.Length).ToList();
HashSet<string> cards = new HashSet<string>();

Console.WriteLine();
Console.Write("Randomising... ");

Random random = new Random();
while (cards.Count < cardCount)
{
  string card;
  do
  {
    card = string.Join('.', sample.OrderBy(i => random.Next()).Take(25).OrderBy(i => i));
  } while (cards.Contains(card));
  cards.Add(card);
}

Console.WriteLine("Done.");

string outputFileName = $"Output{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";
Console.WriteLine();
Console.Write($"Writing {outputFileName}... ");

using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo("Output.xlsx")))
{
  using (ExcelWorksheet sheet = excelPackage.Workbook.Worksheets[0])
  {
    int i;
    for (i = 1; i < cardCount; i++)
    {
      sheet.Cells[1, 1, 5, 5].Copy(sheet.Cells[(i * 5) + 1, 1, (i * 5) + 5, 5]);
    }

    for (i = 0; i < cardCount * 5; i++)
    {
      sheet.Row(i + 1).Height = sheet.Row(1).Height;
    }

    i = 0;
    foreach (string card in cards)
    {
      int[] cardValues = card.Split('.').Select(c => int.Parse(c)).OrderBy(i => random.Next()).ToArray();

      for (int j = 0; j < 25; j++)
      {
        sheet.Cells[(i * 5) + (j % 5) + 1, (j / 5) + 1].Value = source[cardValues[j]];
      }

      i++;
    }

    excelPackage.SaveAs(new FileInfo(outputFileName));
  }
}

Console.WriteLine("Done.");

Console.WriteLine();
Console.Write($"Opening {outputFileName}... ");

ProcessStartInfo startInfo = new ProcessStartInfo();
startInfo.FileName = "EXCEL.EXE";
startInfo.Arguments = outputFileName;
startInfo.UseShellExecute = true;
Process.Start(startInfo);

Console.WriteLine("Done.");

Console.WriteLine();
Console.Write("Press Enter to exit.");
Console.ReadLine();
