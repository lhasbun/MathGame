using ClosedXML.Excel;
using System.Data;

var activeSession = true;

Console.WriteLine("Please enter your username:");

var name = Console.ReadLine();
var date = DateTime.Now;


while (activeSession)
{
    activeSession = Menu(activeSession);
}

Console.WriteLine("Thank you for playing!");
Console.WriteLine("Press any key to exit.");
Console.ReadKey();

void Calculus()
{
    DataSet ds = FillDataset("calculus");
    var table = ds.Tables[0];
    var rand = new Random();
    QandA(table, rand);
}

bool Menu(bool activeSession)
{
    

    Console.WriteLine("--------------------------------------------");
    Console.WriteLine("""
              ____                 
             /___/\_                               
            _\   \/_/\__                             
          __\       \/_/\                       
          \   __    __ \ \          QUICK MATHS 1.0:           
         __\  \_\   \_\ \ \   __            CALCULUS    
        /_/\\   __   __  \ \_/_/\           
        \_\/_\__\/\__\/\__\/_\_\/          
           \_\/_/\       /_\_\/              
              \_\/       \_\/     
       
        """);
    Console.WriteLine($"Hello {name}. Your current score is {name}");
    Console.WriteLine($@"Select a mathematical domain:
    (C)alculus    
    (Q)uit");
    Console.WriteLine("--------------------------------------------");

    var selection = Console.ReadLine();
    var selectMade = false;

    while (selectMade == false)
    {
        if (selection == null)
        {
            Console.WriteLine("Please select a valid option.");
            break;
        }
        else if (selection.Trim().ToLower() == "c")
        {
            Calculus();
        }
        else if (selection.Trim().ToLower() == "q")
        {
            selectMade = true;
            activeSession = false;
            Console.WriteLine("Goodbye!");
            Console.Clear();
        }
        else
        {
            Console.WriteLine("Invalid selection. Please try again.");
            break;
        }
    }

    return activeSession;
}

DataSet FillDataset(string domain) 
{

    var filePath = @"C:\Users\lehas\OneDrive\Escritorio\GitHub Repos\MathGame\MathGame\questions_answers.xlsx";
    var sheetNum = 0;

    if (domain.Trim().ToLower() == "calculus")
    {
        sheetNum = 1;
    }

    // Open the Excel file using ClosedXML.
    // Keep in mind the Excel file cannot be open when trying to read it
    using (XLWorkbook workBook = new XLWorkbook(filePath))
    {
        //Read the first Sheet from Excel file.
        IXLWorksheet workSheet = workBook.Worksheet(sheetNum);

        //Create a new DataTable.
        DataTable dt = new DataTable();

        //Loop through the Worksheet rows.
        bool firstRow = true;
        foreach (IXLRow row in workSheet.Rows())
        {
            //Use the first row to add columns to DataTable.
            if (firstRow)
            {
                foreach (IXLCell cell in row.Cells())
                {
                    dt.Columns.Add(cell.Value.ToString());
                }
                firstRow = false;
            }
            else
            {
                //Add rows to DataTable.
                dt.Rows.Add();
                int i = 0;

                foreach (IXLCell cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
                {
                    dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                    i++;
                }
            }
        }

        DataSet ds = new DataSet();
        ds.Tables.Add(dt);
        return ds;
    }
}

static void QandA(DataTable table, Random rand)
{
    for (int i = 0; i < 10; i++)
    {
        // Get a random row from the DataTable
        var randomRow = rand.Next(1, table.Rows.Count);

        Console.WriteLine($"Question {i + 1}: {table.Rows[randomRow][0]}");
        Console.WriteLine($"1) {table.Rows[randomRow][1]}");
        Console.WriteLine($"2) {table.Rows[randomRow][2]}");
        Console.WriteLine($"3) {table.Rows[randomRow][3]}");
        Console.WriteLine($"4) {table.Rows[randomRow][4]}");
        Console.WriteLine($"5) {table.Rows[randomRow][5]}");
        Console.WriteLine("Please select an answer (1-5) or 'q' to quit:");

        var userAnswer = Console.ReadLine();
        userAnswer = $"A{userAnswer}";
        var answerSelection = table.Rows[randomRow][userAnswer.ToString()].ToString();
        var correctAnswer = table.Rows[randomRow][6].ToString();

        if (userAnswer == null || userAnswer.Trim().ToLower() == "Aq")
        {
            Console.WriteLine("Exiting the quiz.");
            break;
        }
        else
        {
            if (userAnswer == correctAnswer)
            {
                Console.WriteLine("Correct!");
            }
            else
            {
                Console.WriteLine($"Incorrect. The correct answer is: {correctAnswer}");
            }
        }
    }
}