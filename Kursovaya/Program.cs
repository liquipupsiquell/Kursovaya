using Microsoft.VisualBasic;
using System.ComponentModel;
using System.Reflection.Metadata;
using Aspose.Words;
using static System.Net.Mime.MediaTypeNames;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

class Program
{
    public static int p = 0;
    private static void Main(string[] args)
    {
        int f = 0;
        Console.Write("Введите размер массива : ");
        f = Convert.ToInt32(Console.ReadLine());
        int [] arr = new int[f];
        int [] arr1 = new int[f];
        string k = "";
        int z1 = 0;
        string[] k1 = { };
        Word(k,k1,arr,arr1,f,z1);
        ShakerSort(arr,z1,f);
        Console.WriteLine("Конечный массив");
        for (int i = 1; i < f + z1; i++)
        {
            Console.WriteLine(arr[i]);         
        }
        Ex(arr, arr1, f,z1);
   
    }   
    public static void Word(string k,string[] k1, int[] arr,int[] arr1, int f,int z1) 
    {   
        Aspose.Words.Document doc = new Aspose.Words.Document("Elements.docx");
        k = Convert.ToString(doc.Range.Text);
        int l1 = 0;  
        k1 =k.Split(new string[] {" "}, StringSplitOptions.RemoveEmptyEntries);
        Console.WriteLine("Начальный массив");
        for (int z = 9; z < f + 9; z++)
        {
            bool isNum = int.TryParse(k1[z], out l1);
            if (isNum)
            {
                try
                {
                    arr[z - 9] = Int32.Parse(k1[z]);
                    arr1[z - 9] = Int32.Parse(k1[z]);
                    Console.WriteLine(arr[z-9]);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
            }
            else 
            {
                z1++;
            }
        }
    }
    public static void ShakerSort(int[] arr,int z1,int f)
    {
        int left = 1;
        int right = f;
        while (left <= right)
        {
            for (int i = left; i < right-1; i++)
            {
                if (arr[i] > arr[i + 1])
                {
                    Swap(arr, i, i + 1);
                    p = p + 1;
                }
            }
            right--;      
            for (int i = right-1; i > left; i--)
            {
                if (arr[i - 1] > arr[i])
                {
                    Swap(arr, i - 1, i);
                    p = p + 1;
                }
            }
            left++;           
        }
        Console.WriteLine("Количество перестановок {0}", p);
    }
    public static void Swap(int[] arr, int i, int j)
    {
        int temp = arr[i];
        arr[i] = arr[j];
        arr[j] = temp;
    }
    public static void Ex(int[] arr,int[] arr1, int f,int z1)
    {
        int i = 1;
        var path = Path.Combine(Environment.CurrentDirectory, "Export", "Elem.xlsx");
        var wb = new XLWorkbook();
        var sh = wb.Worksheets.Add("Elements");
        sh.Cell(1, 1).SetValue("Начальный массив");
        sh.Cell(1, 2).SetValue("Конечный массив");
        sh.Cell(1, 3).SetValue("Кол-во перестановок");
        sh.Cell(2, 3).SetValue(p);
        for (i=1;i<f+z1;i++)
        {
            sh.Cell(i + 1, 1).SetValue(arr1[i]);
            sh.Cell(i + 1, 2).SetValue(arr[i]);
        }
        wb.SaveAs(path);
    }
}
