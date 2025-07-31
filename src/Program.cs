using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Excel ファイル読み書きサンプル - DocumentFormat.OpenXml使用");
            Console.WriteLine("==========================================================");
            
            // デバッグ用のブレークポイント設置箇所
            string excelFilePath = Path.Combine(Directory.GetCurrentDirectory(), "sample.xlsx");
            
            try
            {
                // ステップ1: サンプルデータの作成
                Console.WriteLine("Step 1: サンプルデータを作成します...");
                var sampleData = CreateSampleData();
                
                // ステップ2: Excelファイルの書き込み
                Console.WriteLine("Step 2: Excelファイルに書き込みます...");
                ExcelHandler.WriteExcel(excelFilePath, sampleData, "社員リスト");
                
                // ステップ3: ファイルの存在確認
                Console.WriteLine("Step 3: ファイルの存在を確認します...");
                bool fileExists = File.Exists(excelFilePath);
                Console.WriteLine($"ファイル存在: {fileExists}");
                
                if (fileExists)
                {
                    // ステップ4: ファイル情報の表示
                    var fileInfo = new FileInfo(excelFilePath);
                    Console.WriteLine($"ファイルサイズ: {fileInfo.Length} bytes");
                    Console.WriteLine($"作成日時: {fileInfo.CreationTime}");
                    
                    // ステップ5: ファイルの読み込み
                    Console.WriteLine("Step 4: Excelファイルを読み込みます...");
                    var readData = ExcelHandler.ReadExcel(excelFilePath);
                    
                    // ステップ6: データの表示
                    Console.WriteLine("Step 5: データを表示します...");
                    DisplayData(readData);
                    
                    // ステップ7: データの処理
                    Console.WriteLine("Step 6: データを処理します...");
                    ProcessData(readData);
                    
                    // ステップ8: 複数シートのサンプル
                    Console.WriteLine("Step 7: 複数シートのファイルを作成します...");
                    CreateMultipleSheetSample();
                }
            }
            catch (Exception ex)
            {
                // エラー処理でもブレークポイントを設置可能
                Console.WriteLine($"エラーが発生しました: {ex.Message}");
                Console.WriteLine($"スタックトレース: {ex.StackTrace}");
            }
            
            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
        }
        
        /// <summary>
        /// サンプルデータを作成
        /// </summary>
        /// <returns>サンプルデータ</returns>
        static List<string[]> CreateSampleData()
        {
            var data = new List<string[]>();
            
            // ヘッダー行
            string[] headers = { "社員ID", "名前", "年齢", "部署", "給与", "入社日" };
            data.Add(headers);
            
            // データ行（ループでブレークポイント設置可能）
            string[][] employees = {
                new string[] { "EMP001", "田中太郎", "30", "開発部", "500000", "2020-04-01" },
                new string[] { "EMP002", "佐藤花子", "25", "デザイン部", "400000", "2021-07-15" },
                new string[] { "EMP003", "鈴木一郎", "35", "営業部", "600000", "2019-01-10" },
                new string[] { "EMP004", "高橋美咲", "28", "人事部", "450000", "2022-03-01" },
                new string[] { "EMP005", "山田和夫", "42", "管理部", "700000", "2018-08-20" }
            };
            
            foreach (var employee in employees)
            {
                // ここにブレークポイントを設置して各行の処理を確認
                data.Add(employee);
                Console.WriteLine($"従業員データを追加: {employee[1]} ({employee[3]})");
            }
            
            return data;
        }
        
        /// <summary>
        /// データを表示
        /// </summary>
        /// <param name="data">表示するデータ</param>
        static void DisplayData(List<string[]> data)
        {
            Console.WriteLine("\n=== データ表示 ===");
            
            for (int i = 0; i < data.Count; i++)
            {
                var row = data[i];
                Console.Write($"Row {i + 1}: ");
                
                for (int j = 0; j < row.Length; j++)
                {
                    // 各セルの値を確認できる
                    string cellValue = row[j];
                    Console.Write($"[{cellValue}] ");
                }
                Console.WriteLine();
            }
        }
        
        /// <summary>
        /// データを処理（給与統計の計算）
        /// </summary>
        /// <param name="data">処理するデータ</param>
        static void ProcessData(List<string[]> data)
        {
            if (data.Count <= 1) return; // ヘッダーのみの場合
            
            Console.WriteLine("\n=== データ処理（給与統計） ===");
            
            // 給与の合計を計算（デバッグで変数の値を確認）
            long totalSalary = 0;
            int employeeCount = 0;
            long maxSalary = 0;
            long minSalary = long.MaxValue;
            string highestPaidEmployee = "";
            
            for (int i = 1; i < data.Count; i++) // ヘッダーをスキップ
            {
                var row = data[i];
                if (row.Length > 4) // 給与列が存在するか確認
                {
                    if (long.TryParse(row[4], out long salary))
                    {
                        totalSalary += salary;
                        employeeCount++;
                        
                        // 最高給与と最低給与をチェック
                        if (salary > maxSalary)
                        {
                            maxSalary = salary;
                            highestPaidEmployee = row[1];
                        }
                        if (salary < minSalary)
                        {
                            minSalary = salary;
                        }
                        
                        // ここでsalary変数の値を確認できる
                        Console.WriteLine($"{row[1]}({row[3]}): {salary:N0}円");
                    }
                }
            }
            
            if (employeeCount > 0)
            {
                double averageSalary = (double)totalSalary / employeeCount;
                Console.WriteLine($"\n【統計結果】");
                Console.WriteLine($"従業員数: {employeeCount}人");
                Console.WriteLine($"給与合計: {totalSalary:N0}円");
                Console.WriteLine($"平均給与: {averageSalary:N0}円");
                Console.WriteLine($"最高給与: {maxSalary:N0}円 ({highestPaidEmployee})");
                Console.WriteLine($"最低給与: {minSalary:N0}円");
            }
        }
        
        /// <summary>
        /// 複数シートのサンプルファイルを作成
        /// </summary>
        static void CreateMultipleSheetSample()
        {
            var multiSheetData = new Dictionary<string, List<string[]>>();
            
            // 部署別売上データシート
            multiSheetData["部署別売上"] = new List<string[]>
            {
                new string[] { "部署", "Q1売上", "Q2売上", "Q3売上", "Q4売上", "年間合計" },
                new string[] { "開発部", "1200", "1350", "1100", "1450", "5100" },
                new string[] { "営業部", "2200", "2100", "2300", "2400", "9000" },
                new string[] { "デザイン部", "800", "900", "850", "950", "3500" },
                new string[] { "人事部", "400", "420", "380", "440", "1640" }
            };
            
            // 月別経費データシート
            multiSheetData["月別経費"] = new List<string[]>
            {
                new string[] { "月", "人件費", "オフィス費", "システム費", "その他", "合計" },
                new string[] { "1月", "3000", "500", "200", "300", "4000" },
                new string[] { "2月", "3100", "500", "250", "280", "4130" },
                new string[] { "3月", "3050", "520", "200", "350", "4120" }
            };
            
            string multiSheetFile = Path.Combine(Directory.GetCurrentDirectory(), "multi_sheet_sample.xlsx");
            ExcelHandler.WriteMultipleSheets(multiSheetFile, multiSheetData);
        }
    }
}
