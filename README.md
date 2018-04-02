Using `Fluent API` to configure POCO excel behaviors, and then provides IEnumerable&lt;T&gt; has save to and load from excel functionalities.

# 修改内容
- PropertyConfiguration
  - HasConvert 功能, 设置 `字段转换` 器
- IEnumerableNpoiExtensions
  - 执行 Formatter 优先级
  - 执行 CellConfig.Convert
- FluentConfiguration 功能
  - SetIgnore 批量设置忽略字段
  - AutoIndex 自动索引
  - FromAnnotations 从 Annotations 初始化字段
- 示范代码:

```csharp
// 学生模型
public class Student
{
    [Required]
    [Display(Name = "简历 ID")]
    public long ID { get; set; }

    /// <summary>
    /// 性别, 0:未知, 1:男, 2:女
    /// </summary>
    [Required]
    [Display(Name = "性别")]
    public int Gender { get; set; }

    public object GenderExtendData { get; set; }
    public object AttachmentData { get; set; }
}

//配置
var fc = Excel.Setting.For<Student>();
var cGenderMap = new string[]{"未知", "男", "女"};
fc.Property(x => x.Gender).HasConvert(val => cGenderMap[(int)val]));
fc.FromAnnotations()
    .SetIgnore(x => x.ExtendData, x => x.AttachmentData)
    .AutoIndex();
```

# Features
- [x] Decouple the configuration from the POCO model by using `fluent api`.
- [x] Support none configuration POCO, so that if English is your mother language, none any more configurations;

The first features will be very useful for English not their mother language developers.

# IMPORTAMT
1. The repo fork from my [NPOI.Extension](https://github.com/xyting/NPOI.Extension), and remove all the attributes based features, and will only support `Fluent API`.
2. All the issues found in [NPOI.Extension](https://github.com/xyting/NPOI.Extension) will be and only be fixed by [FluentExcel](https://github.com/Arch/FluentExcel), so, please update your codes use `FluentExcel`.

# Overview

![FluentExcel demo](images/demo.PNG)

# Get Started

The following demo codes come from [sample](samples), download and run it for more information.

## Using Package Manager Console to install FluentExcel

        PM> Install-Package FluentExcel
    
## Reference FluentExcel in code

        using FluentExcel;
    
## Use Fluent Api to configure POCO's excel behaviors

We can use `fluent api` to configure the model excel behaviors.

```csharp
using System;
using FluentExcel;

namespace samples
{
    class Program
    {
        static void Main(string[] args)
        {
            // global call this
            FluentConfiguration();

            var len = 20;
            var reports = new Report[len];
            for (int i = 0; i < len; i++)
            {
                reports[i] = new Report
                {
                    City = "ningbo",
                    Building = "世茂首府",
                    HandleTime = DateTime.Now,
                    Broker = "rigofunc 18957139**7",
                    Customer = "yingting 18957139**7",
                    Room = "2#1703",
                    Brokerage = 125 * i,
                    Profits = 25 * i
                };
            }

            var excelFile = @"/Users/rigofunc/Documents/sample.xlsx";

            // save to excel file
            reports.ToExcel(excelFile);

            // load from excel
            var loadFromExcel = Excel.Load<Report>(excelFile);
        }

        /// <summary>
        /// Use fluent configuration api. (doesn't poison your POCO)
        /// </summary>
        static void FluentConfiguration()
        {
            var fc = Excel.Setting.For<Report>();

            fc.HasStatistics("合计", "SUM", 6, 7)
              .HasFilter(firstColumn: 0, lastColumn: 2, firstRow: 0)
              .HasFreeze(columnSplit: 2, rowSplit: 1, leftMostColumn: 2, topMostRow: 1);

            fc.Property(r => r.City)
              .HasExcelIndex(0)
              .HasExcelTitle("城市")
              .IsMergeEnabled();

            // or
            //fc.Property(r => r.City).HasExcelCell(0,"城市", allowMerge: true);

            fc.Property(r => r.Building)
              .HasExcelIndex(1)
              .HasExcelTitle("楼盘")
              .IsMergeEnabled();

            // configures the ignore when exporting or importing.
            fc.Property(r => r.Area)
              .HasExcelIndex(8)
              .HasExcelTitle("Area")
              .IsIgnored(exportingIsIgnored: false, importingIsIgnored: true);

            fc.Property(r => r.HandleTime)
              .HasExcelIndex(2)
              .HasExcelTitle("成交时间")
              .HasDataFormatter("yyyy-MM-dd");

            // or 
            //fc.Property(r => r.HandleTime).HasExcelCell(2, "成交时间", formatter: "yyyy-MM-dd", allowMerge: false);
            // or
            //fc.Property(r => r.HandleTime).HasExcelCell(2, "成交时间", "yyyy-MM-dd");


            fc.Property(r => r.Broker)
              .HasExcelIndex(3)
              .HasExcelTitle("经纪人");

            fc.Property(r => r.Customer)
              .HasExcelIndex(4)
              .HasExcelTitle("客户");

            fc.Property(r => r.Room)
              .HasExcelIndex(5)
              .HasExcelTitle("房源");

            fc.Property(r => r.Brokerage)
              .HasExcelIndex(6)
              // the formatter is Excel formatter, not the C# formatter
              .HasDataFormatter("￥0.00")
              .HasExcelTitle("佣金(元)");

            fc.Property(r => r.Profits)
              .HasExcelIndex(7)
              .HasExcelTitle("收益(元)");
        }
    }
}       
```

## Export POCO to excel & Load IEnumerable&lt;T&gt; from excel.

```csharp
var excelFile = @"/Users/rigofunc/Documents/sample.xlsx";

// save to excel file
reports.ToExcel(excelFile);

// load from excel
var loadFromExcel = Excel.Load<Report>(excelFile);       
```
