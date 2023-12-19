<p align="center"><img width="650" src="./excelize.svg" alt="Excelize logo"></p>

<p align="center">
    <a href="https://github.com/xuri/excelize/actions/workflows/go.yml"><img src="https://github.com/xuri/excelize/actions/workflows/go.yml/badge.svg" alt="Build Status"></a>
    <a href="https://codecov.io/gh/qax-os/excelize"><img src="https://codecov.io/gh/qax-os/excelize/branch/master/graph/badge.svg" alt="Code Coverage"></a>
    <a href="https://goreportcard.com/report/github.com/xuri/excelize/v2"><img src="https://goreportcard.com/badge/github.com/xuri/excelize/v2" alt="Go Report Card"></a>
    <a href="https://pkg.go.dev/github.com/xuri/excelize/v2"><img src="https://img.shields.io/badge/go.dev-reference-007d9c?logo=go&logoColor=white" alt="go.dev"></a>
    <a href="https://opensource.org/licenses/BSD-3-Clause"><img src="https://img.shields.io/badge/license-bsd-orange.svg" alt="Licenses"></a>
    <a href="https://www.paypal.com/paypalme/xuri"><img src="https://img.shields.io/badge/Donate-PayPal-green.svg" alt="Donate"></a>
</p>

# Excelize

## 简介

Excelize 是 Go 语言编写的用于操作 Office Excel 文档基础库，基于 ECMA-376，ISO/IEC 29500 国际标准。可以使用它来读取、写入由 Microsoft Excel&trade; 2007 及以上版本创建的电子表格文档。支持 XLAM / XLSM / XLSX / XLTM / XLTX 等多种文档格式，高度兼容带有样式、图片(表)、透视表、切片器等复杂组件的文档，并提供流式读写函数，用于处理包含大规模数据的工作簿。可应用于各类报表平台、云计算、边缘计算等系统。使用本类库要求使用的 Go 语言为 1.18 或更高版本，请注意，Go 1.21.0 中存在[不兼容的更改](https://github.com/golang/go/issues/61881)，导致 Excelize 基础库无法在该版本上正常工作，如果您使用的是 Go 1.21.x，请升级到 Go 1.21.1 及更高版本。完整的使用文档请访问 [go.dev](https://pkg.go.dev/github.com/xuri/excelize/v2) 或查看 [参考文档](https://xuri.me/excelize/)。

## 快速上手

### 安装

```bash
go get github.com/xuri/excelize
```

- 如果您使用 [Go Modules](https://go.dev/blog/using-go-modules) 管理软件包，请使用下面的命令来安装最新版本。

```bash
go get github.com/xuri/excelize/v2
```

### 创建 Excel 文档

下面是一个创建 Excel 文档的简单例子：

```go
package main

import (
    "fmt"

    "github.com/xuri/excelize/v2"
)

func main() {
    f := excelize.NewFile()
    defer func() {
        if err := f.Close(); err != nil {
            fmt.Println(err)
        }
    }()
    // 创建一个工作表
    index, err := f.NewSheet("Sheet2")
    if err != nil {
        fmt.Println(err)
        return
    }
    // 设置单元格的值
    f.SetCellValue("Sheet2", "A2", "Hello world.")
    f.SetCellValue("Sheet1", "B2", 100)
    // 设置工作簿的默认工作表
    f.SetActiveSheet(index)
    // 根据指定路径保存文件
    if err := f.SaveAs("Book1.xlsx"); err != nil {
        fmt.Println(err)
    }
}
```

### 读取 Excel 文档

下面是读取 Excel 文档的例子：

```go
package main

import (
    "fmt"

    "github.com/xuri/excelize/v2"
)

func main() {
    f, err := excelize.OpenFile("Book1.xlsx")
    if err != nil {
        fmt.Println(err)
        return
    }
    defer func() {
        // 关闭工作簿
        if err := f.Close(); err != nil {
            fmt.Println(err)
        }
    }()
    // 获取工作表中指定单元格的值
    cell, err := f.GetCellValue("Sheet1", "B2")
    if err != nil {
        fmt.Println(err)
        return
    }
    fmt.Println(cell)
    // 获取 Sheet1 上所有单元格
    rows, err := f.GetRows("Sheet1")
    if err != nil {
        fmt.Println(err)
        return
    }
    for _, row := range rows {
        for _, colCell := range row {
            fmt.Print(colCell, "\t")
        }
        fmt.Println()
    }
}
```

### 在 Excel 文档中创建图表

使用 Excelize 生成图表十分简单，仅需几行代码。您可以根据工作表中的已有数据构建图表，或向工作表中添加数据并创建图表。

<p align="center"><img width="650" src="./test/images/chart.png" alt="使用 Excelize 在 Excel 电子表格文档中创建图表"></p>

```go
package main

import (
    "fmt"

    "github.com/xuri/excelize/v2"
)

func main() {
    f := excelize.NewFile()
    defer func() {
        if err := f.Close(); err != nil {
            fmt.Println(err)
        }
    }()
    for idx, row := range [][]interface{}{
        {nil, "Apple", "Orange", "Pear"}, {"Small", 2, 3, 3},
        {"Normal", 5, 2, 4}, {"Large", 6, 7, 8},
    } {
        cell, err := excelize.CoordinatesToCellName(1, idx+1)
        if err != nil {
            fmt.Println(err)
            return
        }
        f.SetSheetRow("Sheet1", cell, &row)
    }
    if err := f.AddChart("Sheet1", "E1", &excelize.Chart{
        Type: excelize.Col3DClustered,
        Series: []excelize.ChartSeries{
            {
                Name:       "Sheet1!$A$2",
                Categories: "Sheet1!$B$1:$D$1",
                Values:     "Sheet1!$B$2:$D$2",
            },
            {
                Name:       "Sheet1!$A$3",
                Categories: "Sheet1!$B$1:$D$1",
                Values:     "Sheet1!$B$3:$D$3",
            },
            {
                Name:       "Sheet1!$A$4",
                Categories: "Sheet1!$B$1:$D$1",
                Values:     "Sheet1!$B$4:$D$4",
            }},
        Title: []excelize.RichTextRun{
            {
                Text: "Fruit 3D Clustered Column Chart",
            },
        },
    }); err != nil {
        fmt.Println(err)
        return
    }
    // 根据指定路径保存文件
    if err := f.SaveAs("Book1.xlsx"); err != nil {
        fmt.Println(err)
    }
}
```

### 向 Excel 文档中插入图片

```go
package main

import (
    "fmt"
    _ "image/gif"
    _ "image/jpeg"
    _ "image/png"

    "github.com/xuri/excelize/v2"
)

func main() {
    f, err := excelize.OpenFile("Book1.xlsx")
    if err != nil {
        fmt.Println(err)
        return
    }
    defer func() {
        // 关闭工作簿
        if err := f.Close(); err != nil {
            fmt.Println(err)
        }
    }()
    // 插入图片
    if err := f.AddPicture("Sheet1", "A2", "image.png", nil); err != nil {
        fmt.Println(err)
    }
    // 在工作表中插入图片，并设置图片的缩放比例
    if err := f.AddPicture("Sheet1", "D2", "image.jpg",
        &excelize.GraphicOptions{ScaleX: 0.5, ScaleY: 0.5}); err != nil {
        fmt.Println(err)
    }
    // 在工作表中插入图片，并设置图片的打印属性
    enable, disable := true, false
    if err := f.AddPicture("Sheet1", "H2", "image.gif",
        &excelize.GraphicOptions{
            PrintObject:     &enable,
            LockAspectRatio: false,
            OffsetX:         15,
            OffsetY:         10,
            Locked:          &disable,
        }); err != nil {
        fmt.Println(err)
    }
    // 保存工作簿
    if err = f.Save(); err != nil {
        fmt.Println(err)
    }
}
```

## 社区合作

欢迎您为此项目贡献代码，提出建议或问题、修复 Bug 以及参与讨论对新功能的想法。 XML 符合标准： [part 1 of the 5th edition of the ECMA-376 Standard for Office Open XML](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)。

## 开源许可

本项目遵循 BSD 3-Clause 开源许可协议，访问 [https://opensource.org/licenses/BSD-3-Clause](https://opensource.org/licenses/BSD-3-Clause) 查看许可协议文件。

Excel 徽标是 [Microsoft Corporation](https://aka.ms/trademarks-usage) 的商标，项目的图片是一种改编。

gopher.{ai,svg,png} 由 [Takuya Ueda](https://twitter.com/tenntenn) 创作，遵循 [Creative Commons 3.0 Attributions license](http://creativecommons.org/licenses/by/3.0/) 创作共用授权条款。
