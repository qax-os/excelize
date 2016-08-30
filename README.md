![Excelize](./excelize.png "Excelize")

# Excelize

[![Build Status](https://travis-ci.org/Luxurioust/excelize.svg?branch=master)](https://travis-ci.org/Luxurioust/excelize)
[![Code Coverage](https://codecov.io/gh/Luxurioust/excelize/branch/master/graph/badge.svg)](https://codecov.io/gh/Luxurioust/excelize)
[![GoDoc](https://godoc.org/github.com/Luxurioust/excelize?status.svg)](https://godoc.org/github.com/Luxurioust/excelize)
[![Licenses](https://img.shields.io/badge/license-bsd-orange.svg)](https://opensource.org/licenses/BSD-3-Clause)
[![Join the chat at https://gitter.im/xuri-excelize/Lobby](https://img.shields.io/badge/GITTER-join%20chat-green.svg)](https://gitter.im/xuri-excelize/Lobby)

## Introduction

Excelize is a library written in pure Golang and providing a set of function that allow you to write to and read from XLSX files. Support read and write XLSX file geterated by Office Excel 2007 and later. The full API docs can be viewed using go’s built in documentation tool, or online at [godoc.org](https://godoc.org/github.com/Luxurioust/excelize).

## Basic Usage

### Installation

Golang version requirements 1.6.0 or higher.

```
go get github.com/Luxurioust/excelize
```

### Create XLSX files

Here is a minimal example usage that will create XLSX file.

```
package main

import (
    "fmt"
    "github.com/Luxurioust/excelize"
)

func main() {
    xlsx := excelize.CreateFile()
    xlsx = excelize.NewSheet(xlsx, 2, "Sheet2")
    xlsx = excelize.NewSheet(xlsx, 3, "Sheet3")
    xlsx = excelize.SetCellInt(xlsx, "Sheet2", "A23", 10)
    xlsx = excelize.SetCellStr(xlsx, "Sheet3", "B20", "Hello")
    err := excelize.Save(xlsx, "~/Workbook.xlsx")
    if err != nil {
        fmt.Println(err)
    }
}
```

### Writing XLSX files

The following constitutes the bare minimum required to write an XLSX document.

```
package main

import (
    "fmt"
    "github.com/Luxurioust/excelize"
)

func main() {
    xlsx, err := excelize.OpenFile("~/Workbook.xlsx")
    if err != nil {
        fmt.Println(err)
    }
    xlsx = excelize.SetCellInt(xlsx, "Sheet2", "B2", 100)
    xlsx = excelize.SetCellStr(xlsx, "Sheet2", "C11", "Hello")
    xlsx = excelize.NewSheet(xlsx, 3, "TestSheet")
    xlsx = excelize.SetCellInt(xlsx, "Sheet3", "A23", 10)
    xlsx = excelize.SetCellStr(xlsx, "Sheet3", "b230", "World")
    xlsx = excelize.SetActiveSheet(xlsx, 2)
    if err != nil {
        fmt.Println(err)
    }
    err = excelize.Save(xlsx, "~/Workbook.xlsx")
}
```

## Contributing

Contributions are welcome! Open a pull request to fix a bug, or open an issue to discuss a new feature or change.

## Licenses

This program is under the terms of the BSD 3-Clause License. See [https://opensource.org/licenses/BSD-3-Clause](https://opensource.org/licenses/BSD-3-Clause).