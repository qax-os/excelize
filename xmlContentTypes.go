// Copyright 2016 - 2025 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.
//
// Package excelize providing a set of functions that allow you to write to and
// read from XLAM / XLSM / XLSX / XLTM / XLTX files. Supports reading and
// writing spreadsheet documents generated by Microsoft Excel™ 2007 and later.
// Supports complex components by high compatibility, and provided streaming
// API for generating or reading data from a worksheet with huge amounts of
// data. This library needs Go version 1.23 or later.

package excelize

import (
	"encoding/xml"
	"sync"
)

// xlsxTypes directly maps the types' element of content types for relationship
// parts, it takes a Multipurpose Internet Mail Extension (MIME) media type as a
// value.
type xlsxTypes struct {
	mu        sync.Mutex
	XMLName   xml.Name       `xml:"http://schemas.openxmlformats.org/package/2006/content-types Types"`
	Defaults  []xlsxDefault  `xml:"Default"`
	Overrides []xlsxOverride `xml:"Override"`
}

// xlsxOverride directly maps the override element in the namespace
// http://schemas.openxmlformats.org/package/2006/content-types
type xlsxOverride struct {
	PartName    string `xml:",attr"`
	ContentType string `xml:",attr"`
}

// xlsxDefault directly maps the default element in the namespace
// http://schemas.openxmlformats.org/package/2006/content-types
type xlsxDefault struct {
	Extension   string `xml:",attr"`
	ContentType string `xml:",attr"`
}
