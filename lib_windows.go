//go:build windows

// Copyright 2025 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package excelize

import (
	"io"
)

// readAll is like io.ReadAll, but uses mmap if possible.
func readAll(r io.Reader) ([]byte, error) {
	return io.ReadAll(r)
}
