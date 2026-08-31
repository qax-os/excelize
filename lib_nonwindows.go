//go:build !windows

// Copyright 2025 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package excelize

import (
	"io"
	"os"
	"runtime"
	"syscall"
)

// readAll is like io.ReadAll, but uses mmap if possible.
func readAll(r io.Reader) ([]byte, error) {
	if fder, ok := r.(interface {
		Fd() uintptr
		Stat() (os.FileInfo, error)
	}); ok {
		if fi, err := fder.Stat(); err == nil {
			if b, err := syscall.Mmap(
				int(fder.Fd()),
				0, int(fi.Size()),
				syscall.PROT_READ,
				syscall.MAP_PRIVATE|syscall.MAP_POPULATE,
			); err == nil {
				runtime.SetFinalizer(&b, func(_ any) error {
					return syscall.Munmap(b)
				})
				return b, nil
			}
		}
	}
	return io.ReadAll(r)
}
