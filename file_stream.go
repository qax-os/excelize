// Copyright 2016 - 2026 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package excelize

import (
	"bytes"
	"encoding/binary"
	"io"
	"math"
	"os"
)

// workbookCountWriter tracks bytes accepted by the destination, including
// partial writes, for the io.WriterTo contract.
type workbookCountWriter struct {
	writer  io.Writer
	written int64
}

// Write forwards output and detects invalid short writes.
func (w *workbookCountWriter) Write(p []byte) (int, error) {
	n, err := w.writer.Write(p)
	w.written += int64(n)
	if n < len(p) && err == nil {
		err = io.ErrShortWrite
	}
	return n, err
}

// zip64SpoolNeeded checks serialized entry sizes before emitting any ZIP bytes.
// Entries of at least 2 GiB use a seekable temporary archive conservatively:
// compression can expand incompressible data past ZIP64's 4 GiB boundary.
// Ordinary exports need no additional temporary file, including on WebAssembly.
func (f *File) zip64SpoolNeeded() (bool, error) {
	const threshold = int64(math.MaxUint32 / 2)
	for _, stream := range f.streams {
		size := int64(stream.rawData.buf.Len())
		if stream.rawData.tmp != nil {
			info, err := stream.rawData.tmp.Stat()
			if err != nil {
				return false, err
			}
			size += info.Size()
		}
		if size >= threshold {
			return true, nil
		}
	}
	large := false
	f.Pkg.Range(func(_, value interface{}) bool {
		data, _ := value.([]byte)
		large = int64(len(data)) >= threshold
		return !large
	})
	if large {
		return true, nil
	}
	f.tempFiles.Range(func(name, value interface{}) bool {
		if _, exists := f.Pkg.Load(name); exists {
			return true
		}
		info, err := os.Stat(value.(string))
		if err == nil {
			large = info.Size() >= threshold
		}
		return !large
	})
	return large, nil
}

// writeToDirect packages ordinary workbooks directly into the caller's writer.
func (f *File) writeToDirect(output io.Writer) (int64, error) {
	counter := &workbookCountWriter{writer: output}
	archive := f.ZipWriter(counter)
	if err := f.writeToZipPrepared(archive); err != nil {
		_ = archive.Close()
		return counter.written, err
	}
	err := archive.Close()
	return counter.written, err
}

// writeToSpooled retains ZIP64 local-header compatibility without holding the
// complete compressed archive in memory. Only potentially large entries take
// this path; encrypted output retains the existing WriteToBuffer path.
func (f *File) writeToSpooled(output io.Writer) (int64, error) {
	directory := ""
	if f.options != nil {
		directory = f.options.TmpDir
	}
	file, err := os.CreateTemp(directory, "excelize-save-*")
	if err != nil {
		return 0, err
	}
	defer func() { _ = file.Close(); _ = os.Remove(file.Name()) }()
	if _, err = f.writeToDirect(file); err != nil {
		return 0, err
	}
	if err = f.patchZip64LFH(file, file); err != nil {
		return 0, err
	}
	if _, err = file.Seek(0, io.SeekStart); err != nil {
		return 0, err
	}
	return io.Copy(output, file)
}

// patchZip64LFH applies the same version correction as writeZip64LFH using
// bounded reads and random-access writes instead of a complete bytes.Buffer.
func (f *File) patchZip64LFH(input io.ReaderAt, output io.WriterAt) error {
	if len(f.zip64Entries) == 0 {
		return nil
	}
	signature := []byte{0x50, 0x4b, 0x03, 0x04}
	buffer := make([]byte, 64<<10)
	header := make([]byte, 30)
	name := make([]byte, math.MaxUint16)
	var offset int64
	for {
		n, readErr := input.ReadAt(buffer, offset)
		if readErr != nil && readErr != io.EOF {
			return readErr
		}
		cursor := 0
		for cursor < n {
			index := bytes.Index(buffer[cursor:n], signature)
			if index < 0 {
				break
			}
			index += cursor
			position := offset + int64(index)
			if _, err := input.ReadAt(header, position); err != nil {
				if err == io.EOF {
					return nil
				}
				return err
			}
			length := int(binary.LittleEndian.Uint16(header[26:28]))
			if _, err := input.ReadAt(name[:length], position+30); err != nil {
				if err == io.EOF {
					return nil
				}
				return err
			}
			if inStrSlice(f.zip64Entries, string(name[:length]), true) != -1 {
				if n, err := output.WriteAt([]byte{45, 0}, position+4); err != nil {
					return err
				} else if n != 2 {
					return io.ErrShortWrite
				}
			}
			cursor = index + 1
		}
		if readErr == io.EOF {
			return nil
		}
		offset += int64(n) - int64(len(signature)-1)
	}
}
