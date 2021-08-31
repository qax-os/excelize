package excelize

import (
	"archive/zip"
	"io"
)

type StreamFile struct {
	f  *File
	zw *zip.Writer
	w  io.Writer //需要写入的流
}

type SheetStreamWriter struct {
	sw      *StreamWriter
	File    *StreamFile
	zipfile io.Writer
}

/* write to stream dirct,
step 1. NewStream()
step 2. NewStreamWriter() create sheet stream ,SetRow write to stream ,only one sheet can be writed at a time,
more sheets must write one by one
setp 3. SheetStreamWriter flush
setp 4. stream Close()
*/
func NewStream(w io.Writer) *StreamFile {
	return &StreamFile{f: NewFile(), w: w, zw: zip.NewWriter(w)}
}

func (s *StreamFile) Close() error {
	s.f.calcChainWriter()
	s.f.commentsWriter()
	s.f.contentTypesWriter()
	s.f.drawingsWriter()
	s.f.vmlDrawingWriter()
	s.f.workBookWriter()
	s.f.workSheetWriter()
	s.f.relsWriter()
	s.f.sharedStringsWriter()
	s.f.styleSheetWriter()

	var err error
	s.f.Pkg.Range(func(path, content interface{}) bool {
		if err != nil {
			return false
		}
		if _, ok := s.f.streams[path.(string)]; ok {
			return true
		}
		var fi io.Writer
		fi, err = s.zw.Create(path.(string))
		if err != nil {
			return false
		}
		_, err = fi.Write(content.([]byte))
		return true
	})
	if err != nil {
		return err
	}
	return s.zw.Close()
}

func (s *StreamFile) NewStreamWriter(sheet string) (*SheetStreamWriter, error) {
	s.f.NewSheet(sheet)
	sw, err := s.f.NewStreamWriter(sheet)
	if err != nil {
		return nil, err
	}
	sheetPath := s.f.sheetMap[trimSheetName(sheet)]
	fi, err := s.zw.Create(sheetPath)
	if err != nil {
		return nil, err
	}
	return &SheetStreamWriter{sw: sw, File: s, zipfile: fi}, nil
}

func (s *SheetStreamWriter) Flush() error {
	err := s.sw.Flush()
	if err != nil {
		return err
	}
	_, err = s.sw.rawData.buf.WriteTo(s.zipfile)
	if err != nil {
		return err
	}

	err = s.File.zw.Flush()
	return err
}

func (s *SheetStreamWriter) SetRow(axis string, values []interface{}, needflush bool, opts ...RowOpts) (err error) {
	err = s.sw.SetRow(axis, values, opts...)
	if err != nil {
		return
	}
	if needflush {
		_, err = s.sw.rawData.buf.WriteTo(s.zipfile)
		if err != nil {
			return
		}
		err = s.File.zw.Flush()
	}
	return
}
