package docanalysis

import "errors"

var (
	// ErrNoFile ErrNoFile
	ErrNoFile = errors.New("file not exists")
	// ErrNoSupport ErrNoSupport
	ErrNoSupport = errors.New("no support file type")
)
