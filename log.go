package docanalysis

// Logger Logger
type Logger interface {
	Infow(msg string, keysAndValues ...interface{})
	Debugw(msg string, keysAndValues ...interface{})
	Errorw(msg string, keysAndValues ...interface{})
}

var defaultLogger = &logger{}

type logger struct {
}

func (l *logger) Infow(msg string, keysAndValues ...interface{}) {
	return
}
func (l *logger) Debugw(msg string, keysAndValues ...interface{}) {
	return
}
func (l *logger) Errorw(msg string, keysAndValues ...interface{}) {
	return
}
