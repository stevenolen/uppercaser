# magic windows incantations to not open a command prompt
default: all

all: windows mac linux

windows:
	GOOS=windows GOARCH=amd64 go build -ldflags -H=windowsgui -o out/xlsx-uppercaser.exe

mac:
	GOOS=darwin GOARCH=amd64 go build -o out/xlsx-uppercaser-darwin

linux:
	GOOS=linux GOARCH=amd64 go build -o out/xlsx-uppercaser

PHONY:
	.windows
