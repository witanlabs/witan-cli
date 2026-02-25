VERSION ?= dev
LDFLAGS := -ldflags "-X github.com/witanlabs/witan-cli/cmd.Version=$(VERSION)"

build:
	go build $(LDFLAGS) -o witan .

build-all:
	GOOS=darwin GOARCH=arm64 go build $(LDFLAGS) -o witan-darwin-arm64 .
	GOOS=darwin GOARCH=amd64 go build $(LDFLAGS) -o witan-darwin-amd64 .
	GOOS=linux GOARCH=amd64 go build $(LDFLAGS) -o witan-linux-amd64 .
	GOOS=linux GOARCH=arm64 go build $(LDFLAGS) -o witan-linux-arm64 .
	GOOS=windows GOARCH=amd64 go build $(LDFLAGS) -o witan-windows-amd64.exe .
	GOOS=windows GOARCH=arm64 go build $(LDFLAGS) -o witan-windows-arm64.exe .

dist:
	./scripts/build-dist.sh $(VERSION)

pypi-wheels:
	./scripts/build-pypi-wheels.sh $(VERSION)

test:
	go test ./...

vet:
	go vet ./...

format:
	gofmt -w .

format-check:
	@gofmt -l . | (! grep .)

clean:
	rm -f witan witan-darwin-* witan-linux-* witan-windows-*
	rm -rf dist

.PHONY: build build-all dist pypi-wheels test vet format format-check clean
