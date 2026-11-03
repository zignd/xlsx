# Development tasks for the xlsx package.
#
# The CI workflows call the same commands so that local and CI behaviour stay
# in sync.

GO ?= go
COVERAGE_FILE ?= coverage.out
BENCHTIME ?= 1s

.PHONY: all build test test-race bench cover lint vet fmt tidy check clean

all: check

build:
	$(GO) build ./...

test:
	$(GO) test ./...

test-race:
	$(GO) test -race ./...

bench:
	$(GO) test -run '^$$' -bench . -benchmem -benchtime=$(BENCHTIME) ./...

cover:
	$(GO) test -covermode=atomic -coverprofile=$(COVERAGE_FILE) ./...
	$(GO) tool cover -func=$(COVERAGE_FILE)

lint:
	golangci-lint run

vet:
	$(GO) vet ./...

fmt:
	gofmt -w .

tidy:
	$(GO) mod tidy

check: tidy fmt vet test-race lint

clean:
	rm -f $(COVERAGE_FILE)
	$(GO) clean -testcache
