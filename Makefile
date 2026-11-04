# Development tasks for the xlsx package.
#
# The CI workflows call the same commands so that local and CI behaviour stay
# in sync.

GO ?= go
COVERAGE_FILE ?= coverage.out
COVERAGE_MIN ?= 70
BENCHTIME ?= 1s

.PHONY: all build test test-race bench cover cover-check interop examples lint vet fmt tidy check clean

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
	$(GO) test -covermode=atomic -coverprofile=$(COVERAGE_FILE) .
	$(GO) tool cover -func=$(COVERAGE_FILE)

cover-check:
	$(GO) test -covermode=atomic -coverprofile=$(COVERAGE_FILE) .
	@total=$$($(GO) tool cover -func=$(COVERAGE_FILE) | awk '/^total:/ {gsub(/%/, "", $$3); print $$3}'); \
	echo "Total coverage: $$total%"; \
	awk -v total="$$total" -v min="$(COVERAGE_MIN)" 'BEGIN { if (total < min) { printf "coverage %.2f%% is below the %s%% gate\n", total, min; exit 1 } }'

interop:
	@tmp=$$(mktemp -d); \
	$(GO) run ./internal/testgen -out $$tmp/conformance.xlsx; \
	python3 scripts/check_openpyxl.py $$tmp/conformance.xlsx; \
	rm -rf $$tmp

examples:
	@tmp=$$(mktemp -d); trap 'rm -rf "$$tmp"' EXIT; \
	$(GO) run ./examples/basic -out $$tmp/basic.xlsx; \
	$(GO) run ./examples/read -in $$tmp/basic.xlsx -max-rows 3; \
	$(GO) run ./examples/complex -out $$tmp/complex.xlsx -rows 1000

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
