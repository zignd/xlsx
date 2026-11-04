# Contributing

Thanks for taking the time to contribute.

## Development setup

- Go 1.24 or newer.
- `golangci-lint` v2 (optional, required for `make lint`).

```sh
git clone https://github.com/zignd/xlsx.git
cd xlsx
go test ./...
```

## Before opening a pull request

Run the full local check:

```sh
make check
```

At minimum, run:

```sh
gofmt -w .
go vet ./...
go test -race ./...
golangci-lint run
```

The interoperability target needs `python3` and `openpyxl`:

```sh
python3 -m pip install openpyxl
make interop
```

For OOXML changes, also run:

```sh
go test -run '^$' -fuzz=FuzzReader -fuzztime=20s .
```

Please add tests for behaviour changes and benchmark coverage for hot-path
changes. Keep the public API backwards compatible unless the change is
intentional and documented in `CHANGELOG.md`.

## Documentation expectations

Changes to the public API should update `README.md`, `doc.go` and the relevant
examples. Conformance changes should update the compatibility notes in the
README and `scripts/check_openpyxl.py`.

## Commit messages

Use clear, imperative commit messages. Conventional Commits prefixes
(`feat:`, `fix:`, `docs:`, `test:`, `ci:`, `chore:`) are appreciated but not
required.

## Release process

Maintainers release through the **Tag release** GitHub Actions workflow:

1. Open **Actions → Tag release → Run workflow**.
2. Choose `patch`, `minor`, or `major`; optionally provide a prerelease suffix
   such as `rc.1`. To publish an exact version, fill in the `version` input
   instead (for example `v1.2.3`).
3. The workflow runs the test suite, creates an annotated SemVer tag, and pushes
   it.
4. The same workflow calls the reusable **Release** workflow, which validates
   the tag, runs the checks, and publishes the GitHub release.

Tags must match `vMAJOR.MINOR.PATCH[-prerelease]`. Do not move or delete a tag
after it has been published; publish a new patch release instead.
