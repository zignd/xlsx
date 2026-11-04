# HTTP export example

Serves the same workbook two ways so you can compare the trade-offs:

| Endpoint | Writer | Behaviour |
| --- | --- | --- |
| `/buffered.xlsx` | `xlsx.NewWriter` | Writes worksheets to temporary files, then streams the final ZIP to the HTTP response during `Commit`. |
| `/streamed.xlsx` | `xlsx.NewStreamWriter` | Writes ZIP entries directly to the response as rows are committed. No temporary worksheet files are created. |

Both endpoints write `.xlsx` downloads with chunked transfer; neither sets
`Content-Length` because the final ZIP size is not known up front.

## Run

From the repository root:

```sh
go run ./examples/http -addr 127.0.0.1:8080 -rows 10000
```

Then download the files:

```sh
curl -o /tmp/buffered.xlsx http://127.0.0.1:8080/buffered.xlsx
curl -o /tmp/streamed.xlsx http://127.0.0.1:8080/streamed.xlsx
```

Open either file in Excel, LibreOffice or another XLSX reader.

You can also open <http://127.0.0.1:8080/> in a browser to see the available
endpoints.

## Flags

| Flag | Default | Description |
| --- | --- | --- |
| `-addr` | `127.0.0.1:8080` | HTTP listen address. |
| `-rows` | `10000` | Number of data rows per generated workbook. |

## What the example demonstrates

- Setting the correct XLSX `Content-Type` and `Content-Disposition` headers
- Passing `r.Context()` to cancel generation when the client disconnects
- `xlsx.NewWriter(http.ResponseWriter, ...)` for temp-file-backed exports
- `xlsx.NewStreamWriter(http.ResponseWriter, ...)` for zero-temp streaming
- `Options.FlushEveryRows` to push compressed bytes periodically
- `Workbook.Flush` calling `http.Flusher` when the response writer supports it
- `X-Accel-Buffering: no` for reverse proxies such as nginx

## Trade-offs

### Buffered (`NewWriter`)

- Uses temporary disk space.
- The response can be fully generated before the first byte is sent, so errors
  before `Commit` can still produce a clean HTTP error status.
- Safe when the server has local disk and you want a reliable download.

### Streamed (`NewStreamWriter`)

- No temporary worksheet files.
- Starts sending bytes as soon as rows are flushed.
- Requires one worksheet at a time: finish a sheet with `Commit` before writing
  another.
- If an error occurs after the first byte, the response is already committed and
  the client receives a truncated file. For maximum reliability, prefer the
  buffered writer or generate into a staging store first.
- The ZIP central directory is only written by `Commit`, so the download is not
  a complete/openable XLSX file until the response ends.

## Production notes

- Do not set `Content-Length` for the streaming endpoint unless you buffered the
  whole ZIP first.
- Avoid double-compressing the response; XLSX is already a ZIP archive.
- For object storage, use the provider's multipart/streaming upload API. XLSX
  cannot be resumed with range requests because the ZIP is not seekable while
  it is being produced.
- Keep `FlushEveryRows` reasonably high. Flushing after every row adds overhead;
  500–5,000 rows is usually a better starting point.
