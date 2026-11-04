# Saving to your server

BYM Office opens, edits and writes files in the browser. Your server only stores bytes: it does not
convert, render or understand documents. To open a file from a URL and save it back with the Save
button (and Ctrl/Cmd+S), the server needs two routes:

| Request | Answer |
|---|---|
| `GET /files/report.xlsx` | The file, with a strong `ETag` (e.g. `"<sha1 of the bytes>"`) |
| `PUT /files/report.xlsx` with the file as the raw body | `2xx` and the new `ETag`; **412** if `If-Match` does not match the stored file |

```html
<div id="editor" style="height: 600px"></div>
<script src="https://cdn.jsdelivr.net/npm/bymoffice/dist/bymoffice.min.js"></script>
<script>
  BymOffice.mount(document.getElementById('editor'), { file: '/files/report.xlsx', saveTo: true });
</script>
```

The `PUT` carries:

```
PUT /files/report.xlsx
Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
X-Bym-Format: xlsx                (docx or xlsx)
X-Bym-Converted-From: xls         (only for a doc/xls/RTF file opened with convert: true)
If-Match: "…"                     (only when the GET gave a strong ETag and the file is saved to the same URL)

<the file's bytes>
```

- The body is the file itself, not a multipart form. Read it as raw bytes.
- On **412** nothing is overwritten. After a save by the user the editor asks "Overwrite / Cancel" (`saveTo: { onConflict: 'fail' }` leaves that to your page); an auto-save only reports the conflict, and the question comes with the user's next save.
- **3xx is a failure**, never followed (an expired session's redirect to a login page must not swallow the file).
- 401/403 → "no permission", 423 → "locked"; every other status is reported with its number. The page gets a `SaveError` in `onError` with `code` and `status`.
- Without an `ETag` saving still works, but without conflict protection, and `autoSave` waits instead of saving blindly.

All options of `saveTo` (headers, method, credentials, timeout, conflict handling) are in the
[developer guide](kullanim-kilavuzu.md#aynı-yere-kaydetme-saveto) (Turkish).

## Tested with

On 2026-11-03 the same browser test ran against small servers in each of these stacks: open an
xlsx and a docx from the server, edit, save with the Save button, save again with the new ETag, get a
412 after someone else wrote the file, and load the bundled font and the script from the server's
static files. All passed.

| Stack | Version |
|---|---|
| Node.js (`http`) | 20.19 |
| Bun | 1.x |
| PHP built-in server, plain PHP | 8.x |
| Laravel | 13.40 |
| Python `http.server`, Flask | 3.x, Flask 3.1 |
| Django | 6.1 |
| Go `net/http` | 1.x |
| Java (JDK `HttpServer`) | 17 |
| ASP.NET Core (minimal API, Kestrel) | 8.0 |
| Ruby (WEBrick) | 2.6 |
| Dart (`dart:io`) | 3.11 |

The Laravel page was also tried in Firefox and WebKit. The examples below are the code of those
servers, trimmed; the Node one is in the repository ([`examples/server/node.mjs`](../examples/server/node.mjs))
and is tested on every change. Not tried: real proxies (nginx, IIS, Apache) in front of the app,
cross-origin saving (CORS), and mobile WebViews against these servers.

## Request size limits

A save sends the whole file. Frameworks cap request bodies, and the defaults are smaller than many
real files. When a save is refused this way the editor reports "could not be saved (HTTP 413)" (or 400 / 500), the
changes stay in the editor and the user can try again once the limit is raised.

| Where | Default | What happens above it | How to raise it |
|---|---|---|---|
| Django | 2.5 MB (`DATA_UPLOAD_MAX_MEMORY_SIZE`) | **400** (`request.body` raises `RequestDataTooBig`) | `DATA_UPLOAD_MAX_MEMORY_SIZE = 100 * 1024 * 1024` in settings |
| Laravel / PHP | 8 MB (`post_max_size`) | **413** (`ValidatePostSize`) | `post_max_size` in `php.ini`; also `memory_limit`, since the body is held in memory: with the default 128 MB a 50 MB save failed with **500** |
| ASP.NET Core (Kestrel) | 30,000,000 bytes (about 28.6 MB) | **413** | In code, `builder.WebHost.ConfigureKestrel(o => o.Limits.MaxRequestBodySize = …)`, or `[RequestSizeLimit]` on the endpoint |
| nginx (in front of any app) | 1 MB (`client_max_body_size`) | 413 | `client_max_body_size 100m;` — from the nginx documentation, not measured here |
| IIS (in front of ASP.NET) | 30,000,000 bytes (`maxAllowedContentLength`) | 404.13 | `<requestLimits maxAllowedContentLength="…">` in `web.config` — from the IIS documentation, not measured here |
| Node `http`, Go `net/http`, Flask | no limit | — | Set one yourself (the Node example answers 413 above 50 MB) |

The values in the first three rows were measured on 2026-11-03; with the limits raised, Laravel,
Django and ASP.NET Core accepted a 50 MB save.

## CSRF

Frameworks that protect state-changing requests want their token on the `PUT` too. Pass it in
`saveTo.headers`:

| Framework | Header | Token in the page |
|---|---|---|
| Laravel | `X-CSRF-TOKEN` | `<meta name="csrf-token" content="{{ csrf_token() }}">` |
| Django | `X-CSRFToken` | `{% csrf_token %}`, or the `csrftoken` cookie (call `get_token(request)` in the view) |

ASP.NET Core minimal APIs do not check antiforgery tokens on a `PUT` with a raw body; if your app adds
that check, send its header the same way.

## Laravel

`routes/web.php`:

```php
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Route;
use Illuminate\Support\Facades\Storage;

Route::get('/files/{name}', function (string $name) {
    abort_unless(user_may_open($name), 404);   // your own check; never a path from the request as is
    $path = "office/$name";
    abort_unless(Storage::exists($path), 404);
    $bytes = Storage::get($path);
    $types = [
        'xlsx' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'docx' => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    ];
    return response($bytes, 200, [
        'Content-Type' => $types[strtolower(pathinfo($name, PATHINFO_EXTENSION))] ?? 'application/octet-stream',
        'ETag' => '"' . sha1($bytes) . '"',
        'Cache-Control' => 'no-store',
    ]);
});

Route::put('/files/{name}', function (Request $request, string $name) {
    abort_unless(user_may_save($name), 403);
    $path = "office/$name";
    $want = $request->header('If-Match');
    if ($want !== null && Storage::exists($path) && $want !== '"' . sha1(Storage::get($path)) . '"') {
        return response('', 412);
    }
    $body = $request->getContent();
    Storage::put($path, $body);
    return response('', 200, ['ETag' => '"' . sha1($body) . '"']);
});
```

The Blade view:

```blade
<meta name="csrf-token" content="{{ csrf_token() }}">
<div id="editor" style="height: 600px"></div>
<script src="https://cdn.jsdelivr.net/npm/bymoffice/dist/bymoffice.min.js"></script>
<script>
  const file = '/files/report.xlsx';
  BymOffice.mount(document.getElementById('editor'), {
    file,
    saveTo: {
      url: file,
      headers: { 'X-CSRF-TOKEN': document.querySelector('meta[name=csrf-token]').content },
    },
  });
</script>
```

To serve the script yourself instead of the CDN, copy the package's `dist/` folder to
`public/vendor/bymoffice/` and use `<script src="/vendor/bymoffice/bymoffice.min.js">`; the font
files are found beside the script.

## Django

```python
import hashlib, os
from django.http import Http404, HttpResponse
from django.middleware.csrf import get_token
from django.shortcuts import render

TYPES = {
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
}

def etag(b):
    return '"' + hashlib.sha1(b).hexdigest() + '"'

def editor(request):
    get_token(request)               # sets the csrftoken cookie for the page
    return render(request, 'editor.html')

def files(request, name):
    if not user_may_open(request, name):   # your own check
        raise Http404
    p = os.path.join(STORAGE, name)
    if request.method == 'GET':
        if not os.path.isfile(p):
            raise Http404
        b = open(p, 'rb').read()
        r = HttpResponse(b, content_type=TYPES.get(os.path.splitext(name)[1], 'application/octet-stream'))
        r['ETag'] = etag(b)
        r['Cache-Control'] = 'no-store'
        return r
    if request.method == 'PUT':
        if not user_may_save(request, name):   # your own check
            return HttpResponse(status=403)
        body = request.body
        want = request.headers.get('If-Match')
        if want and os.path.isfile(p) and want != etag(open(p, 'rb').read()):
            return HttpResponse(status=412)
        open(p, 'wb').write(body)
        r = HttpResponse(status=200)
        r['ETag'] = etag(body)
        return r
    return HttpResponse(status=405)

# urls.py: path('files/<str:name>', files)
```

`editor.html`:

```html
{% csrf_token %}
<div id="editor" style="height: 600px"></div>
<script src="https://cdn.jsdelivr.net/npm/bymoffice/dist/bymoffice.min.js"></script>
<script>
  const file = '/files/report.xlsx';
  BymOffice.mount(document.getElementById('editor'), {
    file,
    saveTo: {
      url: file,
      headers: { 'X-CSRFToken': document.querySelector('[name=csrfmiddlewaretoken]').value },
    },
  });
</script>
```

## ASP.NET Core (minimal API)

```csharp
using System.Security.Cryptography;

var builder = WebApplication.CreateBuilder(args);
// Kestrel refuses bodies over 30,000,000 bytes by default.
builder.WebHost.ConfigureKestrel(o => o.Limits.MaxRequestBodySize = 100 * 1024 * 1024);
var app = builder.Build();

var types = new Dictionary<string, string> {
    [".xlsx"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    [".docx"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
};
string ETag(byte[] b) => "\"" + Convert.ToHexString(SHA1.HashData(b)).ToLowerInvariant() + "\"";
string Store(string name) => Path.Combine(app.Environment.ContentRootPath, "storage", name);

app.UseDefaultFiles();
app.UseStaticFiles();   // wwwroot/: the editor page, and dist/ if you serve the script yourself

app.MapGet("/files/{name}", (string name) => {
    if (!UserMayOpen(name) || !File.Exists(Store(name))) return Results.NotFound();   // your own check
    var b = File.ReadAllBytes(Store(name));
    return Results.Bytes(b, types.GetValueOrDefault(Path.GetExtension(name), "application/octet-stream"),
        entityTag: new Microsoft.Net.Http.Headers.EntityTagHeaderValue(ETag(b)));
});

app.MapPut("/files/{name}", async (string name, HttpRequest req, HttpResponse res) => {
    if (!UserMaySave(name)) return Results.NotFound();
    using var ms = new MemoryStream();
    await req.Body.CopyToAsync(ms);
    var body = ms.ToArray();
    var want = req.Headers.IfMatch.ToString();
    if (want != "" && File.Exists(Store(name)) && want != ETag(File.ReadAllBytes(Store(name)))) return Results.StatusCode(412);
    await File.WriteAllBytesAsync(Store(name), body);
    res.Headers.ETag = ETag(body);
    return Results.Ok();
});

app.Run();
```

## Node

[`examples/server/node.mjs`](../examples/server/node.mjs): Node's `http` module, no dependencies,
a 50 MB body limit, names limited to plain `*.docx` / `*.xlsx` names in one folder (letters of any
language), saves of one file one after another (so the `If-Match` check and the write are one step),
and each save written beside the old file and renamed, so a reader never gets half a file.

```sh
node examples/server/node.mjs ./files      # http://127.0.0.1:8080/?file=/files/report.xlsx
```

## Notes for every stack

- **Check permissions in both routes.** The examples above leave that to you (`user_may_open`, `UserMaySave` …).
  Never build a path from the request without checking it: a name like `../.env` must not reach the file system.
- **The ETag check and the write are not atomic** in the Laravel, Django and ASP.NET Core examples: two saves at
  the same instant can both pass. Lock the file (or compare-and-swap in your database) around the check and the
  write; the Node example does this by running the saves of one file one after another.
- **Saving replaces the file.** No versions are kept; keep them on the server if you need them.
- **Converted old formats.** A doc/xls/RTF opened with `convert: true` is saved as docx/xlsx and the request says so
  (`X-Bym-Converted-From`). The editor refuses to save it to the original URL unless you allow it with
  `saveTo: { allowConverted: true }`; then rename the stored file to the new extension.
- **Password-protected files** are saved encrypted with the same password; the server stores them as they come
  (they are CFB files, not ZIP, so a check that the body is a ZIP archive would refuse them).

## Without a browser

The same package runs in Node and Bun for server-side work: `open(bytes)`, `save()`, `createBlank()`,
`saveAsDocx()` / `saveAsXlsx()` for old formats, and `workbook()` / `document()` for the content.
On 2026-11-03 Node and Bun opened and saved an xlsx and two docx files byte for byte. From another language
(PHP, Python, .NET) run a small Node script as a child process; a PHP page calling one through
`proc_open` took about 80 ms per call.

```js
import { open } from 'bymoffice';
import { readFileSync, writeFileSync } from 'node:fs';

const file = await open(new Uint8Array(readFileSync('report.xls')));
writeFileSync('report.xlsx', await file.saveAsXlsx());
```
