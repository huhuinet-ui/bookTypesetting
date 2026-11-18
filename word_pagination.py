"""Read final document page numbers using nonprinting bookmarks and Office."""
import json
import os
from pathlib import Path
import subprocess
import tempfile


SCRIPT = r'''
param([string]$ConfigPath)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$app = $null
$document = $null
try {
    # Use WPS when installed, matching the app commonly used to print the DOCX.
    $engines = @('kwps.Application', 'Word.Application')
    if ($config.engine) { $engines = @($config.engine) }
    $selected = $null
    foreach ($engine in $engines) {
        try {
            $app = New-Object -ComObject $engine
            $selected = $engine
            break
        } catch { $lastError = $_.Exception.Message }
    }
    if ($null -eq $app) { throw ('无法启动 WPS 或 Microsoft Word：' + $lastError) }
    $app.Visible = $false
    $app.DisplayAlerts = 0
    $document = $app.Documents.Open($config.document, $false, $true)
    $document.Repaginate()
    [void]$document.Fields.Update()
    $document.Repaginate()
    $pages = @{}
    $physical = @{}
    foreach ($name in $config.bookmarks) {
        if (-not $document.Bookmarks.Exists($name)) { throw ('缺少分页书签：' + $name) }
        $range = $document.Bookmarks.Item($name).Range
        $range.Collapse(1)
        # 1 is the printed (section-adjusted) page; 3 is the physical page.
        $pages[$name] = [int]$range.Information(1)
        $physical[$name] = [int]$range.Information(3)
    }
    $total = [int]$document.ComputeStatistics(2)
    if ($config.pdf) { $document.SaveAs([ref]$config.pdf, [ref]17) }
    @{pages=$pages;physical=$physical;total=$total;engine=$selected} |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $config.result -Encoding UTF8
} catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
} finally {
    if ($null -ne $document) { $document.Close($false) | Out-Null }
    if ($null -ne $app) { $app.Quit() | Out-Null }
}
'''


def inspect_document(path, bookmarks, engine=None, pdf=None):
    """Measure this complete DOCX, preserving the caller's document unchanged."""
    with tempfile.TemporaryDirectory(prefix='book-pagination-') as directory:
        root = Path(directory)
        script = root / 'measure.ps1'
        config = root / 'config.json'
        result = root / 'result.json'
        script.write_text(SCRIPT, encoding='utf-8-sig')
        config.write_text(json.dumps({
            'document': os.path.abspath(str(path)), 'bookmarks': list(bookmarks),
            'engine': engine, 'pdf': os.path.abspath(str(pdf)) if pdf else None,
            'result': str(result),
        }, ensure_ascii=False), encoding='utf-8-sig')
        process = subprocess.run(
            ['powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', str(script), str(config)],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=360)
        if process.returncode or not result.is_file():
            raw = process.stderr or process.stdout
            message = raw.decode('utf-8', 'replace').strip()
            raise RuntimeError('Word/WPS 最终分页校验失败：' + message)
        return json.loads(result.read_text(encoding='utf-8-sig'))


def measured_pages(result, anchors):
    """Use printed page numbers, never physical positions in the PDF."""
    missing = [name for name, _ in anchors if name not in result['pages']]
    if missing:
        raise RuntimeError('分页校验缺少书签：' + ', '.join(missing))
    return {title: int(result['pages'][name]) for name, title in anchors}


def save_paginated(build, output, inspect=inspect_document, max_passes=6):
    """Rebuild until TOC and title parity agree with the complete final layout."""
    page_map, pad, engine = {}, 3, None
    previous = set()
    output = Path(output)
    with tempfile.TemporaryDirectory(prefix='book-final-', dir=str(output.parent)) as directory:
        candidate = Path(directory) / 'book.docx'
        for attempt in range(max_passes):
            document, anchors = build(page_map, pad)
            document.save(str(candidate))
            result = inspect(candidate, [name for name, _ in anchors], engine=engine)
            engine = result['engine']
            actual = measured_pages(result, anchors)
            # Include the remaining body pages, not the front matter, in padding.
            first = anchors[0][0] if anchors else None
            body_total = result['total'] - result['physical'][first] + result['pages'][first] if first else result['total']
            actual_pad = max(1, len(str(body_total)))
            if actual == page_map and actual_pad == pad:
                os.replace(str(candidate), str(output))
                return actual, body_total, pad, engine
            state = (tuple(sorted(actual.items())), actual_pad)
            if state in previous:
                break
            previous.add(state)
            page_map, pad = actual, actual_pad
    raise RuntimeError('整本书的分页尚未稳定，已停止导出，避免目录页码或标题左右位置错误。请检查临界换页处的长标题、图片和字体。')
