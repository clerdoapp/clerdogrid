<?php
declare(strict_types=1);

namespace Grid;

/**
 * Converts a JSON document from the first Grid web app to .xlsx. Streams sheet XML to temp files so a 2M-cell
 * document never has to be held as one string; the zip is assembled from those
 * files via ZipArchive (or a stored-zip fallback when the extension is missing).
 */
final class Xlsx
{
    private const BUILTIN_NUMFMT = [
        'General' => 0, '0' => 1, '0.00' => 2, '#,##0' => 3, '#,##0.00' => 4,
        '0%' => 9, '0.00%' => 10, '0.00E+00' => 11, '# ?/?' => 12, '# ??/??' => 13,
        'mm-dd-yy' => 14, 'd-mmm-yy' => 15, 'd-mmm' => 16, 'mmm-yy' => 17,
        'h:mm AM/PM' => 18, 'h:mm:ss AM/PM' => 19, 'h:mm' => 20, 'h:mm:ss' => 21, 'm/d/yy h:mm' => 22, '@' => 49,
    ];

    // --------------------------------------------------------------- builder

    /** Build the workbook; returns the path of a temp .xlsx file. */
    public static function build(array $doc): string
    {
        $dir = sys_get_temp_dir() . '/gridx_' . bin2hex(random_bytes(8));
        mkdir($dir, 0700, true);
        $parts = []; // zipPath => filePath

        $styles = new XlsxStyles($doc['styles'] ?? []);
        $strings = new SharedStrings($dir . '/sst.tmp');
        $sheets = $doc['sheets'];
        $sheetNames = [];
        foreach ($sheets as $i => $sheet) {
            $n = $i + 1;
            $path = $dir . "/sheet{$n}.xml";
            self::writeSheet($path, $sheet, $doc, $styles, $strings);
            $parts["xl/worksheets/sheet{$n}.xml"] = $path;
            if (is_file($path . '.rels')) {
                $parts["xl/worksheets/_rels/sheet{$n}.xml.rels"] = $path . '.rels';
            }
            $sheetNames[] = $sheet['name'];
        }
        $parts['xl/sharedStrings.xml'] = $strings->finish();
        $parts['xl/styles.xml'] = self::writeFile($dir . '/styles.xml', $styles->xml());
        $parts['xl/workbook.xml'] = self::writeFile($dir . '/workbook.xml', self::workbookXml($doc, $sheets));
        $parts['xl/_rels/workbook.xml.rels'] = self::writeFile($dir . '/wb.rels', self::workbookRels(count($sheets)));
        $parts['_rels/.rels'] = self::writeFile($dir . '/root.rels', self::rootRels());
        $parts['[Content_Types].xml'] = self::writeFile($dir . '/ct.xml', self::contentTypes(count($sheets)));
        $parts['docProps/app.xml'] = self::writeFile($dir . '/app.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Clerdo Grid</Application></Properties>');
        $parts['docProps/core.xml'] = self::writeFile($dir . '/core.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>' . self::esc((string)($doc['title'] ?? '')) . '</dc:title><dc:creator>Clerdo Grid</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">' . gmdate('Y-m-d\TH:i:s\Z') . '</dcterms:created></cp:coreProperties>');

        $out = $dir . '.xlsx';
        if (class_exists(\ZipArchive::class)) {
            $zip = new \ZipArchive();
            if ($zip->open($out, \ZipArchive::CREATE | \ZipArchive::OVERWRITE) !== true) {
                throw new \RuntimeException('cannot create zip');
            }
            // [Content_Types].xml first for picky readers.
            $zip->addFile($parts['[Content_Types].xml'], '[Content_Types].xml');
            foreach ($parts as $zp => $fp) {
                if ($zp !== '[Content_Types].xml') {
                    $zip->addFile($fp, $zp);
                }
            }
            $zip->close();
        } else {
            StoredZip::write($out, ['[Content_Types].xml' => $parts['[Content_Types].xml']] + $parts);
        }
        foreach ($parts as $fp) {
            @unlink($fp);
        }
        @unlink($dir . '/sst.tmp');
        @rmdir($dir);
        return $out;
    }

    private static function writeFile(string $path, string $content): string
    {
        file_put_contents($path, $content);
        return $path;
    }

    private static function writeSheet(string $path, array $sheet, array $doc, XlsxStyles $styles, SharedStrings $sst): void
    {
        $fh = fopen($path, 'wb');
        fwrite($fh, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n");
        fwrite($fh, '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');

        // Group cells by row without materialising a dense grid.
        $rows = [];
        foreach ($sheet['cells'] ?? [] as $addr => $cell) {
            [$c, $r] = self::splitAddr((string)$addr);
            $rows[$r][$c] = $cell;
        }
        ksort($rows, SORT_NUMERIC);
        $maxR = $rows === [] ? 1 : (int)array_key_last($rows);
        $maxC = 1;
        foreach ($rows as $cols) {
            $maxC = max($maxC, (int)max(array_keys($cols)));
        }
        $tabColor = '';
        if (!empty($sheet['color']) && is_string($sheet['color'])) {
            $tabColor = '<tabColor rgb="' . self::argb($sheet['color']) . '"/>';
        }
        $pr = '<sheetPr>' . $tabColor . '</sheetPr>';
        fwrite($fh, $pr);
        fwrite($fh, '<dimension ref="A1:' . self::colLetter($maxC) . $maxR . '"/>');

        $freeze = $sheet['freeze'] ?? ['rows' => 0, 'cols' => 0];
        $fr = (int)($freeze['rows'] ?? 0);
        $fc = (int)($freeze['cols'] ?? 0);
        $view = '<sheetView workbookViewId="0"' . (($sheet['gridlines'] ?? true) ? '' : ' showGridLines="0"') . '>';
        if ($fr > 0 || $fc > 0) {
            $top = self::colLetter($fc + 1) . ($fr + 1);
            $pane = $fr > 0 && $fc > 0 ? 'bottomRight' : ($fr > 0 ? 'bottomLeft' : 'topRight');
            $view .= '<pane' . ($fc > 0 ? ' xSplit="' . $fc . '"' : '') . ($fr > 0 ? ' ySplit="' . $fr . '"' : '') . ' topLeftCell="' . $top . '" activePane="' . $pane . '" state="frozen"/>';
            $view .= '<selection pane="' . $pane . '" activeCell="' . $top . '" sqref="' . $top . '"/>';
        }
        $view .= '</sheetView>';
        fwrite($fh, '<sheetViews>' . $view . '</sheetViews>');
        fwrite($fh, '<sheetFormatPr defaultRowHeight="' . self::pxToPt(24) . '"/>');

        $colWidths = $sheet['colWidths'] ?? [];
        $hiddenCols = array_flip($sheet['hiddenCols'] ?? []);
        if ($colWidths !== [] || $hiddenCols !== []) {
            $cols = [];
            foreach ($colWidths as $letter => $px) {
                if (is_string($letter) && preg_match('/^[A-Z]{1,3}$/', $letter)) {
                    $cols[$letter] = ['w' => (float)$px];
                }
            }
            foreach ($hiddenCols as $letter => $_) {
                if (is_string($letter) && preg_match('/^[A-Z]{1,3}$/', $letter)) {
                    $cols[$letter]['hidden'] = true;
                }
            }
            if ($cols !== []) {
                fwrite($fh, '<cols>');
                foreach ($cols as $letter => $def) {
                    $idx = self::colIndex($letter);
                    $w = isset($def['w']) ? max(0, round($def['w'] / 7, 2)) : 13;
                    fwrite($fh, '<col min="' . $idx . '" max="' . $idx . '" width="' . $w . '" customWidth="1"' . (!empty($def['hidden']) ? ' hidden="1"' : '') . '/>');
                }
                fwrite($fh, '</cols>');
            }
        }

        fwrite($fh, '<sheetData>');
        $rowHeights = $sheet['rowHeights'] ?? [];
        $hiddenRows = array_flip(array_map('intval', $sheet['hiddenRows'] ?? []));
        $hyperlinks = [];
        $allRows = $rows;
        foreach ($rowHeights as $rk => $_) {
            $allRows[(int)$rk] ??= [];
        }
        foreach ($hiddenRows as $rk => $_) {
            $allRows[(int)$rk] ??= [];
        }
        ksort($allRows, SORT_NUMERIC);
        foreach ($allRows as $r => $cols) {
            $attrs = ' r="' . $r . '"';
            if (isset($rowHeights[(string)$r]) || isset($rowHeights[$r])) {
                $attrs .= ' ht="' . self::pxToPt((float)($rowHeights[(string)$r] ?? $rowHeights[$r])) . '" customHeight="1"';
            }
            if (isset($hiddenRows[$r])) {
                $attrs .= ' hidden="1"';
            }
            fwrite($fh, '<row' . $attrs . '>');
            ksort($cols, SORT_NUMERIC);
            foreach ($cols as $c => $cell) {
                $ref = self::colLetter($c) . $r;
                fwrite($fh, self::cellXml($ref, $cell, $styles, $sst));
                if (isset($cell['h']['url'])) {
                    $hyperlinks[] = [$ref, (string)$cell['h']['url'], (string)($cell['h']['tip'] ?? '')];
                }
            }
            fwrite($fh, '</row>');
        }
        fwrite($fh, '</sheetData>');

        if (!empty($sheet['protection']['locked'])) {
            fwrite($fh, '<sheetProtection sheet="1" objects="1" scenarios="1"/>');
        }
        $merges = array_values(array_filter($sheet['merges'] ?? [], 'is_string'));
        if ($merges !== []) {
            fwrite($fh, '<mergeCells count="' . count($merges) . '">');
            foreach ($merges as $m) {
                fwrite($fh, '<mergeCell ref="' . self::esc($m) . '"/>');
            }
            fwrite($fh, '</mergeCells>');
        }
        if ($hyperlinks !== []) {
            // Use location-less external links: spreadsheet apps need relationships for them, so
            // write them as HYPERLINK-style tooltips via r:id-free form is not valid.
            // Instead we emit the hyperlinks element referencing rels we write below.
            fwrite($fh, '<hyperlinks>');
            foreach ($hyperlinks as $i => [$ref, $url, $tip]) {
                fwrite($fh, '<hyperlink ref="' . $ref . '" r:id="rIdH' . ($i + 1) . '"' . ($tip !== '' ? ' tooltip="' . self::esc(mb_substr($tip, 0, 255)) . '"' : '') . '/>');
            }
            fwrite($fh, '</hyperlinks>');
        }
        fwrite($fh, '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>');
        fwrite($fh, '</worksheet>');
        fclose($fh);

        if ($hyperlinks !== []) {
            $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
            foreach ($hyperlinks as $i => [$ref, $url]) {
                $rels .= '<Relationship Id="rIdH' . ($i + 1) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="' . self::esc($url) . '" TargetMode="External"/>';
            }
            $rels .= '</Relationships>';
            file_put_contents($path . '.rels', $rels);
        }
    }

    private static function cellXml(string $ref, array $cell, XlsxStyles $styles, SharedStrings $sst): string
    {
        $sIdx = isset($cell['s']) ? $styles->index((string)$cell['s']) : 0;
        $attrs = ' r="' . $ref . '"' . ($sIdx > 0 ? ' s="' . $sIdx . '"' : '');
        $v = $cell['v'] ?? null;
        $f = isset($cell['f']) && is_string($cell['f']) ? '<f>' . self::esc($cell['f']) . '</f>' : '';
        if (is_array($v) && isset($v['e'])) {
            return '<c' . $attrs . ' t="e">' . $f . '<v>' . self::esc((string)$v['e']) . '</v></c>';
        }
        if (is_bool($v)) {
            return '<c' . $attrs . ' t="b">' . $f . '<v>' . ($v ? '1' : '0') . '</v></c>';
        }
        if (is_int($v) || is_float($v)) {
            $num = is_float($v) ? self::fmtFloat($v) : (string)$v;
            return '<c' . $attrs . '>' . $f . '<v>' . $num . '</v></c>';
        }
        if (is_string($v)) {
            if ($f !== '') {
                return '<c' . $attrs . ' t="str">' . $f . '<v>' . self::esc($v) . '</v></c>';
            }
            return '<c' . $attrs . ' t="s"><v>' . $sst->index($v) . '</v></c>';
        }
        if ($f !== '' || $sIdx > 0) {
            return '<c' . $attrs . '>' . $f . '</c>';
        }
        return '';
    }

    private static function fmtFloat(float $v): string
    {
        if (!is_finite($v)) {
            return '0';
        }
        $s = rtrim(rtrim(sprintf('%.15G', $v), '0'), '.');
        return $s === '' || $s === '-' ? '0' : $s;
    }

    private static function workbookXml(array $doc, array $sheets): string
    {
        $x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $x .= '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
        $x .= '<workbookPr' . ((($doc['settings']['dateSystem'] ?? 1900) === 1904) ? ' date1904="1"' : '') . '/>';
        $active = 0;
        foreach ($sheets as $i => $sh) {
            if (($doc['activeSheet'] ?? null) === $sh['id']) {
                $active = $i;
            }
        }
        $x .= '<bookViews><workbookView activeTab="' . $active . '"/></bookViews><sheets>';
        foreach ($sheets as $i => $sh) {
            $x .= '<sheet name="' . self::esc(mb_substr((string)$sh['name'], 0, 31)) . '" sheetId="' . ($i + 1) . '"' . (!empty($sh['hidden']) ? ' state="hidden"' : '') . ' r:id="rId' . ($i + 1) . '"/>';
        }
        $x .= '</sheets>';
        $names = array_values(array_filter($doc['names'] ?? [], static fn($n) => is_array($n) && isset($n['name'], $n['ref'])));
        if ($names !== []) {
            $idIndex = [];
            foreach ($sheets as $i => $sh) {
                $idIndex[$sh['id']] = $i;
            }
            $x .= '<definedNames>';
            foreach ($names as $n) {
                $scope = isset($n['scope'], $idIndex[$n['scope']]) ? ' localSheetId="' . $idIndex[$n['scope']] . '"' : '';
                $x .= '<definedName name="' . self::esc((string)$n['name']) . '"' . $scope . '>' . self::esc(ltrim((string)$n['ref'], '=')) . '</definedName>';
            }
            $x .= '</definedNames>';
        }
        $calc = ($doc['settings']['calcMode'] ?? 'auto') === 'manual' ? 'manual' : 'auto';
        $x .= '<calcPr calcId="191029" calcMode="' . $calc . '" fullCalcOnLoad="1"/>';
        $x .= '</workbook>';
        return $x;
    }

    private static function workbookRels(int $n): string
    {
        $x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        for ($i = 1; $i <= $n; $i++) {
            $x .= '<Relationship Id="rId' . $i . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' . $i . '.xml"/>';
        }
        $x .= '<Relationship Id="rId' . ($n + 1) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
        $x .= '<Relationship Id="rId' . ($n + 2) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>';
        return $x . '</Relationships>';
    }

    private static function rootRels(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            . '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            . '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            . '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
            . '</Relationships>';
    }

    private static function contentTypes(int $n): string
    {
        $x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
        $x .= '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>';
        $x .= '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
        for ($i = 1; $i <= $n; $i++) {
            $x .= '<Override PartName="/xl/worksheets/sheet' . $i . '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
        }
        $x .= '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
        $x .= '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>';
        $x .= '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
        $x .= '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
        return $x . '</Types>';
    }

    // ---------------------------------------------------------------- utils

    public static function esc(string $s): string
    {
        $s = preg_replace('/[\x00-\x08\x0B\x0C\x0E-\x1F]/', '', $s) ?? '';
        return htmlspecialchars($s, ENT_XML1 | ENT_QUOTES, 'UTF-8');
    }

    public static function argb(string $hex): string
    {
        $h = ltrim($hex, '#');
        return strlen($h) === 8 ? strtoupper(substr($h, 6, 2) . substr($h, 0, 6)) : 'FF' . strtoupper(substr($h, 0, 6));
    }

    public static function colLetter(int $n): string
    {
        $s = '';
        while ($n > 0) {
            $n--;
            $s = chr(65 + $n % 26) . $s;
            $n = intdiv($n, 26);
        }
        return $s;
    }

    public static function colIndex(string $letters): int
    {
        $n = 0;
        foreach (str_split($letters) as $ch) {
            $n = $n * 26 + (ord($ch) - 64);
        }
        return $n;
    }

    /** @return array{0:int,1:int} [col, row] 1-based */
    public static function splitAddr(string $a): array
    {
        preg_match('/^([A-Z]+)(\d+)$/', $a, $m);
        return [self::colIndex($m[1] ?? 'A'), (int)($m[2] ?? 1)];
    }

    private static function pxToPt(float $px): string
    {
        return (string)round($px * 0.75, 2);
    }




    public static function builtinNumFmt(string $code): ?int
    {
        return self::BUILTIN_NUMFMT[$code] ?? null;
    }
}

/** Shared string table streamed to a temp file; index lookup kept in memory (strings dedup'd). */
final class SharedStrings
{
    /** @var array<string,int> */
    private array $index = [];
    private int $count = 0;
    private int $unique = 0;
    private $fh;
    private string $path;

    public function __construct(string $path)
    {
        $this->path = $path;
        $this->fh = fopen($path, 'wb');
    }

    public function index(string $s): int
    {
        $this->count++;
        $key = strlen($s) > 64 ? "\0" . md5($s) : $s;
        if (isset($this->index[$key])) {
            return $this->index[$key];
        }
        $i = $this->unique++;
        $this->index[$key] = $i;
        $space = ($s !== '' && (ctype_space($s[0]) || ctype_space($s[-1]))) ? ' xml:space="preserve"' : '';
        fwrite($this->fh, '<si><t' . $space . '>' . Xlsx::esc(mb_substr($s, 0, 32767)) . '</t></si>');
        return $i;
    }

    /** Finalise: returns path of the completed sharedStrings.xml. */
    public function finish(): string
    {
        fclose($this->fh);
        $out = $this->path . '.xml';
        $o = fopen($out, 'wb');
        fwrite($o, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' . $this->count . '" uniqueCount="' . $this->unique . '">');
        $in = fopen($this->path, 'rb');
        stream_copy_to_stream($in, $o);
        fclose($in);
        fwrite($o, '</sst>');
        fclose($o);
        $this->index = [];
        return $out;
    }
}

/** Maps Grid style ids → cellXfs indices, generating fonts/fills/borders/numFmts. */
final class XlsxStyles
{
    private array $fonts = ['<font><sz val="11"/><name val="Calibri"/></font>'];
    private array $fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
    private array $borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
    private array $numFmts = []; // code => id (custom from 164)
    private array $xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
    /** @var array<string,int> style id → xf index */
    private array $map = [];
    private array $styles;

    public function __construct(array $styles)
    {
        $this->styles = $styles;
    }

    public function index(string $styleId): int
    {
        if (isset($this->map[$styleId])) {
            return $this->map[$styleId];
        }
        $st = $this->styles[$styleId] ?? null;
        if (!is_array($st)) {
            return $this->map[$styleId] = 0;
        }
        $fontId = 0;
        $fillId = 0;
        $borderId = 0;
        $numFmtId = 0;
        $align = '';
        $prot = '';
        if (isset($st['numFmt']) && is_string($st['numFmt']) && $st['numFmt'] !== 'General') {
            $numFmtId = $this->numFmt($st['numFmt']);
        }
        if (isset($st['font']) && is_array($st['font'])) {
            $f = $st['font'];
            $x = '<font>';
            if (!empty($f['bold'])) {
                $x .= '<b/>';
            }
            if (!empty($f['italic'])) {
                $x .= '<i/>';
            }
            if (!empty($f['strike'])) {
                $x .= '<strike/>';
            }
            if (isset($f['underline']) && $f['underline'] !== 'none') {
                $x .= '<u' . ($f['underline'] === 'double' ? ' val="double"' : '') . '/>';
            }
            $x .= '<sz val="' . (is_numeric($f['size'] ?? null) ? (float)$f['size'] : 11) . '"/>';
            if (isset($f['color']) && is_string($f['color'])) {
                $x .= '<color rgb="' . Xlsx::argb($f['color']) . '"/>';
            }
            $x .= '<name val="' . Xlsx::esc(is_string($f['name'] ?? null) ? $f['name'] : 'Calibri') . '"/></font>';
            $fontId = $this->dedup($this->fonts, $x);
        }
        if (isset($st['fill']) && is_string($st['fill'])) {
            $x = '<fill><patternFill patternType="solid"><fgColor rgb="' . Xlsx::argb($st['fill']) . '"/><bgColor indexed="64"/></patternFill></fill>';
            $fillId = $this->dedup($this->fills, $x);
        }
        if (isset($st['border']) && is_array($st['border'])) {
            $x = '<border>';
            foreach (['left', 'right', 'top', 'bottom'] as $side) {
                $e = $st['border'][$side] ?? null;
                if (is_array($e) && isset($e['style'])) {
                    $x .= '<' . $side . ' style="' . Xlsx::esc((string)$e['style']) . '">' . (isset($e['color']) && is_string($e['color']) ? '<color rgb="' . Xlsx::argb($e['color']) . '"/>' : '') . '</' . $side . '>';
                } else {
                    $x .= '<' . $side . '/>';
                }
            }
            $x .= '<diagonal/></border>';
            $borderId = $this->dedup($this->borders, $x);
        }
        if (isset($st['align']) && is_array($st['align'])) {
            $a = $st['align'];
            $attrs = '';
            if (isset($a['h']) && $a['h'] !== 'general') {
                $attrs .= ' horizontal="' . Xlsx::esc((string)$a['h']) . '"';
            }
            if (isset($a['v'])) {
                $attrs .= ' vertical="' . ($a['v'] === 'middle' ? 'center' : Xlsx::esc((string)$a['v'])) . '"';
            }
            if (!empty($a['wrap'])) {
                $attrs .= ' wrapText="1"';
            }
            if (!empty($a['shrink'])) {
                $attrs .= ' shrinkToFit="1"';
            }
            if (!empty($a['indent'])) {
                $attrs .= ' indent="' . (int)$a['indent'] . '"';
            }
            if (!empty($a['rotation'])) {
                $attrs .= ' textRotation="' . (int)$a['rotation'] . '"';
            }
            if ($attrs !== '') {
                $align = '<alignment' . $attrs . '/>';
            }
        }
        if (isset($st['protection']) && is_array($st['protection'])) {
            $p = $st['protection'];
            $prot = '<protection locked="' . (($p['locked'] ?? true) ? 1 : 0) . '" hidden="' . (!empty($p['hidden']) ? 1 : 0) . '"/>';
        }
        $xf = '<xf numFmtId="' . $numFmtId . '" fontId="' . $fontId . '" fillId="' . $fillId . '" borderId="' . $borderId . '" xfId="0"'
            . ($numFmtId ? ' applyNumberFormat="1"' : '') . ($fontId ? ' applyFont="1"' : '') . ($fillId ? ' applyFill="1"' : '') . ($borderId ? ' applyBorder="1"' : '') . ($align ? ' applyAlignment="1"' : '') . ($prot ? ' applyProtection="1"' : '');
        $xf .= ($align === '' && $prot === '') ? '/>' : '>' . $align . $prot . '</xf>';
        $this->xfs[] = $xf;
        return $this->map[$styleId] = count($this->xfs) - 1;
    }

    private function numFmt(string $code): int
    {
        $b = Xlsx::builtinNumFmt($code);
        if ($b !== null) {
            return $b;
        }
        if (!isset($this->numFmts[$code])) {
            $this->numFmts[$code] = 164 + count($this->numFmts);
        }
        return $this->numFmts[$code];
    }

    private function dedup(array &$list, string $xml): int
    {
        $i = array_search($xml, $list, true);
        if ($i !== false) {
            return (int)$i;
        }
        $list[] = $xml;
        return count($list) - 1;
    }

    public function xml(): string
    {
        $x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
        if ($this->numFmts !== []) {
            $x .= '<numFmts count="' . count($this->numFmts) . '">';
            foreach ($this->numFmts as $code => $id) {
                $x .= '<numFmt numFmtId="' . $id . '" formatCode="' . Xlsx::esc((string)$code) . '"/>';
            }
            $x .= '</numFmts>';
        }
        $x .= '<fonts count="' . count($this->fonts) . '">' . implode('', $this->fonts) . '</fonts>';
        $x .= '<fills count="' . count($this->fills) . '">' . implode('', $this->fills) . '</fills>';
        $x .= '<borders count="' . count($this->borders) . '">' . implode('', $this->borders) . '</borders>';
        $x .= '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>';
        $x .= '<cellXfs count="' . count($this->xfs) . '">' . implode('', $this->xfs) . '</cellXfs>';
        $x .= '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>';
        return $x . '</styleSheet>';
    }
}

/** Minimal stored (method 0) zip writer for hosts without ext-zip. */
final class StoredZip
{
    /** @param array<string,string> $files zipPath => localPath */
    public static function write(string $out, array $files): void
    {
        $fh = fopen($out, 'wb');
        $central = '';
        $offset = 0;
        $n = 0;
        foreach ($files as $name => $path) {
            $data = file_get_contents($path);
            $crc = crc32($data);
            $len = strlen($data);
            $nameLen = strlen($name);
            $local = pack('VvvvvvVVVvv', 0x04034b50, 20, 0x0800, 0, 0, 0x21, $crc, $len, $len, $nameLen, 0) . $name;
            fwrite($fh, $local);
            fwrite($fh, $data);
            $central .= pack('VvvvvvvVVVvvvvvVV', 0x02014b50, 20, 20, 0x0800, 0, 0, 0x21, $crc, $len, $len, $nameLen, 0, 0, 0, 0, 0, $offset) . $name;
            $offset += strlen($local) + $len;
            $n++;
        }
        fwrite($fh, $central);
        fwrite($fh, pack('VvvvvVVv', 0x06054b50, 0, 0, $n, $n, strlen($central), $offset, 0));
        fclose($fh);
    }
}
