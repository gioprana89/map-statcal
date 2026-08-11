<?php
declare(strict_types=1);

header('Content-Type: application/json; charset=utf-8');
header('Cache-Control: no-store, no-cache, must-revalidate, max-age=0');
header('Pragma: no-cache');

function respond(int $status, array $payload): never {
    http_response_code($status);
    echo json_encode($payload, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
    exit;
}

function xmlDecode(string $value): string {
    return html_entity_decode($value, ENT_QUOTES | ENT_XML1, 'UTF-8');
}

function columnIndexFromCellRef(string $ref): int {
    if (!preg_match('/^([A-Z]+)/i', $ref, $m)) {
        return 0;
    }

    $letters = strtoupper($m[1]);
    $index = 0;

    for ($i = 0, $len = strlen($letters); $i < $len; $i++) {
        $index = ($index * 26) + (ord($letters[$i]) - 64);
    }

    return $index - 1;
}

/* ---------------------------------------------------------
   XLSX ZIP READER
   Uses ZipArchive when available.
   Otherwise uses a pure-PHP reader for ordinary XLSX ZIPs.
   --------------------------------------------------------- */

function purePhpZipGetEntry(string $zipData, string $wantedName): string|false {
    $eocdPos = strrpos($zipData, "\x50\x4b\x05\x06");

    if ($eocdPos === false) {
        return false;
    }

    $eocd = substr($zipData, $eocdPos + 4, 18);
    $info = unpack(
        'vdisk/vcdDisk/ventriesDisk/ventries/VcdSize/VcdOffset/vcommentLen',
        $eocd
    );

    if (!$info) {
        return false;
    }

    $pos = (int)$info['cdOffset'];
    $entries = (int)$info['entries'];

    for ($i = 0; $i < $entries; $i++) {
        if (substr($zipData, $pos, 4) !== "\x50\x4b\x01\x02") {
            return false;
        }

        $fixed = substr($zipData, $pos + 4, 42);

        $h = unpack(
            'vversionMade/vversionNeeded/vflags/vmethod/vmtime/vmdate/' .
            'Vcrc/VcompressedSize/VuncompressedSize/' .
            'vfilenameLen/vextraLen/vcommentLen/vdiskStart/' .
            'vinternalAttr/VexternalAttr/VlocalOffset',
            $fixed
        );

        if (!$h) {
            return false;
        }

        $filename = substr($zipData, $pos + 46, (int)$h['filenameLen']);

        if ($filename === $wantedName) {
            $localOffset = (int)$h['localOffset'];

            if (substr($zipData, $localOffset, 4) !== "\x50\x4b\x03\x04") {
                return false;
            }

            $localFixed = substr($zipData, $localOffset + 4, 26);

            $local = unpack(
                'vversionNeeded/vflags/vmethod/vmtime/vmdate/' .
                'Vcrc/VcompressedSize/VuncompressedSize/' .
                'vfilenameLen/vextraLen',
                $localFixed
            );

            if (!$local) {
                return false;
            }

            $dataOffset =
                $localOffset +
                30 +
                (int)$local['filenameLen'] +
                (int)$local['extraLen'];

            $compressed = substr(
                $zipData,
                $dataOffset,
                (int)$h['compressedSize']
            );

            $method = (int)$h['method'];

            if ($method === 0) {
                return $compressed;
            }

            if ($method === 8) {
                $inflated = @gzinflate($compressed);
                return $inflated === false ? false : $inflated;
            }

            return false;
        }

        $pos +=
            46 +
            (int)$h['filenameLen'] +
            (int)$h['extraLen'] +
            (int)$h['commentLen'];
    }

    return false;
}

function xlsxGetEntry(string $filePath, string $entryName): string|false {
    if (class_exists('ZipArchive')) {
        $zip = new ZipArchive();

        if ($zip->open($filePath) === true) {
            try {
                return $zip->getFromName($entryName);
            } finally {
                $zip->close();
            }
        }
    }

    $zipData = @file_get_contents($filePath);

    if ($zipData === false) {
        return false;
    }

    return purePhpZipGetEntry($zipData, $entryName);
}

/* ---------------------------------------------------------
   LIGHTWEIGHT XML READER
   Avoids requiring SimpleXML / DOM extensions.
   --------------------------------------------------------- */

function readSharedStrings(string $filePath): array {
    $xml = xlsxGetEntry($filePath, 'xl/sharedStrings.xml');

    if ($xml === false) {
        return [];
    }

    $strings = [];

    if (!preg_match_all('/<si\b[^>]*>(.*?)<\/si>/s', $xml, $items)) {
        return $strings;
    }

    foreach ($items[1] as $itemXml) {
        $text = '';

        if (preg_match_all('/<t\b[^>]*>(.*?)<\/t>/s', $itemXml, $texts)) {
            foreach ($texts[1] as $part) {
                $text .= xmlDecode(strip_tags($part));
            }
        }

        $strings[] = $text;
    }

    return $strings;
}

function attributeValue(string $attributes, string $name): string {
    $pattern = '/\b' . preg_quote($name, '/') . '\s*=\s*"([^"]*)"/i';

    if (preg_match($pattern, $attributes, $m)) {
        return xmlDecode($m[1]);
    }

    return '';
}

function cellTextFromXml(
    string $attributes,
    string $innerXml,
    array $sharedStrings
): string {
    $type = attributeValue($attributes, 't');

    if ($type === 'inlineStr') {
        $text = '';

        if (preg_match_all('/<t\b[^>]*>(.*?)<\/t>/s', $innerXml, $texts)) {
            foreach ($texts[1] as $part) {
                $text .= xmlDecode(strip_tags($part));
            }
        }

        return $text;
    }

    $raw = '';

    if (preg_match('/<v\b[^>]*>(.*?)<\/v>/s', $innerXml, $m)) {
        $raw = xmlDecode(strip_tags($m[1]));
    }

    if ($type === 's') {
        $index = (int)$raw;
        return $sharedStrings[$index] ?? '';
    }

    if ($type === 'b') {
        return $raw === '1' ? 'TRUE' : 'FALSE';
    }

    return $raw;
}

function readXlsxRows(string $filePath): array {
    if (!is_file($filePath)) {
        throw new RuntimeException('Membership database file was not found.');
    }

    if (!function_exists('gzinflate')) {
        throw new RuntimeException('PHP zlib support is not available.');
    }

    $sharedStrings = readSharedStrings($filePath);

    // For standard Excel workbooks the first worksheet is sheet1.xml.
    $sheetXml = xlsxGetEntry($filePath, 'xl/worksheets/sheet1.xml');

    if ($sheetXml === false) {
        throw new RuntimeException('Unable to read the first worksheet.');
    }

    $rows = [];

    if (!preg_match_all('/<row\b[^>]*>(.*?)<\/row>/s', $sheetXml, $rowMatches)) {
        return $rows;
    }

    foreach ($rowMatches[1] as $rowXml) {
        $row = [];

        if (preg_match_all(
            '/<c\b([^>]*)>(.*?)<\/c>/s',
            $rowXml,
            $cellMatches,
            PREG_SET_ORDER
        )) {
            foreach ($cellMatches as $cellMatch) {
                $attributes = $cellMatch[1];
                $innerXml = $cellMatch[2];

                $ref = attributeValue($attributes, 'r');
                $columnIndex = columnIndexFromCellRef($ref);

                $row[$columnIndex] = cellTextFromXml(
                    $attributes,
                    $innerXml,
                    $sharedStrings
                );
            }
        }

        if ($row) {
            ksort($row);
            $maxIndex = max(array_keys($row));
            $normalized = [];

            for ($i = 0; $i <= $maxIndex; $i++) {
                $normalized[$i] = $row[$i] ?? '';
            }

            $rows[] = $normalized;
        }
    }

    return $rows;
}

$id = isset($_GET['id']) ? trim((string)$_GET['id']) : '';

if ($id === '') {
    respond(400, [
        'success' => false,
        'message' => 'Please enter a Member ID.'
    ]);
}

if (strlen($id) > 30 || !preg_match('/^[A-Za-z0-9._-]+$/', $id)) {
    respond(400, [
        'success' => false,
        'message' => 'Invalid Member ID format.'
    ]);
}

/*
 * Package default:
 * public_html/private_data/data_member_statcal.xlsx
 *
 * The included .htaccess blocks direct browser access.
 * For stronger isolation, place private_data outside public_html
 * and change the path below.
 */
$excelPath = __DIR__ . '/private_data/data_member_statcal.xlsx';

try {
    $rows = readXlsxRows($excelPath);

    if (count($rows) < 2) {
        throw new RuntimeException('Membership database contains no records.');
    }

    $headers = array_map(
        static fn($value) => trim((string)$value),
        $rows[0]
    );

    $idColumnIndex = null;

    foreach ($headers as $index => $header) {
        if (strcasecmp($header, 'Id') === 0 || strcasecmp($header, 'ID') === 0) {
            $idColumnIndex = $index;
            break;
        }
    }

    if ($idColumnIndex === null) {
        throw new RuntimeException('The Id column was not found in the membership database.');
    }

    $member = null;

    for ($r = 1, $count = count($rows); $r < $count; $r++) {
        $row = $rows[$r];
        $rowId = trim((string)($row[$idColumnIndex] ?? ''));

        if (hash_equals($id, $rowId)) {
            $record = [];

            foreach ($headers as $colIndex => $header) {
                if ($header === '') {
                    continue;
                }

                $record[$header] = isset($row[$colIndex])
                    ? trim((string)$row[$colIndex])
                    : '';
            }

            $member = $record;
            break;
        }
    }

    if ($member === null) {
        respond(404, [
            'success' => false,
            'message' => 'Member ID not found in the STATCAL membership database.'
        ]);
    }

    $allowedFields = [
        'Id',
        'ID',
        'Name',
        'Institution',
        'Province',
        'Location',
        'Occupation',
        'Highest Education',
        'Membership Start Date',
        'Membership Expiry Date',
        'Membership Status'
    ];

    $safeMember = [];

    foreach ($allowedFields as $field) {
        if (array_key_exists($field, $member)) {
            $safeMember[$field] = $member[$field];
        }
    }

    respond(200, [
        'success' => true,
        'member' => $safeMember
    ]);

} catch (Throwable $e) {
    error_log('STATCAL member lookup error: ' . $e->getMessage());

    respond(500, [
        'success' => false,
        'message' => 'The membership verification service is temporarily unavailable.'
    ]);
}
