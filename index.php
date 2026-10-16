<?php
declare(strict_types=1);

$sheetExtensions = ['xlsx', 'xls', 'csv'];
$musicExtensions = ['mp3', 'wav', 'ogg', 'm4a', 'aac', 'flac'];
$spreadsheetFiles = [];
$musicFiles = [];
foreach ($sheetExtensions as $extension) {
    foreach (glob(__DIR__ . '/*.' . $extension) ?: [] as $file) {
        if (is_file($file)) {
            $spreadsheetFiles[] = basename($file);
        }
    }
}
sort($spreadsheetFiles, SORT_NATURAL | SORT_FLAG_CASE);
foreach ($musicExtensions as $extension) {
    foreach (glob(__DIR__ . '/*.' . $extension) ?: [] as $file) {
        if (is_file($file)) {
            $musicFiles[] = basename($file);
        }
    }
}
sort($musicFiles, SORT_NATURAL | SORT_FLAG_CASE);
$defaultSpreadsheet = in_array('invitations.xlsx', $spreadsheetFiles, true) ? 'invitations.xlsx' : '';
$secureCookie = isset($_SERVER['HTTPS']) && $_SERVER['HTTPS'] !== '' && $_SERVER['HTTPS'] !== 'off';
session_set_cookie_params([
    'lifetime' => 0,
    'path' => '/',
    'secure' => $secureCookie,
    'httponly' => true,
    'samesite' => 'Strict',
]);
session_start();

function respondJson(array $payload, int $status = 200): never
{
    http_response_code($status);
    header('Content-Type: application/json; charset=utf-8');
    echo json_encode($payload, JSON_UNESCAPED_SLASHES | JSON_INVALID_UTF8_SUBSTITUTE);
    exit;
}

function removeTemporaryDirectory(string $directory): void
{
    foreach (scandir($directory) ?: [] as $name) {
        if ($name === '.' || $name === '..') {
            continue;
        }
        $path = $directory . '/' . $name;
        if (is_dir($path) && !is_link($path)) {
            removeTemporaryDirectory($path);
        } else {
            unlink($path);
        }
    }
    rmdir($directory);
}

function writeAttendanceWorkbook(array $records): void
{
    $target = __DIR__ . '/Attendance.xlsx';
    $temporaryWorkbook = $target . '.tmp';

    $headers = ['SN', 'Name', 'Designation', 'Institution', 'Whatsapp', 'Check-in time'];
    $rows = [$headers];
    foreach ($records as $record) {
        $rows[] = [
            $record['sn'], $record['name'], $record['designation'],
            $record['institution'], $record['whatsapp'], $record['checkedInAt'],
        ];
    }

    $sheetRows = '';
    foreach ($rows as $rowIndex => $row) {
        $number = $rowIndex + 1;
        $cells = '';
        foreach ($row as $columnIndex => $value) {
            $column = '';
            for ($index = $columnIndex + 1; $index > 0; $index = intdiv($index - 1, 26)) {
                $column = chr(65 + (($index - 1) % 26)) . $column;
            }
            $text = htmlspecialchars((string) $value, ENT_XML1 | ENT_QUOTES, 'UTF-8');
            $cells .= '<c r="' . $column . $number . '" t="inlineStr"><is><t xml:space="preserve">' . $text . '</t></is></c>';
        }
        $sheetRows .= '<row r="' . $number . '">' . $cells . '</row>';
    }

    $parts = [
        '[Content_Types].xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
        '_rels/.rels' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Attendance" sheetId="1" r:id="rId1"/></sheets></workbook>',
        'xl/_rels/workbook.xml.rels' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
        'xl/worksheets/sheet1.xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' . $sheetRows . '</sheetData></worksheet>',
    ];

    if (!class_exists('ZipArchive')) {
        throw new RuntimeException('Enable the PHP zip extension to write Attendance.xlsx.');
    }
    @unlink($temporaryWorkbook);
    $zip = new ZipArchive();
    if ($zip->open($temporaryWorkbook, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
        throw new RuntimeException('Could not create Attendance.xlsx.');
    }
    foreach ($parts as $relativePath => $contents) {
        $zip->addFromString($relativePath, $contents);
    }
    if (!$zip->close() || !is_file($temporaryWorkbook)) {
        @unlink($temporaryWorkbook);
        throw new RuntimeException('Could not create Attendance.xlsx.');
    }
    if (!@rename($temporaryWorkbook, $target)) {
        @unlink($temporaryWorkbook);
        throw new RuntimeException('Could not save Attendance.xlsx. Close it if it is open in Excel.');
    }
}

if (isset($_GET['action'])) {
    $action = (string) $_GET['action'];
    if ($action === 'template' && $_SERVER['REQUEST_METHOD'] === 'POST') {
        $input = json_decode((string) file_get_contents('php://input'), true);
        $template = trim((string) ($input['template'] ?? ''));
        if (!preg_match('/^[A-Za-z0-9_]{1,100}$/', $template)) {
            respondJson(['error' => 'Template name may contain only letters, numbers, and underscores.'], 422);
        }
        $_SESSION['bhsms_template'] = $template;
        session_write_close();
        respondJson(['saved' => true, 'template' => $template]);
    }
    if ($action === 'credentials' && $_SERVER['REQUEST_METHOD'] === 'GET') {
        $username = (string) ($_SESSION['bhsms_user'] ?? (getenv('BHSMS_USER') ?: 'InnocentWTS'));
        $configured = !empty($_SESSION['bhsms_password']) || (bool) (getenv('BHSMS_PASSWORD') ?: '');
        $publicUrl = (string) ($_SESSION['app_public_url'] ?? (getenv('APP_PUBLIC_URL') ?: ''));
        respondJson(['username' => $username, 'configured' => $configured, 'publicUrl' => $publicUrl]);
    }
    if ($action === 'credentials' && $_SERVER['REQUEST_METHOD'] === 'POST') {
        $input = json_decode((string) file_get_contents('php://input'), true);
        $username = trim((string) ($input['username'] ?? ''));
        $password = (string) ($input['password'] ?? '');
        $publicUrl = rtrim(trim((string) ($input['publicUrl'] ?? '')), '/');
        $hasSavedPassword = !empty($_SESSION['bhsms_password']) || (bool) (getenv('BHSMS_PASSWORD') ?: '');
        if ($username === '' || strlen($username) > 150 || strlen($password) > 500 || ($password === '' && !$hasSavedPassword)) {
            respondJson(['error' => 'Enter a valid WhatsApp API username and password.'], 422);
        }
        if ($publicUrl !== '' && (!filter_var($publicUrl, FILTER_VALIDATE_URL) || !in_array(strtolower((string) parse_url($publicUrl, PHP_URL_SCHEME)), ['http', 'https'], true))) {
            respondJson(['error' => 'Enter a valid public URL beginning with https://.'], 422);
        }
        $_SESSION['bhsms_user'] = $username;
        if ($password !== '') {
            $_SESSION['bhsms_password'] = $password;
        }
        if ($publicUrl !== '') {
            $_SESSION['app_public_url'] = $publicUrl;
        } else {
            unset($_SESSION['app_public_url']);
        }
        $publicUrl = (string) ($_SESSION['app_public_url'] ?? (getenv('APP_PUBLIC_URL') ?: ''));
        session_write_close();
        respondJson(['saved' => true, 'username' => $username, 'configured' => true, 'publicUrl' => $publicUrl]);
    }
    if ($action === 'attendance' && $_SERVER['REQUEST_METHOD'] === 'GET') {
        $path = __DIR__ . '/data/attendance.json';
        $records = is_file($path) ? json_decode((string) file_get_contents($path), true) : [];
        respondJson(['records' => is_array($records) ? $records : []]);
    }
    if ($action === 'attendance' && $_SERVER['REQUEST_METHOD'] === 'POST') {
        $input = json_decode((string) file_get_contents('php://input'), true);
        $sn = trim((string) ($input['sn'] ?? ''));
        if ($sn === '' || strlen($sn) > 100) {
            respondJson(['error' => 'A valid SN value is required.'], 422);
        }
        $directory = __DIR__ . '/data';
        if (!is_dir($directory) && !mkdir($directory, 0775, true)) {
            respondJson(['error' => 'Could not create the attendance data folder.'], 500);
        }
        $path = $directory . '/attendance.json';
        $records = is_file($path) ? json_decode((string) file_get_contents($path), true) : [];
        $records = is_array($records) ? $records : [];
        foreach ($records as $record) {
            if (($record['sn'] ?? '') === $sn) {
                respondJson(['record' => $record, 'duplicate' => true]);
            }
        }
        $record = [
            'sn' => $sn,
            'name' => trim((string) ($input['name'] ?? '')),
            'designation' => trim((string) ($input['designation'] ?? '')),
            'institution' => trim((string) ($input['institution'] ?? '')),
            'whatsapp' => trim((string) ($input['whatsapp'] ?? '')),
            'checkedInAt' => date(DATE_ATOM),
        ];
        $records[] = $record;
        if (file_put_contents($path, json_encode($records, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE), LOCK_EX) === false) {
            respondJson(['error' => 'Could not save attendance data.'], 500);
        }
        try {
            writeAttendanceWorkbook($records);
        } catch (Throwable $error) {
            respondJson(['error' => $error->getMessage(), 'record' => $record], 500);
        }
        respondJson(['record' => $record, 'duplicate' => false]);
    }
    if ($action === 'create-and-send' && $_SERVER['REQUEST_METHOD'] === 'POST') {
        $input = json_decode((string) file_get_contents('php://input'), true);
        $row = is_array($input['record'] ?? null) ? $input['record'] : [];
        $sn = trim((string) ($row['sn'] ?? ''));
        if (!preg_match('/^[A-Za-z0-9._-]{1,80}$/', $sn)) {
            respondJson(['error' => 'SN must contain only letters, numbers, dots, underscores, or hyphens.'], 422);
        }
        $imageData = (string) ($input['image'] ?? '');
        if (!preg_match('/^data:image\/jpeg;base64,([A-Za-z0-9+\/=]+)$/', $imageData, $matches) || strlen($matches[1]) > 2_000_000) {
            respondJson(['error' => 'A valid JPEG QR image is required.'], 422);
        }
        $image = base64_decode($matches[1], true);
        if ($image === false || !str_starts_with($image, "\xFF\xD8\xFF")) {
            respondJson(['error' => 'The QR image is not a valid JPEG.'], 422);
        }
        $qrDirectory = __DIR__ . '/QR';
        if (!is_dir($qrDirectory) && !mkdir($qrDirectory, 0775, true)) {
            respondJson(['error' => 'Could not create the QR folder.'], 500);
        }
        if (file_put_contents($qrDirectory . '/' . $sn . '.jpg', $image, LOCK_EX) === false) {
            respondJson(['error' => 'Could not save the QR image. Check folder permissions.'], 500);
        }

        $phone = preg_replace('/\D+/', '', (string) ($row['whatsapp'] ?? ''));
        if (!getenv('BHSMS_CA_BUNDLE') && is_file(__DIR__ . '/data/cacert.pem')) {
            putenv('BHSMS_CA_BUNDLE=' . __DIR__ . '/data/cacert.pem');
        }
        $password = (string) ($_SESSION['bhsms_password'] ?? (getenv('BHSMS_PASSWORD') ?: ''));
        $publicUrl = rtrim((string) ($_SESSION['app_public_url'] ?? (getenv('APP_PUBLIC_URL') ?: '')), '/');
        $missing = [];
        if ($password === '') {
            $missing[] = 'WhatsApp API password (save it in the Spreadsheet tab)';
        }
        if ($publicUrl === '') {
            $host = (string) ($_SERVER['HTTP_HOST'] ?? '');
            $hostName = strtolower((string) parse_url('//' . $host, PHP_URL_HOST));
            $isPrivate = $hostName === '' || $hostName === 'localhost' || filter_var($hostName, FILTER_VALIDATE_IP) && !filter_var($hostName, FILTER_VALIDATE_IP, FILTER_FLAG_NO_PRIV_RANGE | FILTER_FLAG_NO_RES_RANGE) || !str_contains($hostName, '.');
            if (!$isPrivate) {
                $https = !empty($_SERVER['HTTPS']) && $_SERVER['HTTPS'] !== 'off';
                $publicUrl = ($https ? 'https://' : 'http://') . $host . rtrim(dirname((string) ($_SERVER['SCRIPT_NAME'] ?? '')), '/\\');
            }
        }
        if (strlen($phone) > 10 && str_starts_with($phone, '91')) {
            $phone = substr($phone, 2);
        }
        if ($phone === '') {
            $missing[] = 'attendee WhatsApp number in the spreadsheet';
        }
        if ($missing) {
            respondJson(['saved' => true, 'sent' => false, 'error' => 'QR saved, not sent. Missing: ' . implode('; ', $missing) . '.']);
        }
        $parameters = [
            'user' => (string) ($_SESSION['bhsms_user'] ?? (getenv('BHSMS_USER') ?: 'InnocentWTS')),
            'pass' => $password,
            'sender' => getenv('BHSMS_SENDER') ?: 'BUZWAP',
            'phone' => $phone,
            'text' => (string) ($_SESSION['bhsms_template'] ?? (getenv('BHSMS_TEMPLATE') ?: 'education_conclave_qr_pass')),
            'priority' => 'wa',
            'stype' => 'normal',
            'Params' => implode(',', array_map(static fn(string $key): string => str_replace(',', ' ', (string) ($row[$key] ?? '')), ['name', 'designation', 'institution'])),
        ];
        // Same approach as the EMS app: QR Server renders the QR as a direct public PNG, so no upload is needed.
        $imageUrl = $publicUrl !== '' ? $publicUrl . '/QR/' . rawurlencode($sn) . '.jpg' : 'https://api.qrserver.com/v1/create-qr-code/?' . http_build_query(['size' => '200x200', 'format' => 'png', 'data' => $sn, 'file' => 'event-pass-' . $sn . '.png'], '', '&', PHP_QUERY_RFC3986);
        $parameters['htype'] = 'image';
        $parameters['url'] = $imageUrl;
        $apiUrl = getenv('BHSMS_API_URL') ?: 'http://bhashsms.com/api/sendmsg.php';
        $curl = curl_init($apiUrl . '?' . http_build_query($parameters, '', '&', PHP_QUERY_RFC3986));
        curl_setopt_array($curl, [CURLOPT_RETURNTRANSFER => true, CURLOPT_CONNECTTIMEOUT => 10, CURLOPT_TIMEOUT => 30]);
        if (getenv('BHSMS_CA_BUNDLE')) {
            curl_setopt($curl, CURLOPT_CAINFO, (string) getenv('BHSMS_CA_BUNDLE'));
        }
        $response = curl_exec($curl);
        $curlError = curl_error($curl);
        $httpStatus = (int) curl_getinfo($curl, CURLINFO_HTTP_CODE);
        curl_close($curl);
        $sent = $response !== false && $httpStatus >= 200 && $httpStatus < 300 && !preg_match('/\b(error|failed|invalid|denied)\b/i', (string) $response);
        respondJson([
            'saved' => true,
            'sent' => $sent,
            'error' => $sent ? '' : ($curlError ?: trim((string) $response) ?: 'WhatsApp provider returned HTTP ' . $httpStatus),
            'providerResponse' => substr(trim((string) $response), 0, 300),
            'qrFile' => $sn . '.jpg',
        ]);
    }
    respondJson(['error' => 'Unknown action.'], 404);
}

if (isset($_GET['sheet'])) {
    $name = basename((string) $_GET['sheet']);
    $path = realpath(__DIR__ . DIRECTORY_SEPARATOR . $name);
    $extension = strtolower(pathinfo($name, PATHINFO_EXTENSION));
    if ($path === false || dirname($path) !== __DIR__ || !in_array($extension, $sheetExtensions, true)) {
        http_response_code(404);
        exit;
    }
    header('Content-Type: application/octet-stream');
    header('Content-Length: ' . (string) filesize($path));
    readfile($path);
    exit;
}

$spreadsheetFilesJson = json_encode($spreadsheetFiles, JSON_HEX_TAG | JSON_HEX_APOS | JSON_HEX_AMP | JSON_HEX_QUOT);
$musicFilesJson = json_encode($musicFiles, JSON_HEX_TAG | JSON_HEX_APOS | JSON_HEX_AMP | JSON_HEX_QUOT);
$defaultSpreadsheetJson = json_encode($defaultSpreadsheet, JSON_HEX_TAG | JSON_HEX_APOS | JSON_HEX_AMP | JSON_HEX_QUOT);
?>
<!doctype html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="theme-color" content="#f5f4ef">
    <title>Gather / desk | Event attendance</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js" defer></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js" defer></script>
    <script src="jsQR.min.js" defer></script>
    <style>
        :root{color-scheme:light;--paper:#f5f4ef;--white:#fff;--ink:#17221d;--muted:#7d857f;--line:#e6e8e1;--green:#287b58;--green-dark:#1b5e42;--green-pale:#e8f3ec;--orange:#d56b42;--orange-pale:#fbefe8;--blue:#4c72a6;--blue-pale:#edf2f9;--radius:9px}
        *{box-sizing:border-box}body{margin:0;background:var(--paper);color:var(--ink);font:14px 'DM Sans',sans-serif}button,input,select{font:inherit}button{cursor:pointer}.shell{min-height:100vh}.topbar{height:68px;display:flex;align-items:center;justify-content:space-between;padding:0 max(5vw,24px);background:white;border-bottom:1px solid var(--line)}.brand{display:flex;align-items:center;gap:11px;color:var(--ink);text-decoration:none}.brand-mark{width:32px;height:32px;display:grid;place-items:center;border-radius:8px;background:var(--green);color:white}.brand-mark svg{width:18px}.brand-name{font:800 15px Manrope,sans-serif}.top-right{display:flex;align-items:center;gap:18px;color:var(--muted);font-size:12px}.live-dot{display:inline-block;width:7px;height:7px;margin-right:6px;border-radius:50%;background:#43a66f}.avatar{width:32px;height:32px;display:grid;place-items:center;border-radius:50%;background:#f2e6d6;color:#725a41;font:700 11px Manrope,sans-serif}main{max-width:1320px;margin:auto;padding:36px 32px 58px}.event-heading{display:flex;align-items:end;justify-content:space-between;gap:20px;margin-bottom:26px}.eyebrow{margin:0 0 9px;color:var(--green);font:500 10px 'DM Mono',monospace;letter-spacing:1px}h1{margin:0;font:700 27px Manrope,sans-serif}.event-meta{display:flex;align-items:center;gap:8px;margin-top:8px;color:var(--muted);font-size:13px}.event-meta svg{width:15px}.date-chip{padding:9px 12px;border:1px solid var(--line);border-radius:6px;background:white;color:#515b54;font-size:12px}.tabs{display:flex;gap:5px;width:fit-content;max-width:100%;overflow-x:auto;padding:5px;border-radius:8px;background:#e9e9e2}.tab{display:inline-flex;align-items:center;justify-content:center;gap:8px;min-height:39px;padding:0 14px;border:0;border-radius:6px;background:transparent;color:#747c75;font-size:12px;font-weight:600;white-space:nowrap}.tab svg{width:15px;height:15px}.tab.active{background:white;color:var(--ink);box-shadow:0 1px 3px #17221d12}.panel{min-height:490px;margin-top:18px;padding:23px 24px;border:1px solid var(--line);border-radius:var(--radius);background:white;animation:arrive .2s ease-out}@keyframes arrive{from{opacity:.55;transform:translateY(3px)}to{opacity:1;transform:none}}.panel-heading{display:flex;align-items:start;justify-content:space-between;gap:16px;margin-bottom:20px}h2{margin:0;font:700 17px Manrope,sans-serif}.subheading{margin:5px 0 0;color:var(--muted);font-size:12px}.toolbar{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.button{min-height:36px;display:inline-flex;align-items:center;justify-content:center;gap:7px;padding:0 12px;border:1px solid var(--line);border-radius:6px;background:white;color:#48534b;font-size:12px;font-weight:600}.button:hover{background:#f8f9f6;border-color:#cbd2c9}.button svg{width:15px;height:15px}.button.primary{border-color:var(--green);background:var(--green);color:white}.button.primary:hover{background:var(--green-dark)}.button:disabled{opacity:.45;cursor:not-allowed}.field label{display:block;margin-bottom:7px;color:#59635c;font-size:11px;font-weight:600}.field select,.field input[type=file]{width:100%;min-height:39px;padding:9px 10px;border:1px solid #dfe3dc;border-radius:6px;background:white;color:var(--ink);font-size:12px}.field input[type=file]{padding:7px}.field-hint{display:block;margin-top:6px;color:#8b948d;font-size:10px;line-height:1.5}.source-form{display:grid;grid-template-columns:minmax(220px,1fr) minmax(220px,1fr) auto;align-items:end;gap:12px;padding:16px;margin-bottom:17px;border:1px solid var(--line);border-radius:7px;background:#fafbf8}.source-summary{display:flex;align-items:center;justify-content:space-between;gap:12px;margin:14px 0;color:var(--muted);font-size:11px}.source-summary strong{color:var(--ink)}.table-wrap{overflow-x:auto;border:1px solid var(--line);border-radius:7px}table{width:100%;border-collapse:collapse;text-align:left}th{height:38px;padding:0 13px;background:#f8f9f6;color:#858d87;font:500 10px 'DM Mono',monospace;letter-spacing:.35px;text-transform:uppercase;white-space:nowrap}td{height:46px;padding:0 13px;border-top:1px solid #eef0eb;color:#414c45;font-size:12px;white-space:nowrap}td:first-child,th:first-child{width:42px;text-align:center}td input[type=checkbox]{width:15px;height:15px;accent-color:var(--green);vertical-align:middle}.empty{padding:36px 12px;color:var(--muted);text-align:center;white-space:normal}.stats{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:11px;margin:17px 0}.stat{padding:12px 14px;border:1px solid var(--line);border-radius:7px}.stat-label{color:var(--muted);font-size:11px}.stat-value{margin-top:3px;font:700 17px Manrope,sans-serif}.notice{display:flex;align-items:flex-start;gap:9px;padding:11px 12px;border:1px solid #e8e7d9;border-radius:6px;background:#fbfaf2;color:#777355;font-size:11px;line-height:1.5}.notice svg{width:15px;height:15px;flex:none}.qr-batch{display:grid;grid-template-columns:repeat(auto-fill,minmax(205px,1fr));gap:12px;margin-top:19px}.qr-card{display:grid;justify-items:center;gap:9px;padding:14px;border:1px solid var(--line);border-radius:7px;text-align:center}.qr-image{min-width:142px;min-height:142px;display:grid;place-items:center;padding:5px;border:1px solid var(--line);background:white}.qr-image img,.qr-image canvas{width:130px;height:130px}.qr-card strong{font-size:12px}.qr-card span{color:var(--muted);font-size:10px}.qr-card .toolbar{justify-content:center}.attendance-overview{display:flex;align-items:center;justify-content:space-between;gap:20px;padding:16px 18px;margin-bottom:18px;border-radius:7px;background:#f3f7f2}.big-number{font:700 27px Manrope,sans-serif}.attendance-overview p{margin:3px 0 0;color:#7a857c;font-size:11px}.progress{width:38%;height:6px;overflow:hidden;border-radius:6px;background:#dce5dd}.progress span{display:block;height:100%;border-radius:inherit;background:var(--green);transition:width .2s}.camera-box{position:relative;min-height:300px;display:grid;place-items:center;overflow:hidden;border-radius:8px;background:#18231e}.camera-box video{position:absolute;width:100%;height:100%;object-fit:cover}.camera-empty{z-index:1;display:grid;justify-items:center;gap:11px;color:#bac5bd;text-align:center}.camera-empty svg{width:36px;color:#8ea194}.camera-empty p{font-size:12px;line-height:1.6}.camera-actions{position:absolute;z-index:2;bottom:16px;left:50%;display:flex;gap:8px;transform:translateX(-50%)}.scan-result{margin-top:12px;padding:11px 13px;border-radius:6px;background:var(--green-pale);color:var(--green-dark);font-size:12px}.music-layout{display:grid;grid-template-columns:1fr 300px;gap:20px}.dropzone{min-height:210px;display:grid;place-items:center;padding:20px;border:1px dashed #dce2da;border-radius:8px;background:#fafbf8;text-align:center}.dropzone.dragover{border-color:var(--green);background:var(--green-pale)}.drop-content{display:grid;justify-items:center;gap:9px}.drop-content strong{font:700 13px Manrope,sans-serif}.drop-content p{margin:0;color:var(--muted);font-size:11px}.playlist{overflow:hidden;border:1px solid var(--line);border-radius:7px}.playlist-head{padding:11px 13px;background:#f8f9f6;color:#828a83;font:500 10px 'DM Mono',monospace;text-transform:uppercase}.track{display:flex;align-items:center;gap:10px;min-height:47px;padding:7px 11px;border-top:1px solid #eef0eb}.track-info{flex:1;min-width:0}.track-info strong{display:block;overflow:hidden;font-size:11px;text-overflow:ellipsis;white-space:nowrap}.track-info span{color:var(--muted);font-size:10px}.track button{border:0;background:none;color:var(--green)}.music-controls{display:flex;align-items:center;gap:9px;padding:11px;border-top:1px solid var(--line)}.music-controls input{flex:1;accent-color:var(--green)}.music-controls span{color:var(--muted);font:10px 'DM Mono',monospace}.toast{position:fixed;right:22px;bottom:22px;z-index:5;padding:11px 15px;border-radius:7px;background:#20372a;color:white;box-shadow:0 6px 22px #17221d30;font-size:12px}.hidden{display:none!important}.error{color:#a8412c!important}.visually-hidden{position:absolute;width:1px;height:1px;padding:0;margin:-1px;overflow:hidden;clip:rect(0,0,0,0);white-space:nowrap;border:0}
        @media(max-width:760px){.topbar{padding:0 20px}main{padding:27px 19px 40px}.source-form{grid-template-columns:1fr 1fr}.source-form>.button{grid-column:1/-1}.music-layout{grid-template-columns:1fr}.camera-box{min-height:260px}}
        .camera-box{width:300px;height:300px;min-height:0;max-width:100%;margin:0 auto}.camera-box video{width:100%;height:100%;object-fit:cover}
        @media(max-width:540px){.topbar{height:60px}.top-right{gap:10px}.top-right .live-label{display:none}main{padding:23px 13px 32px}.event-heading{align-items:start;flex-direction:column;margin-bottom:19px}h1{font-size:23px}.tabs{width:100%;justify-content:flex-start}.tab{min-width:42px;padding:0 11px}.tab span{display:none}.tab.active span{display:inline}.panel{min-height:420px;padding:17px 12px}.panel-heading{flex-direction:column}.toolbar{width:100%}.source-form{grid-template-columns:1fr}.source-form>.button{grid-column:auto}.stats{gap:7px}.stat{padding:9px}.stat-label{font-size:10px}.stat-value{font-size:15px}.attendance-overview{padding:13px}.progress{width:30%}}
        .field input[type=text],.field input[type=password]{width:100%;min-height:39px;padding:9px 10px;border:1px solid #dfe3dc;border-radius:6px;background:white;color:var(--ink);font-size:12px}
        .sheet-error{grid-column:1/-1;margin:0;color:#a8412c;font-size:11px;line-height:1.5;overflow-wrap:anywhere}
        .tab.active{background:#dff2e5;color:#1b5e42;box-shadow:inset 0 -2px #d56b42}
        .panel{border-top:3px solid #48a878}
        .table-wrap{width:100%;max-width:100%;overflow-x:auto;overscroll-behavior-x:contain;scrollbar-color:#78ae90 #edf3ee;scrollbar-width:thin}
        .table-wrap table{min-width:680px}
        .table-wrap th,.table-wrap td{font-size:12px;letter-spacing:0}
        .table-wrap thead th{background:#287b67;color:#fff;border-bottom:3px solid #efb452}
        .table-wrap tbody tr:nth-child(4n + 1) td{background:#fff}
        .table-wrap tbody tr:nth-child(4n + 2) td{background:#e8f5ec}
        .table-wrap tbody tr:nth-child(4n + 3) td{background:#edf4fb}
        .table-wrap tbody tr:nth-child(4n) td{background:#fff0e6}
        .table-wrap tbody tr:hover td{background:#fff4c9}
        .send-qr-table{min-width:860px}
        .send-qr-table .whatsapp-column{background:#fff}
        .send-qr-table tbody tr:nth-child(4n + 2) .whatsapp-column{background:#e8f5ec}
        .send-qr-table tbody tr:nth-child(4n + 3) .whatsapp-column{background:#edf4fb}
        .send-qr-table tbody tr:nth-child(4n) .whatsapp-column{background:#fff0e6}
        .send-qr-table tbody tr:hover .whatsapp-column{background:#fff4c9}
        .send-qr-table thead .whatsapp-column{background:#287b67}
        .send-qr-table{min-width:860px}.send-qr-table .whatsapp-column{position:sticky;right:0;z-index:1;background:white;box-shadow:-5px 0 8px -8px #17221d80}.send-qr-table thead .whatsapp-column{z-index:2;background:#f8f9f6}
    </style>
</head>
<body>
<div class="shell">
    <header class="topbar"><a class="brand" href="#" aria-label="Event desk home"><span class="brand-mark"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 7.5 12 3l8 4.5v9L12 21l-8-4.5z"/><path d="m4.5 7.8 7.5 4.4 7.5-4.4M12 12.2V21"/></svg></span><span class="brand-name">Gather / desk</span></a><div class="top-right"><span class="live-label"><i class="live-dot"></i>Event workspace</span><span class="avatar" aria-label="Organizer">AS</span></div></header>
    <main>
        <section class="event-heading"><div><p class="eyebrow">EVENT OPERATIONS</p><h1>Attendance desk</h1><div class="event-meta">Spreadsheet-powered guest management</div></div><div class="date-chip">EVENT CHECK-IN</div></section>
        <nav class="tabs" role="tablist" aria-label="Event tools">
            <button class="tab active" role="tab" aria-selected="true" data-tab="scan-qr"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M4 8V5a1 1 0 0 1 1-1h3m8 0h3a1 1 0 0 1 1 1v3m0 8v3a1 1 0 0 1-1 1h-3m-8 0H5a1 1 0 0 1-1-1v-3M8 8h8v8H8z"/><path d="M2 12h3m14 0h3"/></svg><span>Scan QR</span></button>
            <button class="tab" role="tab" aria-selected="false" data-tab="send-qr"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><path d="M14 14h3v3h-3zm4 4h3v3h-3zm3-4v1m-7 4v2"/></svg><span>Send QR</span></button>
            <button class="tab" role="tab" aria-selected="false" data-tab="spreadsheet"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9m6 12V9"/></svg><span>Spreadsheet</span></button>
            <button class="tab" role="tab" aria-selected="false" data-tab="attendance"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="m5 12 4 4L19 6"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/></svg><span>Attendance</span></button>
            <button class="tab" role="tab" aria-selected="false" data-tab="music"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M9 18V5l12-2v13"/><circle cx="6" cy="18" r="3"/><circle cx="18" cy="16" r="3"/></svg><span>Music</span></button>
        </nav>
        <section class="panel" id="panel" role="tabpanel" tabindex="0"></section>
    </main>
</div>
<input class="visually-hidden" id="spreadsheet-input" type="file" accept=".xlsx,.xls,.csv">
<input class="visually-hidden" id="music-input" type="file" accept="audio/*" multiple>
<div id="toast-root" aria-live="polite"></div><audio id="audio-player"></audio>
<script>
(() => {
    const serverSheets = <?= $spreadsheetFilesJson ?: '[]' ?>;
    const serverMusic = <?= $musicFilesJson ?: '[]' ?>;
    const preferredSpreadsheet = <?= $defaultSpreadsheetJson ?: '""' ?>;
    let apiUsername = <?= json_encode((string) ($_SESSION['bhsms_user'] ?? (getenv('BHSMS_USER') ?: 'InnocentWTS')), JSON_HEX_TAG | JSON_HEX_APOS | JSON_HEX_AMP | JSON_HEX_QUOT) ?>;
    let apiConfigured = <?= (!empty($_SESSION['bhsms_password']) || (bool) (getenv('BHSMS_PASSWORD') ?: '')) ? 'true' : 'false' ?>;
    let apiTemplate = <?= json_encode((string) ($_SESSION['bhsms_template'] ?? (getenv('BHSMS_TEMPLATE') ?: 'education_conclave_qr_pass')), JSON_HEX_TAG | JSON_HEX_APOS | JSON_HEX_AMP | JSON_HEX_QUOT) ?>;
    const requiredHeaders = ['sn', 'name', 'designation', 'institution', 'whatsapp'];
    const storageKey = 'gather-attendance-sheet-v1';
    const panel = document.querySelector('#panel');
    const audio = document.querySelector('#audio-player');
    let records = [];
    let sourceName = '';
    let spreadsheetError = '';
    let activeTab = 'scan-qr';
    let attendanceRows = [];
    let tracks = serverMusic.map(name => ({ name, url: encodeURIComponent(name) }));
    let trackIndex = -1;
    let stream = null;
    let detector = null;
    let scanTimer = null;
    let scanBusy = false;
    const escapeHtml = value => String(value ?? '').replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char]);
    const iconInfo = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="12" cy="12" r="9"/><path d="M12 11v5m0-8h.01"/></svg>';
    const iconUpload = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M12 16V4m-5 5 5-5 5 5M4 20h16"/></svg>';
    const toast = message => { const root = document.querySelector('#toast-root'); root.innerHTML = `<div class="toast">${escapeHtml(message)}</div>`; setTimeout(() => { root.innerHTML = ''; }, 2800); };
    const saveCache = () => localStorage.setItem(storageKey, JSON.stringify({ records, sourceName }));
    const phoneNumber = value => String(value ?? '').replace(/\D/g, '');

    function renderSpreadsheet() {
        const selectedSheet = serverSheets.includes(sourceName) ? sourceName : preferredSpreadsheet;
        panel.innerHTML = `<div class="panel-heading"><div><h2>Choose the guest spreadsheet</h2><p class="subheading">The first row must contain SN, Name, Designation, Institution, and Whatsapp.</p></div></div><div class="source-form"><div class="field"><label for="server-sheet">Spreadsheet in this app folder</label><select id="server-sheet"><option value="">Choose a fileâ€¦</option>${serverSheets.map(name => `<option value="${escapeHtml(name)}" ${name === selectedSheet ? 'selected' : ''}>${escapeHtml(name)}</option>`).join('')}</select><span class="field-hint">${preferredSpreadsheet ? `${escapeHtml(preferredSpreadsheet)} is the default when available.` : 'Select a .xlsx, .xls, or .csv file beside the PHP app.'}</span></div><div class="field"><label for="spreadsheet-input-visible">Browse for a spreadsheet</label><input id="spreadsheet-input-visible" type="file" accept=".xlsx,.xls,.csv"><span class="field-hint" id="file-name">No file selected</span></div><button class="button primary" id="load-sheet">Load spreadsheet</button></div><div class="source-form"><div class="field"><label for="whatsapp-username">WhatsApp API username</label><input id="whatsapp-username" type="text" autocomplete="username" value="${escapeHtml(apiUsername)}"></div><div class="field"><label for="whatsapp-password">WhatsApp API password</label><input id="whatsapp-password" type="password" autocomplete="new-password" placeholder="Enter API password"></div><button class="button" id="save-api-credentials">Save credentials</button></div><p class="field-hint" id="credentials-status">${apiConfigured ? 'Credentials are available for this server session.' : 'Credentials are not set for this session.'} Password is kept in the server session only.</p><div class="source-form"><div class="field"><label for="whatsapp-template">WhatsApp template name</label><input id="whatsapp-template" type="text" autocomplete="off" value="${escapeHtml(apiTemplate)}"></div><button class="button" id="save-template">Save template</button></div><div class="notice">${iconInfo}<span>Column headings are read from the first row. The spreadsheet stays on your device; this browser keeps a working snapshot for check-in.</span></div><div class="source-summary"><span>Loaded source: <strong id="source-name">${escapeHtml(sourceName || 'No spreadsheet loaded')}</strong></span><span id="record-count">${records.length} records</span></div><div class="table-wrap"><table><thead><tr><th>SN</th><th>Name</th><th>Designation</th><th>Institution</th><th>Whatsapp</th></tr></thead><tbody>${records.length ? records.slice(0, 100).map(row => `<tr><td>${escapeHtml(row.sn)}</td><td>${escapeHtml(row.name)}</td><td>${escapeHtml(row.designation)}</td><td>${escapeHtml(row.institution)}</td><td>${escapeHtml(row.whatsapp)}</td></tr>`).join('') : '<tr><td colspan="5" class="empty">Choose a spreadsheet to load attendee records.</td></tr>'}</tbody></table></div>${records.length > 100 ? `<p class="field-hint">Showing the first 100 of ${records.length} records.</p>` : ''}`;
        const loadButton = document.querySelector('#load-sheet');
        const errorMessage = document.createElement('p');
        errorMessage.id = 'sheet-error';
        errorMessage.className = `sheet-error${spreadsheetError ? '' : ' hidden'}`;
        errorMessage.setAttribute('role', 'alert');
        errorMessage.textContent = spreadsheetError;
        loadButton.insertAdjacentElement('afterend', errorMessage);
        const fileInput = document.querySelector('#spreadsheet-input');
        document.querySelector('#spreadsheet-input-visible').addEventListener('change', event => {
            const file = event.target.files[0];
            document.querySelector('#file-name').textContent = file ? file.name : 'No file selected';
            fileInput.files = event.target.files;
        });
        document.querySelector('#server-sheet').addEventListener('change', () => { fileInput.value = ''; document.querySelector('#file-name').textContent = 'No file selected'; });
        document.querySelector('#load-sheet').addEventListener('click', loadSelectedSheet);
        document.querySelector('#save-api-credentials').addEventListener('click', saveApiCredentials);
        document.querySelector('#save-template').addEventListener('click', saveTemplate);
    }

    async function loadSelectedSheet() {
        const serverName = document.querySelector('#server-sheet').value;
        const file = document.querySelector('#spreadsheet-input').files[0];
        showSpreadsheetError('');
        if (!serverName && !file) { showSpreadsheetError('Choose a spreadsheet first.'); return; }
        try {
            let data;
            let name;
            if (file) { data = await file.arrayBuffer(); name = file.name; }
            else { name = serverName; const response = await fetch(`?sheet=${encodeURIComponent(serverName)}`); if (!response.ok) throw new Error('Could not read that file from the app folder.'); data = await response.arrayBuffer(); }
            applySpreadsheet(data, name);
        } catch (error) { showSpreadsheetError(error.message || 'Could not load this spreadsheet.'); }
    }

    function showSpreadsheetError(message) {
        spreadsheetError = String(message || '');
        const output = document.querySelector('#sheet-error');
        if (output) {
            output.textContent = spreadsheetError;
            output.classList.toggle('hidden', spreadsheetError === '');
        } else if (spreadsheetError) toast(spreadsheetError);
    }

    function applySpreadsheet(data, name) {
        const workbook = XLSX.read(data, { type: 'array', raw: false });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const matrix = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
        if (!matrix.length) throw new Error('The selected spreadsheet is empty.');
        const headers = matrix[0].map(value => String(value).replace(/^\uFEFF/, '').trim().toLowerCase());
        const indexes = Object.fromEntries(requiredHeaders.map(header => [header, headers.indexOf(header)]));
        const missing = requiredHeaders.filter(header => indexes[header] < 0);
        if (missing.length) throw new Error(`Missing column${missing.length > 1 ? 's' : ''}: ${missing.join(', ')}.`);
        const nextRecords = matrix.slice(1).filter(row => String(row[indexes.sn] ?? '').trim() !== '').map(row => ({
            sn: String(row[indexes.sn]).trim(), name: String(row[indexes.name] ?? '').trim(),
            designation: String(row[indexes.designation] ?? '').trim(), institution: String(row[indexes.institution] ?? '').trim(),
            whatsapp: String(row[indexes.whatsapp] ?? '').trim()
        }));
        if (!nextRecords.length) throw new Error('No attendee rows with an SN value were found.');
        records = nextRecords;
        sourceName = name;
        saveCache();
        render();
        toast(`${records.length} spreadsheet records loaded.`);
    }

    async function loadSpreadsheetFromFolder(name) {
        try {
            const response = await fetch(`?sheet=${encodeURIComponent(name)}`);
            if (!response.ok) throw new Error(`Could not load ${name}.`);
            applySpreadsheet(await response.arrayBuffer(), name);
        } catch (error) { showSpreadsheetError(error.message || 'Could not load the default spreadsheet.'); }
    }

    async function saveTemplate() {
        const input = document.querySelector('#whatsapp-template');
        const button = document.querySelector('#save-template');
        const template = input.value.trim();
        if (!/^[A-Za-z0-9_]{1,100}$/.test(template)) { toast('Template name may contain only letters, numbers, and underscores.'); return; }
        button.disabled = true;
        try {
            const response = await fetch('?action=template', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ template })
            });
            const result = await response.json();
            if (!response.ok) throw new Error(result.error || 'Could not save the template.');
            apiTemplate = result.template;
            toast('WhatsApp template saved.');
        } catch (error) { toast(error.message || 'Could not save the template.'); }
        finally { button.disabled = false; }
    }

    async function saveApiCredentials() {
        const usernameInput = document.querySelector('#whatsapp-username');
        const passwordInput = document.querySelector('#whatsapp-password');
        const status = document.querySelector('#credentials-status');
        const button = document.querySelector('#save-api-credentials');
        const username = usernameInput.value.trim();
        if (!username || !passwordInput.value) { status.textContent = 'Enter both the API username and password.'; return; }
        button.disabled = true;
        try {
            const response = await fetch('?action=credentials', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ username, password: passwordInput.value })
            });
            const result = await response.json();
            if (!response.ok) throw new Error(result.error || 'Could not save credentials.');
            apiUsername = result.username;
            apiConfigured = true;
            passwordInput.value = '';
            status.textContent = 'Credentials saved to the server session.';
            toast('WhatsApp API credentials saved.');
        } catch (error) { status.textContent = error.message || 'Could not save credentials.'; }
        finally { button.disabled = false; }
    }

    function renderSendQr() {
        panel.innerHTML = `<div class="panel-heading"><div><h2>Create and send QR passes</h2><p class="subheading">Select rows to create an SN QR, save it in QR/, and send it to the attendee's WhatsApp.</p></div><button class="button primary" id="create-qr" ${records.length ? '' : 'disabled'}>Create and Send QR</button></div><div class="source-summary"><span>Source: <strong>${escapeHtml(sourceName || 'No spreadsheet loaded')}</strong></span><span>${records.length} records</span></div><div class="table-wrap"><table><thead><tr><th><input id="select-all" type="checkbox" aria-label="Select all rows" ${records.length ? '' : 'disabled'}></th><th>SN</th><th>Name</th><th>Designation</th><th>Institution</th><th>Whatsapp</th><th>Send status</th></tr></thead><tbody>${records.length ? records.map((row, index) => `<tr><td><input type="checkbox" class="row-select" value="${index}" aria-label="Select ${escapeHtml(row.name || row.sn)}"></td><td>${escapeHtml(row.sn)}</td><td>${escapeHtml(row.name)}</td><td>${escapeHtml(row.designation)}</td><td>${escapeHtml(row.institution)}</td><td>${escapeHtml(row.whatsapp)}</td><td data-delivery-status="${index}">Not sent</td></tr>`).join('') : '<tr><td colspan="7" class="empty">Load a spreadsheet in the Spreadsheet tab to see records here.</td></tr>'}</tbody></table></div>`;
        const sendTable = panel.querySelector('.table-wrap table');
        sendTable.classList.add('send-qr-table');
        sendTable.querySelectorAll('thead th:last-child, tbody tr td:last-child:not([colspan])').forEach(cell => cell.classList.add('whatsapp-column'));
        document.querySelector('#select-all')?.addEventListener('change', event => panel.querySelectorAll('.row-select').forEach(box => { box.checked = event.target.checked; }));
        document.querySelector('#create-qr')?.addEventListener('click', createSelectedQrs);
    }

    function qrJpeg(sn) {
        if (!window.QRCode) return Promise.reject(new Error('QR generator is unavailable. Check your connection.'));
        return new Promise((resolve, reject) => {
            const holder = document.createElement('div');
            holder.style.cssText = 'position:fixed;left:-10000px;top:0';
            document.body.append(holder);
            new QRCode(holder, { text: String(sn), width: 260, height: 260, colorDark: '#17221d', colorLight: '#ffffff', correctLevel: QRCode.CorrectLevel.M });
            setTimeout(() => {
                const canvas = holder.querySelector('canvas');
                if (canvas) { const url = canvas.toDataURL('image/jpeg', 0.95); holder.remove(); resolve(url); return; }
                const image = holder.querySelector('img');
                if (image) { image.onload = () => { const canvasImage = document.createElement('canvas'); canvasImage.width = 260; canvasImage.height = 260; canvasImage.getContext('2d').drawImage(image, 0, 0, 260, 260); holder.remove(); resolve(canvasImage.toDataURL('image/jpeg', 0.95)); }; image.onerror = () => { holder.remove(); reject(new Error('QR image could not be created.')); }; return; }
                holder.remove(); reject(new Error('QR image could not be created.'));
            }, 50);
        });
    }

    async function createSelectedQrs() {
        const indexes = [...panel.querySelectorAll('.row-select:checked')].map(box => Number(box.value));
        if (!indexes.length) { toast('Select at least one spreadsheet row.'); return; }
        const button = document.querySelector('#create-qr'); button.disabled = true;
        let sentCount = 0;
        let failedCount = 0;
        try {
            for (const index of indexes) {
                const row = records[index];
                const status = panel.querySelector(`[data-delivery-status="${CSS.escape(String(index))}"]`);
                status.textContent = 'Generating QR and sendingâ€¦';
                try {
                    const image = await qrJpeg(row.sn);
                    const response = await fetch('?action=create-and-send', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ record: row, image })
                    });
                    const delivery = await response.json();
                    if (!response.ok || !delivery.saved) throw new Error(delivery.error || 'Could not save this QR image.');
                    if (delivery.sent) {
                        sentCount++;
                        status.textContent = 'Sent';
                    } else {
                        failedCount++;
                        status.textContent = `Saved, not sent: ${delivery.error || 'Provider did not confirm delivery.'}`;
                    }
                } catch (error) {
                    failedCount++;
                    status.textContent = `Failed: ${error.message || 'Could not create or send the QR.'}`;
                }
            }
            toast(`${sentCount} sent, ${failedCount} failed or not sent.`);
        } catch (error) { toast(error.message || 'Could not process QR passes.'); }
        finally { button.disabled = false; }
    }

    function renderAttendance() {
        const arrived = attendanceRows.length;
        const percent = records.length ? Math.min(100, Math.round(arrived / records.length * 100)) : 0;
        const sorted = [...attendanceRows].sort((a, b) => b.checkedInAt.localeCompare(a.checkedInAt));
        panel.innerHTML = `<div class="panel-heading"><div><h2>Attendance</h2><p class="subheading">Guests welcomed at the door. Each scan is also appended to Attendance.xlsx.</p></div><button class="button" id="export-attendance" ${attendanceRows.length ? '' : 'disabled'}>Export attendance</button></div><div class="attendance-overview"><div><div class="big-number">${arrived} <span style="color:#929c93;font-size:16px;font-weight:500">/ ${records.length}</span></div><p>guests checked in Â· ${percent}% of loaded records</p></div><div class="progress" role="progressbar" aria-valuenow="${percent}" aria-valuemin="0" aria-valuemax="100"><span style="width:${percent}%"></span></div></div><div class="table-wrap"><table><thead><tr><th>SN</th><th>Name</th><th>Designation</th><th>Institution</th><th>Whatsapp</th><th>Check-in time</th></tr></thead><tbody>${sorted.length ? sorted.map(row => `<tr><td>${escapeHtml(row.sn)}</td><td>${escapeHtml(row.name)}</td><td>${escapeHtml(row.designation)}</td><td>${escapeHtml(row.institution)}</td><td>${escapeHtml(row.whatsapp)}</td><td>${escapeHtml(new Date(row.checkedInAt).toLocaleString())}</td></tr>`).join('') : '<tr><td colspan="6" class="empty">No guests have been scanned yet.</td></tr>'}</tbody></table></div>`;
        document.querySelector('#export-attendance')?.addEventListener('click', exportAttendance);
    }

    function renderScan() {
        panel.innerHTML = `<div class="panel-heading"><div><h2>Scan attendee QR</h2><p class="subheading">Scan the code to check in the matching SN from the loaded spreadsheet.</p></div></div><div class="camera-box"><video id="camera-video" playsinline muted></video><div class="camera-empty" id="camera-empty"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M14 5h-4L8 8H5a2 2 0 0 0-2 2v8a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-8a2 2 0 0 0-2-2h-3z"/><circle cx="12" cy="13" r="3"/></svg><p>Camera preview appears here.<br>Allow access when prompted.</p></div><div class="camera-actions"><button class="button primary" id="start-camera">Start camera</button><button class="button hidden" id="stop-camera">Stop camera</button></div></div><div id="scan-result" class="scan-result hidden" aria-live="polite"></div>`;
        document.querySelector('#start-camera').addEventListener('click', startScanner);
        document.querySelector('#stop-camera').addEventListener('click', stopScanner);
    }

    async function startScanner() {
        if (!records.length) { toast('Load a spreadsheet before scanning.'); return; }
        if (!navigator.mediaDevices?.getUserMedia) { toast('Camera access needs HTTPS (or http://localhost). Open this page over HTTPS.'); return; }
        if (typeof jsQR !== 'function') { toast('QR decoder is not loaded. Reload the page.'); return; }
        const status = document.querySelector('#scan-result');
        try {
            detector = { detect: detectWithJsQr };
            try {
                stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: 'environment' }, width: { ideal: 1280 }, height: { ideal: 720 } } });
            } catch (constraintError) {
                stream = await navigator.mediaDevices.getUserMedia({ video: true });
            }
            const video = document.querySelector('#camera-video'); video.srcObject = stream; await video.play();
            document.querySelector('#camera-empty').classList.add('hidden'); document.querySelector('#start-camera').classList.add('hidden'); document.querySelector('#stop-camera').classList.remove('hidden');
            status.textContent = 'Camera is ready. Point it at a QR code.'; status.classList.remove('hidden');
            scanTimer = setInterval(scanFrame, 250);
        } catch (error) { toast(error.name === 'NotAllowedError' ? 'Camera permission was denied.' : 'Unable to start the camera: ' + (error.message || error.name)); stopScanner(); }
    }

    async function detectWithJsQr(video) {
        const canvas = detectWithJsQr.canvas ||= document.createElement('canvas');
        const scale = Math.min(1, 720 / video.videoWidth);
        canvas.width = Math.round(video.videoWidth * scale); canvas.height = Math.round(video.videoHeight * scale);
        const context = canvas.getContext('2d', { willReadFrequently: true });
        context.drawImage(video, 0, 0, canvas.width, canvas.height);
        const image = context.getImageData(0, 0, canvas.width, canvas.height);
        const code = jsQR(image.data, image.width, image.height);
        return code ? [{ rawValue: code.data }] : [];
    }

    async function scanFrame() {
        const video = document.querySelector('#camera-video'); if (scanBusy || !detector || video.readyState < 2) return;
        scanBusy = true;
        try {
            const found = await detector.detect(video).catch(() => (typeof jsQR === 'function' ? detectWithJsQr(video) : []));
            if (!found.length) { scanBusy = false; return; }
            clearInterval(scanTimer);
            const sn = found[0].rawValue.trim(); const row = records.find(record => String(record.sn).trim() === sn);
            const result = document.querySelector('#scan-result');
            if (!row) {
                result.textContent = `SN ${sn} was not found in the loaded spreadsheet.`;
                result.classList.remove('hidden');
                resumeScanning(1400);
                return;
            }
            const response = await fetch('?action=attendance', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(row)
            });
            const saved = await response.json();
            if (!response.ok) throw new Error(saved.error || 'Could not save this check-in.');
            if (!attendanceRows.some(record => record.sn === saved.record.sn)) attendanceRows.push(saved.record);
            result.textContent = saved.duplicate ? `${row.name} is already checked in.` : `${row.name} checked in successfully.`;
            result.classList.remove('hidden');
            if (!saved.duplicate) await announceWelcome(row);
            if (activeTab === 'attendance') renderAttendance();
            resumeScanning(900);
        } catch (error) {
            const result = document.querySelector('#scan-result');
            if (result) { result.textContent = error.message || 'Could not save this check-in.'; result.classList.remove('hidden'); }
            resumeScanning(1600);
        }
    }

    function resumeScanning(delay) {
        scanTimer = setTimeout(() => {
            scanBusy = false;
            if (stream) scanTimer = setInterval(scanFrame, 250);
        }, delay);
    }

    function announceWelcome(row) {
        const resumeMusic = !audio.paused;
        audio.pause();
        if (!('speechSynthesis' in window) || !('SpeechSynthesisUtterance' in window)) {
            if (resumeMusic) audio.play().catch(() => {});
            return Promise.resolve();
        }
        window.speechSynthesis.cancel();
        const announcement = ['Welcome', row.name, row.designation, row.institution].filter(value => String(value || '').trim()).join(', ');
        const utterance = new SpeechSynthesisUtterance(announcement);
        return new Promise(resolve => {
            let finished = false;
            let timeout;
            const complete = () => {
                if (finished) return;
                finished = true;
                clearTimeout(timeout);
                if (resumeMusic) audio.play().catch(() => {});
                resolve();
            };
            utterance.onend = complete;
            utterance.onerror = complete;
            timeout = setTimeout(complete, 20000);
            window.speechSynthesis.speak(utterance);
        });
    }

    function stopScanner() {
        if (scanTimer) clearInterval(scanTimer); scanTimer = null;
        if (stream) stream.getTracks().forEach(track => track.stop()); stream = null;
        const video = document.querySelector('#camera-video'); if (video) video.srcObject = null;
        document.querySelector('#camera-empty')?.classList.remove('hidden'); document.querySelector('#start-camera')?.classList.remove('hidden'); document.querySelector('#stop-camera')?.classList.add('hidden');
    }

    function exportAttendance() {
        const rows = [['SN', 'Name', 'Designation', 'Institution', 'Whatsapp', 'Check-in time'], ...attendanceRows.map(row => [row.sn, row.name, row.designation, row.institution, row.whatsapp, row.checkedInAt])];
        const csv = rows.map(row => row.map(value => `"${String(value).replace(/"/g, '""')}"`).join(',')).join('\r\n');
        const link = document.createElement('a'); link.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' })); link.download = 'event-attendance.csv'; link.click(); URL.revokeObjectURL(link.href);
    }

    function renderMusic() {
        panel.innerHTML = `<div class="panel-heading"><div><h2>Event music</h2><p class="subheading">Audio files in the application folder play in a random order and continue between tracks.</p></div></div><div class="music-layout"><div class="dropzone"><div class="drop-content"><strong>${tracks.length ? `${tracks.length} music file${tracks.length === 1 ? '' : 's'} found` : 'No music files found'}</strong><p>Place MP3, WAV, OGG, M4A, AAC, or FLAC files beside the app.</p><button class="button primary" id="play-pause" ${tracks.length ? '' : 'disabled'}>${audio.paused ? 'Play music' : 'Pause music'}</button></div></div><div class="playlist"><div class="playlist-head">Application folder Â· ${tracks.length} tracks</div><div id="track-list"></div><div class="music-controls"><input type="range" id="seek" min="0" max="100" value="0" aria-label="Seek audio"><span id="current-time">0:00</span></div></div></div>`;
        document.querySelector('#play-pause').addEventListener('click', () => {
            if (audio.paused && trackIndex < 0) playRandomTrack();
            else if (audio.paused) audio.play().catch(() => toast('The music file could not be played.'));
            else audio.pause();
        });
        document.querySelector('#seek').addEventListener('input', event => { if (audio.duration) audio.currentTime = audio.duration * Number(event.target.value) / 100; });
        renderTracks();
    }

    function renderTracks() {
        const list = document.querySelector('#track-list'); if (!list) return;
        list.innerHTML = tracks.length ? tracks.map((track, index) => `<div class="track"><button data-play="${index}" aria-label="Play ${escapeHtml(track.name)}">â–¶</button><div class="track-info"><strong>${escapeHtml(track.name)}</strong><span>${index === trackIndex ? 'Now playing' : 'In random queue'}</span></div></div>`).join('') : '<div class="empty">No audio files in the application folder.</div>';
        list.querySelectorAll('[data-play]').forEach(button => button.addEventListener('click', () => playTrack(Number(button.dataset.play))));
    }
    function playTrack(index) { if (!tracks[index]) return; trackIndex = index; audio.src = tracks[index].url; audio.play().catch(() => toast('This audio file could not be played.')); renderTracks(); }
    function playRandomTrack() {
        if (!tracks.length) return;
        let nextIndex = Math.floor(Math.random() * tracks.length);
        if (tracks.length > 1 && nextIndex === trackIndex) nextIndex = (nextIndex + 1 + Math.floor(Math.random() * (tracks.length - 1))) % tracks.length;
        playTrack(nextIndex);
    }
    audio.addEventListener('timeupdate', () => { const seek = document.querySelector('#seek'); const time = document.querySelector('#current-time'); if (!seek || !time) return; seek.value = audio.duration ? String(audio.currentTime / audio.duration * 100) : '0'; time.textContent = `${Math.floor(audio.currentTime / 60)}:${String(Math.floor(audio.currentTime % 60)).padStart(2, '0')}`; });
    audio.addEventListener('ended', playRandomTrack);
    audio.addEventListener('play', () => { const button = document.querySelector('#play-pause'); if (button) button.textContent = 'Pause music'; });
    audio.addEventListener('pause', () => { const button = document.querySelector('#play-pause'); if (button) button.textContent = 'Play music'; });

    async function refreshAttendance() {
        try {
            const response = await fetch('?action=attendance');
            const result = await response.json();
            if (!response.ok) throw new Error(result.error || 'Could not load attendance history.');
            attendanceRows = result.records || [];
            if (activeTab === 'attendance') renderAttendance();
        } catch (error) { toast(error.message || 'Could not load attendance history.'); }
    }

    function render() {
        if (activeTab !== 'scan-qr') stopScanner();
        if (activeTab === 'spreadsheet') renderSpreadsheet();
        else if (activeTab === 'send-qr') renderSendQr();
        else if (activeTab === 'scan-qr') renderScan();
        else if (activeTab === 'attendance') renderAttendance();
        else renderMusic();
    }
    document.querySelectorAll('.tab').forEach(tab => tab.addEventListener('click', () => {
        activeTab = tab.dataset.tab;
        document.querySelectorAll('.tab').forEach(item => { const selected = item === tab; item.classList.toggle('active', selected); item.setAttribute('aria-selected', String(selected)); });
        render();
    }));
    try { const saved = JSON.parse(localStorage.getItem(storageKey)); if (saved && Array.isArray(saved.records)) { records = saved.records; sourceName = saved.sourceName || ''; } } catch { localStorage.removeItem(storageKey); }
    render();
    refreshAttendance();
    if (!records.length && preferredSpreadsheet) loadSpreadsheetFromFolder(preferredSpreadsheet);
})();
</script>
</body>
</html>
