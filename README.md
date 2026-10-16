# Event Attendance Desk

Single-page PHP application for spreadsheet-based QR passes, check-in, spoken welcomes, and event music.

## Setup

- Use PHP 8.0 or newer with cURL, `proc_open`, and the system `zip` command available.
- Put the attendee `.xlsx`, `.xls`, or `.csv` spreadsheet in the same folder as `index.php`, or choose it from the Spreadsheet tab.
- `invitations.xlsx` is selected and loaded automatically when present. Scan QR is the first tab and opens by default.
- The first row must contain `SN`, `Name`, `Designation`, `Institution`, and `Whatsapp` headers. Header matching is case-insensitive.
- Put supported audio files (`mp3`, `wav`, `ogg`, `m4a`, `aac`, or `flac`) beside `index.php`. Start playback from the Music tab; the browser requires a user gesture before it plays audio.
- Ensure the PHP process can write to the app folder. It creates `QR/`, `data/attendance.json`, and `Attendance.xlsx` there as guests are processed.
- Serve the app over HTTPS for camera access on other devices. QR scanning uses the browser `BarcodeDetector` API, available in current Chromium-based browsers.

## WhatsApp Delivery

Enter the API username and password in the Spreadsheet tab. The password is saved in the PHP server session, not browser storage. Alternatively, configure credentials through the server environment. Never commit the API password.

```sh
export APP_PUBLIC_URL='https://your-domain.example/eventattendance'
# Optional instead of entering credentials in the app:
export BHSMS_USER='your-bhashsms-username'
export BHSMS_PASSWORD='your-bhashsms-password'
```

`BHSMS_USER` defaults to `InnocentWTS`, `BHSMS_SENDER` defaults to `BUZWAP`, and `BHSMS_API_URL` defaults to the HTTPS BhashSMS endpoint. Override them through the server environment if your account uses different values. The configured public URL must allow BhashSMS to fetch `QR/<SN>.jpg`. The request uses the `conclavereminder` template and sends Name, Designation, and Institution as template parameters; ensure the provider template matches those fields.

Attendance scans are kept in `data/attendance.json` and written to `Attendance.xlsx`; repeat scans of the same SN do not add duplicate rows. Speech announcements use the browser's built-in speech synthesis and pause/resume currently playing music.