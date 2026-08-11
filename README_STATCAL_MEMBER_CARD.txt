STATCAL MEMBER ID CARD - UPLOAD GUIDE

Recommended structure:

public_html/
├── index.html
├── member_lookup.php
├── logo.png
├── style.css
├── tombol_css.css
├── script.js
├── map_statcal.html
├── data_long_lat.xlsx
└── private_data/
    ├── .htaccess
    └── data_member_statcal.xlsx

INSTALL:
1. Back up your current index.html.
2. Rename index_member_card.html to index.html and upload it.
3. Upload member_lookup.php beside index.html.
4. Create a folder named private_data.
5. Upload .htaccess and data_member_statcal.xlsx into private_data.
6. Test a valid ID such as 20260001.

HOW IT WORKS:
- The visitor enters a Member ID.
- index.html calls member_lookup.php.
- member_lookup.php reads the XLSX on the SERVER.
- Only the matching member record is returned to the browser.
- The ID Card is displayed.
- Download ID Card PDF creates an ID-1 proportion PDF.

The PHP endpoint does NOT require Composer or PhpSpreadsheet.
It uses ZipArchive when available and includes a pure-PHP ZIP/XLSX fallback.
The fallback requires standard PHP zlib support.

STRONGER PRIVACY:
For stronger isolation, place private_data outside public_html.
Then change in member_lookup.php:

$excelPath = __DIR__ . '/private_data/data_member_statcal.xlsx';

to something like:

$excelPath = dirname(__DIR__) . '/private_data/data_member_statcal.xlsx';

SECURITY NOTE:
A Member ID alone is verification, not strong authentication.
Because Member IDs may be predictable, consider adding a second private
verification value in a future version.
