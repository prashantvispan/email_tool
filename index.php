<?php
require 'vendor/autoload.php'; // Path to your Composer autoload file

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

function extractEmailsFromSheet($filePath)
{
    $emails = [];

    // Load the spreadsheet
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($filePath);

    // Get the first sheet in the workbook
    $sheet = $spreadsheet->getActiveSheet();

    // Iterate through all rows with data
    foreach ($sheet->getRowIterator() as $row) {
        // Iterate through all columns in the row
        foreach ($row->getCellIterator() as $cell) {
            $cellValue = $cell->getValue();
            if (filter_var($cellValue, FILTER_VALIDATE_EMAIL)) {
                $emails[] = $cellValue;
            }
        }
    }

    return $emails;
}

function getMXRecords($email)
{
    $domain = substr(strrchr($email, "@"), 1); // Extract domain from email
    $mxRecords = @dns_get_record($domain, DNS_MX);

    // Handle DNS query failure
    if ($mxRecords === false) {
        // Log the error if needed
        error_log("Failed to retrieve MX records for domain: $domain");
        return []; // Return an empty array or handle the error condition as required
    }

    return $mxRecords;
}

?>

<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <title>Upload Excel File and Group Emails by Provider</title>
    <link rel="stylesheet" href="https://bootswatch.com/5/cosmo/bootstrap.min.css">
</head>

<body class="container card p-4 col-md-6">
    <h2>Upload Excel File to Group Emails by Provider</h2>
    <hr>
    <form action="index.php" method="post" enctype="multipart/form-data">
        <div class="form-group">
            <input type="file" name="file" accept=".xlsx, .xls, .csv" class="form-control" required>
            <button type="submit" class="btn btn-primary mt-4">Upload and Group Emails</button>
        </div>
    </form>

    <?php
    // Handle file upload and email extraction
    if ($_SERVER["REQUEST_METHOD"] == "POST" && isset($_FILES["file"])) {
        $uploadOk = true;

        // Check file size (optional)
        if ($_FILES['file']['size'] > 5000000) { // Adjust the file size limit as needed
            echo "File is too large.";
            $uploadOk = false;
        }

        // Allow only specific file formats (optional)
        $fileType = pathinfo($_FILES['file']['name'], PATHINFO_EXTENSION);
        if ($fileType != "xlsx" && $fileType != "xls" && $fileType != "csv") {
            echo "Only Excel files are allowed.";
            $uploadOk = false;
        }

        // Process file if upload is valid
        if ($uploadOk && move_uploaded_file($_FILES['file']['tmp_name'], 'temp_uploaded_file.xlsx')) {
            // Extract emails from uploaded file
            $emails = extractEmailsFromSheet('temp_uploaded_file.xlsx');

            // Group emails by provider
            $groupedEmails = [];

            foreach ($emails as $email) {
                $mxRecords = getMXRecords($email);
                $domain = substr(strrchr($email, "@"), 1); // Extract domain from email

                // Check if the domain exists in groupedEmails, if not add it
                if (!isset($groupedEmails[$domain])) {
                    $groupedEmails[$domain] = [
                        'emails' => [],
                        'domains' => []
                    ];
                }

                // Check if the MX records for the domain match Google
                $isGoogleDomain = false;
                foreach ($mxRecords as $mxRecord) {
                    if (strpos($mxRecord['target'], 'google.com') !== false) {
                        $isGoogleDomain = true;
                        break;
                    }
                }

                // Add email to the appropriate group
                if ($isGoogleDomain) {
                    $groupedEmails['google.com']['emails'][] = $email;
                    $groupedEmails['google.com']['domains'][] = $domain;
                } else {
                    $groupedEmails[$domain]['emails'][] = $email;
                    $groupedEmails[$domain]['domains'][] = $domain;
                }
            }

            // Display total count of emails
            $totalCount = count($emails);
            echo '<hr /><h3>Total Count of Emails: ' . $totalCount . '</h3>';

            // Accumulate download links for each group
            $downloadLinks = [];
            foreach ($groupedEmails as $provider => $group) {
                if (empty($group['emails'])) {
                    continue; // Skip if no emails for this provider
                }

                // Generate Excel content in memory
                $outputSpreadsheet = new Spreadsheet();
                $outputSheet = $outputSpreadsheet->getActiveSheet();
                $rowIndex = 1;

                foreach ($group['emails'] as $email) {
                    $outputSheet->setCellValue('A' . $rowIndex, $email);
                    $rowIndex++;
                }

                // Save Excel content to a variable
                $writer = new Xlsx($outputSpreadsheet);
                ob_start();
                $writer->save('php://output');
                $excelContent = ob_get_clean();

                // Generate download link and store in array
                $downloadLinks[] = [
                    'provider' => $provider,
                    'link' => 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' . base64_encode($excelContent),
                    'filename' => strtolower($provider) . '_emails.xlsx'
                ];
            }

            // Display download links
            foreach ($downloadLinks as $link) {
                echo '<p><a href="' . $link['link'] . '" class="btn btn-info">Download ' . $link['provider'] . ' Emails</a></p>';
            }

            // Delete the temporary uploaded file
            unlink('temp_uploaded_file.xlsx');
        } else {
            echo "Failed to upload or process file.";
        }
    }
    ?>
</body>

</html>
