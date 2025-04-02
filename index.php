<?php

// Fehleranzeige deaktivieren
error_reporting(0);
ini_set('display_errors', 0);

require_once 'vendor/autoload.php';

use PhpOffice\PhpWord\TemplateProcessor;

// Standardwerte für die Stunden
$eingaben_stunden = [
    '1. Stunde' => '',
    '2. Stunde' => '',
    '3. Stunde' => '',
    '4. Stunde' => '',
    '5. Stunde' => '',
    '6. Stunde' => ''
];

// Wochentag auswählen oder Standard setzen
$auswahl_wochentag = isset($_POST['wochentag']) ? $_POST['wochentag'] : '';

// Zustandsvariablen
$uebertragen_erfolgreich = false;
$htmlOutput = "";
$fehler_meldung = "";

// Berechnung der aktuellen Kalenderwoche
$heute = new DateTime();
$kw = $heute->format('W');

$montag = clone $heute;
$montag->modify('monday this week');
$datumv = $montag->format('d.m.Y');

$freitag = clone $heute;
$freitag->modify('friday this week');
$datumb = $freitag->format('d.m.Y');

// Wenn Formular abgeschickt wurde
if ($_SERVER['REQUEST_METHOD'] == 'POST' && isset($_POST['speichern'])) {
    $auswahl_wochentag = $_POST['aktueller_wochentag'];
    if (isset($_POST['eingaben_stunden']) && is_array($_POST['eingaben_stunden'])) {
        $eingaben_stunden = $_POST['eingaben_stunden'];
    }

    // Datei-Upload prüfen
    if (isset($_FILES['wordfile']) && $_FILES['wordfile']['error'] == 0) {
        $fileTmpPath = $_FILES['wordfile']['tmp_name'];
        $fileName = $_FILES['wordfile']['name'];
        $fileExt = pathinfo($fileName, PATHINFO_EXTENSION);

        if ($fileExt === 'docx') {
            try {
                $templateProcessor = new TemplateProcessor($fileTmpPath);

                $wochentage_map = [
                    'Montag' => 'mo',
                    'Dienstag' => 'di',
                    'Mittwoch' => 'mi',
                    'Donnerstag' => 'do',
                    'Freitag' => 'fr'
                ];

                // Platzhalter ersetzen
                $templateProcessor->setValue('{{kw}}', $kw);
                $templateProcessor->setValue('{{datumv}}', $datumv);
                $templateProcessor->setValue('{{datumb}}', $datumb);

                if (array_key_exists($auswahl_wochentag, $wochentage_map)) {
                    $prefix = $wochentage_map[$auswahl_wochentag];
                    foreach ($eingaben_stunden as $stunde => $eingabe) {
                        $stunden_nummer = explode('.', $stunde)[0];
                        $platzhalter = "{{" . $prefix . $stunden_nummer . "}}";
                        if (!empty($eingabe)) {
                            $templateProcessor->setValue($platzhalter, $eingabe);
                        }
                    }
                }

                // Temporäre Datei erzeugen und zum Download anbieten
                $tempFile = tempnam(sys_get_temp_dir(), 'doc_') . '.docx';
                $templateProcessor->saveAs($tempFile);

                header("Content-Description: File Transfer");
                header("Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document");
                $kwFilename = "KW" . $kw . ".docx";
                header("Content-Disposition: attachment; filename=\"$kwFilename\"");

                header("Content-Transfer-Encoding: binary");
                header("Expires: 0");
                header("Cache-Control: must-revalidate");
                header("Pragma: public");
                header("Content-Length: " . filesize($tempFile));
                flush();
                readfile($tempFile);
                unlink($tempFile);
                exit;

            } catch (Exception $e) {
                $fehler_meldung = "Fehler beim Verarbeiten der Datei: " . $e->getMessage();
            }
        } else {
            $fehler_meldung = "Bitte lade eine gültige .docx-Datei hoch.";
        }
    }

    // Reset für neue Eingaben
    $eingaben_stunden = array_fill_keys(array_keys($eingaben_stunden), '');
    $auswahl_wochentag = '';
}
?>

<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Wochennachweis</title>
    <link rel="stylesheet" href="styles.css">
</head>
<body>
    <div class="container">
        <form method="POST" action="" enctype="multipart/form-data">
            <h1>Wochennachweis</h1>

            <div class="wochentage">
                <?php
                $wochentage = ['Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag'];
                foreach ($wochentage as $tag) {
                    $buttonClass = ($tag == $auswahl_wochentag) ? 'btn selected' : 'btn';
                    echo '<button type="submit" name="wochentag" value="' . $tag . '" class="' . $buttonClass . '">' . $tag . '</button>';
                }
                ?>
            </div>

            <div class="stunden-container">
                <?php foreach ($eingaben_stunden as $stunde => $eingabe): ?>
                    <div class="stunde">
                        <label for="<?php echo $stunde; ?>"><?php echo $stunde; ?>:</label>
                        <input 
                            type="text" 
                            name="eingaben_stunden[<?php echo $stunde; ?>]" 
                            id="<?php echo $stunde; ?>" 
                            value="<?php echo htmlspecialchars($eingabe); ?>" 
                            class="text-input" 
                            placeholder="Gib deinen Text ein"
                            <?php echo ($auswahl_wochentag ? '' : 'disabled'); ?>>
                    </div>
                <?php endforeach; ?>
            </div>

            <input type="hidden" name="aktueller_wochentag" value="<?php echo htmlspecialchars($auswahl_wochentag); ?>">

            <div class="upload-container">
                <input type="file" name="wordfile" accept=".docx" id="file-upload" class="upload-input">
                <label for="file-upload">Durchsuchen...</label>
            </div>

            <div class="button-container">
                <button type="submit" name="speichern" class="btn">Übertragen</button>
            </div>
        </form>
    </div>

    <?php if (!empty($fehler_meldung)): ?>
        <script>
            alert("Fehler: <?php echo addslashes($fehler_meldung); ?>");
        </script>
    <?php endif; ?>
</body>
</html>
