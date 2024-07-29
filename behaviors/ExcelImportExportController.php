<?php namespace WRvE\ExcelImportExport\Behaviors;

use Backend\Behaviors\ImportExportController;
use October\Rain\Database\Models\DeferredBinding;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Exception;
use PhpOffice\PhpSpreadsheet\Writer\Csv;
use System\Models\File;
use ApplicationException;
use League\Csv\Reader as CsvReader;

class ExcelImportExportController extends ImportExportController
{
    public function __construct($controller)
    {
        parent::__construct($controller);
        $this->viewPath = base_path() . '/modules/backend/behaviors/importexportcontroller/partials';
        $this->assetPath = '/modules/backend/behaviors/importexportcontroller/assets';
    }

    protected function createCsvReader(string $path): CsvReader
    {
        $path = $this->convertToCsv($path);

        return parent::createCsvReader($path);
    }

    private function convertToCsv(string $path)
    {
        $ext = pathinfo($path, PATHINFO_EXTENSION);

        if ($ext === 'csv' || mime_content_type($path) === 'text/csv' || mime_content_type($path) === 'text/plain') {
            return $path;
        }

        $tempCsvPath = $path . '.csv';
        $inputFileType = IOFactory::identify($path);

        try {
            $reader = IOFactory::createReader($inputFileType);
        } catch (Exception $e) {
            throw new ApplicationException('Unsupported file type: ' . $inputFileType);
        }

        $spreadsheet = $reader->load($path);
        $sheet = $spreadsheet->getActiveSheet();

        // Crear una nueva hoja de cálculo para almacenar solo las columnas seleccionadas
        $newSpreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        $newSheet = $newSpreadsheet->getActiveSheet();

        // Definir las columnas a mantener. Util cuando se tienen celdas basura
        $columnsToKeep = ['A', 'B', 'C', 'D', 'E', 'I', 'J', 'L', 'M', 'N', 'P', 'Q'];
        $currentColumn = 1;

        // Copiar los valores de las columnas seleccionadas a la nueva hoja de cálculo
        foreach ($columnsToKeep as $column) {
            $colIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($column);
            foreach ($sheet->getRowIterator() as $row) {
                $cell = $sheet->getCellByColumnAndRow($colIndex, $row->getRowIndex());
                $value = $cell->getValue();

                // Verificar si la celda es una fecha utilizando el formato de celda
                $cellFormat = $sheet->getStyle($cell->getCoordinate())->getNumberFormat()->getFormatCode();
                if (\PhpOffice\PhpSpreadsheet\Shared\Date::isDateTimeFormatCode($cellFormat)) {
                    $value = \PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($value)->format('d/m/Y');
                }

                $newSheet->setCellValueByColumnAndRow($currentColumn, $row->getRowIndex(), $value);
            }
            $currentColumn++;
        }

        $writer = new Csv($newSpreadsheet);
        $writer->setSheetIndex(0);
        $writer->save($tempCsvPath);

        $fileModel = $this->getFileModel();
        $disk = $fileModel->getDisk();
        $disk->put($fileModel->getDiskPath() . '.csv', file_get_contents($tempCsvPath));
        $fileModel->disk_name = $fileModel->disk_name . '.csv';
        $fileModel->save();

        return $path . '.csv';
    }


    /**
     * @return File
     */
    private function getFileModel()
    {
        $sessionKey = $this->importUploadFormWidget->getSessionKey();

        $deferredBinding = DeferredBinding::where('session_key', $sessionKey)
            ->orderBy('id', 'desc')
            ->where('master_field', 'import_file')
            ->first();

        return $deferredBinding->slave_type::find($deferredBinding->slave_id);
    }
}
