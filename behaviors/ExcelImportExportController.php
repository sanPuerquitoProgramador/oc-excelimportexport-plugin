<?php namespace WRvE\ExcelImportExport\Behaviors;

use ApplicationException;
use Backend\Behaviors\ImportExportController;
use League\Csv\Reader as CsvReader;
use October\Rain\Database\Models\DeferredBinding;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Exception;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Csv;
use System\Models\File;

class ExcelImportExportController extends ImportExportController
{
    protected $columnsToKeep = "all";

    public function __construct($controller)
    {
        parent::__construct($controller);
        $this->viewPath = base_path() . '/modules/backend/behaviors/importexportcontroller/partials';
        $this->assetPath = '/modules/backend/behaviors/importexportcontroller/assets';
    }

    public function setColumnsToKeep($columns)
    {
        $this->columnsToKeep = $columns;
    }

    protected function createCsvReader(string $path): CsvReader
    {
        $path = $this->convertToCsv($path);
        return parent::createCsvReader($path);
    }

    /**
     * @throws ApplicationException
     */
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

        // Determinar las columnas a mantener
        if ($this->columnsToKeep === "all" || empty($this->columnsToKeep)) {
            $columnsToKeep = [];
            $firstRow = $sheet->getRowIterator()->current();
            foreach ($firstRow->getCellIterator() as $cell) {
                if (!is_null($cell->getValue()) && $cell->getValue() !== '') {
                    $columnsToKeep[] = $cell->getColumn();
                }
            }
        } else {
            $columnsToKeep = $this->columnsToKeep;
        }

        $currentColumn = 1;

        // Crear una nueva hoja de cálculo para almacenar solo las columnas seleccionadas
        $newSpreadsheet = new Spreadsheet();
        $newSheet = $newSpreadsheet->getActiveSheet();

        // Copiar los valores de las columnas seleccionadas a la nueva hoja de cálculo
        foreach ($columnsToKeep as $column) {
            $colIndex = Coordinate::columnIndexFromString($column);
            foreach ($sheet->getRowIterator() as $row) {
                $rowIndex = $row->getRowIndex();
                $cellCoordinate = Coordinate::stringFromColumnIndex($colIndex) . $rowIndex;
                $cell = $sheet->getCell($cellCoordinate); // Método no deprecado
                $value = $cell->getValue();

                // // Verificar si la celda tiene un valor antes de formatear
                if (!is_null($value) && $value !== '') {
                //     // Verificar si la celda es una fecha utilizando el formato de celda
                //     $cellFormat = $sheet->getStyle($cell->getCoordinate())->getNumberFormat()->getFormatCode();
                //     if (Date::isDateTimeFormatCode($cellFormat)) {
                //         $value = Date::excelToDateTimeObject($value)->format('d/m/Y');
                    // }
                } else {
                    // Si la celda está vacía, establecer el valor a null
                    $value = null;
                }

                $newCellCoordinate = Coordinate::stringFromColumnIndex($currentColumn) . $rowIndex;
                $newSheet->getCell($newCellCoordinate)->setValue($value); // Método no deprecado
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
