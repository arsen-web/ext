<?php

namespace arsenweb\Reader;

use arsenweb\Exceptions\FileException;
use arsenweb\Exceptions\ReaderException;
use arsenweb\Helpers\FileHelper;
use arsenweb\Helpers\IFileHelper;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Reader\Csv;
use PhpOffice\PhpSpreadsheet\Reader\Xls;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class ExtReader implements IExtReader
{
    /** @var self */
    protected static $instance;

    /** @var Spreadsheet */
    protected $spreadsheet;

    /** @var array */
    protected $map;

    /** @var array */
    protected $coordinates;

    /** @var IFileHelper */
    protected $fileHelper;

    protected function __construct()
    {
        $this->fileHelper = new FileHelper();
    }

    /**
     * @return ExtReader
     */
    protected static function getInstance(): self
    {
        if(is_null(static::$instance)) {
            static::$instance = new static();
        }

        return static::$instance;
    }

    /**
     * @param string $pathToTemplate
     * @return $this
     * @throws FileException
     * @throws ReaderException
     */
    public static function read(string $pathToTemplate): IExtReader
    {
        $instance = static::getInstance();

        $instance->fileHelper->validateFilePath($pathToTemplate);

        switch($instance->fileHelper->getTypeFile($pathToTemplate)) {
            case FileHelper::TYPE_XLSX:
                $instance->readXlsx($pathToTemplate);
                break;
            case FileHelper::TYPE_XLS:
                $instance->readXls($pathToTemplate);
                break;
            case FileHelper::TYPE_CSV:
                $instance->readCsv($pathToTemplate);
                break;
            default:
                throw new ReaderException("The file `{$pathToTemplate}` is not supported");
        }

        return $instance;
    }

    /**
     * @param string $pathToTemplate
     * @return self
     * @throws FileException
     */
    public static function readXlsx(string $pathToTemplate): IExtReader
    {
        $instance = static::getInstance();
        $instance->fileHelper->validateFilePath($pathToTemplate);
        $instance->fileHelper->validateFileType($pathToTemplate, FileHelper::TYPE_XLSX);

        $xlsxReader = new Xlsx();
        $instance->spreadsheet = $xlsxReader->load($pathToTemplate);

        return $instance;
    }

    /**
     * @param string $pathToTemplate
     * @return self
     * @throws FileException
     */
    public static function readXls(string $pathToTemplate): IExtReader
    {
        $instance = static::getInstance();
        $instance->fileHelper->validateFilePath($pathToTemplate);
        $instance->fileHelper->validateFileType($pathToTemplate, FileHelper::TYPE_XLS);

        $xlsxReader = new Xls();
        $instance->spreadsheet = $xlsxReader->load($pathToTemplate);

        return $instance;
    }

    /**
     * @param string $pathToTemplate
     * @return self
     * @throws FileException
     */
    public static function readCsv(string $pathToTemplate): IExtReader
    {
        $instance = static::getInstance();
        $instance->fileHelper->validateFilePath($pathToTemplate);
        $instance->fileHelper->validateFileType($pathToTemplate, FileHelper::TYPE_CSV);

        $xlsxReader = new Csv();
        $instance->spreadsheet = $xlsxReader->load($pathToTemplate);

        return $instance;
    }

    /**
     * @param string $tempFile
     * @param string $pathToSave
     * @return IExtReader
     * @throws FileException
     * @throws ReaderException
     */
    public static function readTempFile(string $tempFile, string $pathToSave): IExtReader
    {
        $instance = static::getInstance();

        move_uploaded_file($tempFile, $pathToSave);

        return $instance->read($pathToSave);
    }

    /**
     * @return Spreadsheet
     */
    public function getSpreadsheet(): Spreadsheet
    {
        return $this->spreadsheet;
    }

    /**
     * @param array $map
     * @return ExtReader
     * @throws Exception
     */
    public function setMaskList(array $map): IExtReader
    {
        $this->map = $map;
        $this->setCoordinates();

        return $this;
    }

    /**
     * @return array
     */
    public function getCoordinates(): array
    {
        return $this->coordinates;
    }

    /**
     * @throws Exception
     */
    protected function setCoordinates(): void
    {
        foreach($this->spreadsheet->getWorksheetIterator() as $worksheet) {

            foreach($worksheet->getRowIterator() as $row) {

                $cellIterator = $row->getCellIterator();
                $cellIterator->setIterateOnlyExistingCells(true);

                foreach($cellIterator as $cell) {
                    $mask = trim($cell->getValue());

                    if(in_array($mask, $this->map)) {
                        [$col, $row] = Coordinate::coordinateFromString($cell->getCoordinate());

                        $this->coordinates[$mask] = [
                            'col' => $col,
                            'row' => $row,
                        ];
                    }

                }

            }

        }
    }
}
