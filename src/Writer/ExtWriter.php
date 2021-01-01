<?php

namespace arsenweb\Writer;

use arsenweb\Exceptions\ColumnException;
use arsenweb\Exceptions\WriterException;
use arsenweb\Helpers\ColumnHelper;
use arsenweb\Helpers\FileHelper;
use arsenweb\Helpers\IFileHelper;
use arsenweb\Reader\ExtReader;
use arsenweb\Reader\IExtReader;
use PhpOffice\PhpSpreadsheet\Writer\Csv;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Throwable;

class ExtWriter implements IExtWriter
{
    /** @var string */
    public const DIRECTION_DOWN = 'down';

    /** @var string */
    public const DIRECTION_RIGHT = 'right';

    /** @var self */
    protected static $instance;

    /** @var ExtReader */
    protected $reader;

    /** @var IFileHelper */
    protected $fileHelper;

    /** @var ColumnHelper */
    protected $columnHelper;

    protected function __construct()
    {
        $this->fileHelper = new FileHelper();
        $this->columnHelper = new ColumnHelper();
    }

    /**
     * @return ExtWriter
     */
    protected static function getInstance(): self
    {
        if(is_null(static::$instance)) {
            static::$instance = new static();
        }

        return static::$instance;
    }

    /**
     * @param IExtReader $reader
     * @return IExtWriter
     */
    public static function setReader(IExtReader $reader): IExtWriter
    {
        $instance = static::getInstance();

        $instance->reader = $reader;

        return $instance;
    }

    /**
     * @param string $mask
     * @param string $value
     * @return IExtWriter
     * @throws ColumnException
     * @throws WriterException
     */
    public function setString(string $mask, string $value): IExtWriter
    {
        ['col' => $colNumber, 'row' => $rowNumber] = $this->getCoordinatesByMask($mask);

        $this->reader
            ->getSpreadsheet()
            ->getActiveSheet()
            ->setCellValueByColumnAndRow($colNumber, $rowNumber, (string)$value);

        return $this;
    }

    /**
     * @param string $mask
     * @param int $value
     * @return $this
     * @throws ColumnException
     * @throws WriterException
     */
    public function setInt(string $mask, int $value): IExtWriter
    {
        ['col' => $colNumber, 'row' => $rowNumber] = $this->getCoordinatesByMask($mask);

        $this->reader
            ->getSpreadsheet()
            ->getActiveSheet()
            ->setCellValueByColumnAndRow($colNumber, $rowNumber, (int)$value);

        return $this;
    }

    /**
     * @param string $mask
     * @param $value
     * @param int $precision
     * @return $this
     * @throws ColumnException
     * @throws WriterException
     */
    public function setFloat(string $mask, $value, int $precision = 2): IExtWriter
    {
        ['col' => $colNumber, 'row' => $rowNumber] = $this->getCoordinatesByMask($mask);

        $value = str_replace(',', '.', $value);

        $this->reader
            ->getSpreadsheet()
            ->getActiveSheet()
            ->setCellValueByColumnAndRow($colNumber, $rowNumber, round((float)$value, $precision));

        return $this;
    }

    /**
     * @param string $mask
     * @param array $items
     * @param callable|null $fn
     * @param string $direction
     * @return IExtWriter
     * @throws ColumnException
     * @throws WriterException
     */
    public function setArray(
        string $mask,
        array $items,
        ?callable $fn = null,
        string $direction = self::DIRECTION_DOWN
    ): IExtWriter
    {
        ['col' => $colNumber, 'row' => $rowNumber] = $this->getCoordinatesByMask($mask);

        $items = !is_null($fn) ? $fn($items) : $items;

        foreach($items as $item) {
            if(!is_array($item)) {
                $this->reader
                    ->getSpreadsheet()
                    ->getActiveSheet()
                    ->setCellValueByColumnAndRow($colNumber, $rowNumber, $item);

                switch($direction) {
                    case self::DIRECTION_DOWN:
                        $rowNumber += 1;
                        break;
                    case self::DIRECTION_RIGHT:
                        $colNumber += 1;
                        break;
                }
            }
        }

        return $this;
    }

    /**
     * @param string $mask
     * @param array $rows
     * @param callable $fn
     * @return IExtWriter
     * @throws ColumnException
     * @throws WriterException
     */
    public function setAssociativeArray(string $mask, array $rows, ?callable $fn = null): IExtWriter
    {
        ['col' => $colNumber, 'row' => $rowNumber] = $this->getCoordinatesByMask($mask);

        $rows = !is_null($fn) ? $fn($rows) : $rows;

        foreach($rows as $row) {
            if(!is_array($row)) {
                throw new WriterException("`{$row}` is not an array");
            }

            $colNumberTemp = $colNumber;
            foreach($row as $col) {
                $this->reader
                    ->getSpreadsheet()
                    ->getActiveSheet()
                    ->setCellValueByColumnAndRow($colNumberTemp, $rowNumber, $col);
                $colNumberTemp += 1;
            }
            $rowNumber += 1;
        }

        return $this;
    }

    /**
     * @param string $pathToFile
     * @param string $fileType
     * @return bool
     */
    public function save(string $pathToFile, ?string $fileType = null): bool
    {
        if(is_null($fileType)) {
            $fileType = $this->fileHelper->getTypeFile($pathToFile);
        }

        try {
            switch($fileType) {
                case FileHelper::TYPE_XLSX:
                    $writer = new Xlsx($this->reader->getSpreadsheet());
                    break;
                case FileHelper::TYPE_XLS:
                    $writer = new Xls($this->reader->getSpreadsheet());
                    break;
                case FileHelper::TYPE_CSV:
                    $writer = new Csv($this->reader->getSpreadsheet());
                    break;
                default:
                    throw new WriterException("Invalid type `{$fileType}`");
            }

            $pathToFile = $this->fileHelper->dropTypeFile($pathToFile);

            $writer->save("{$pathToFile}{$fileType}");
        } catch(Throwable $e) {
            return false;
        }

        return true;
    }

    /**
     * @param string $fileType
     * @return string
     * @throws WriterException
     */
    public function getContent(string $fileType): string
    {
        if(!in_array($fileType, FileHelper::TYPE_LIST)) {
            throw new WriterException("Invalid type `{$fileType}`");
        }

        $tempFile = sys_get_temp_dir() . '/' . uniqid('ext_') . $fileType;
        fopen($tempFile, 'x');

        $this->save($fileType, $tempFile);

        $handle = fopen($tempFile, 'r');
        $content = fread($handle, filesize($tempFile));
        unlink($tempFile);

        return $content;
    }

    /**
     * @param string $fileType
     * @return string
     * @throws WriterException
     */
    public function getBase64(string $fileType): string
    {
        return base64_encode($this->getContent($fileType));
    }

    /**
     * @param string $mask
     * @return array
     * @throws ColumnException
     * @throws WriterException
     */
    protected function getCoordinatesByMask(string $mask): array
    {
        $coordinates = $this->reader->getCoordinates();

        if(!isset($coordinates[$mask])) {
            throw new WriterException("Invalid mask `{$mask}`");
        }

        return [
            'col' => $this->columnHelper->getColIndex($coordinates[$mask]['col']),
            'row' => $coordinates[$mask]['row'],
        ];
    }
}
