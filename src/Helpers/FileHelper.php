<?php

namespace arsenweb\Helpers;

use arsenweb\Exceptions\FileException;

class FileHelper implements IFileHelper
{
    /** @var string */
    public const TYPE_XLSX = 'xlsx';

    /** @var string */
    public const TYPE_XLS = 'xls';

    /** @var string */
    public const TYPE_CSV = 'csv';

    public const TYPE_LIST = [
        self::TYPE_XLSX,
        self::TYPE_XLS,
        self::TYPE_CSV,
    ];

    /**
     * Получить тип файла
     *
     * @param string $fileName
     * @return string
     */
    public function getTypeFile(string $fileName): string
    {
        return substr($fileName, strrpos($fileName, '.') + 1);
    }

    /**
     * Удалить тип файла
     *
     * @param string $fileName
     * @return string
     */
    public function dropTypeFile(string $fileName): string
    {
        return str_replace($this->getTypeFile($fileName), '', $fileName);
    }

    /**
     * Валидация файла
     *
     * @param string $pathToFile
     * @throws FileException
     */
    public function validateFilePath(string $pathToFile): void
    {
        if(empty($pathToFile)) {
            throw new FileException('File not listed.');
        }

        if(!file_exists($pathToFile)) {
            throw new FileException("The file `{$pathToFile}` does not exist.");
        }

        if(!is_readable($pathToFile)) {
            throw new FileException("Could not open file `{$pathToFile}` for reading.");
        }

        if(!is_file($pathToFile)) {
            throw new FileException("`{$pathToFile}` is not a file.");
        }
    }

    /**
     * @param string $pathToFile
     * @param string $type
     * @throws FileException
     */
    public function validateFileType(string $pathToFile, string $type): void
    {
        if(!preg_match("/\.{$type}$/", $pathToFile)) {
            throw new FileException("The file` {$pathToFile} `does not match the type. {$type}");
        }
    }
}
