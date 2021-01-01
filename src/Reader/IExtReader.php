<?php

namespace arsenweb\Reader;

use PhpOffice\PhpSpreadsheet\Spreadsheet;

interface IExtReader
{
    public static function read(string $pathToTemplate): IExtReader;

    public static function readXlsx(string $pathToTemplate): IExtReader;

    public static function readXls(string $pathToTemplate): IExtReader;

    public static function readCsv(string $pathToTemplate): IExtReader;

    public static function readTempFile(string $tempFile, string $pathToSave): IExtReader;

    public function getSpreadsheet(): Spreadsheet;

    public function setMaskList(array $map): IExtReader;

    public function getCoordinates(): array;
}
